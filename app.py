"""Montador Visual de Kits – Momento Uau.

Organização deste arquivo:
1. Configuração e constantes
2. Utilidades (formatação, imagens, cálculos, exportação)
3. Estado e ações (callbacks)
4. Estilos (CSS)
5. Componentes de interface
6. Páginas (Montagem, Catálogo, Histórico)
7. Função principal
"""

from __future__ import annotations

import html
import inspect
import io
import re
from dataclasses import dataclass, field
from pathlib import Path
from typing import Any, Callable, Iterable
from uuid import uuid4

import pandas as pd
import streamlit as st
from openpyxl.utils import get_column_letter
from PIL import Image

from database import (
    ASSETS_DIR,
    add_catalog_item,
    delete_catalog_item,
    delete_kit,
    delete_recipe,
    format_address_summary,
    get_catalog_items,
    get_kits_history,
    get_recipes,
    init_db,
    save_kit,
    save_recipe,
    update_catalog_item,
    update_recipe,
)

# =============================================================================
# 1. CONFIGURAÇÃO E CONSTANTES
# =============================================================================

APP_TITLE = "Montador Visual de Kits"
APP_SUBTITLE = "Momento Uau · seleção visual de brindes e cálculo automático"

STATUS_OPTIONS = ["Em montagem", "Conferido", "Finalizado"]
STATUS_CHIP = {"Em montagem": "chip-warn", "Conferido": "chip-info", "Finalizado": "chip-ok"}
MAX_BATCH_KITS = 10
ALLOWED_IMAGE_TYPES = ["png", "jpg", "jpeg", "webp"]

PAGES = {
    "montagem": "Montagem do kit",
    "catalogo": "Catálogo e brindes",
    "historico": "Histórico",
}

FORM_FIELDS = [
    "kit_code", "client_name", "status", "notes",
    "cep", "street", "address_number", "complement", "neighborhood", "city_state",
]
ADDRESS_FIELDS = ["cep", "street", "address_number", "complement", "neighborhood", "city_state"]

STRETCH: dict[str, Any] = (
    {"width": "stretch"}
    if "width" in inspect.signature(st.button).parameters
    else {"use_container_width": True}
)

Item = dict[str, Any]


@dataclass
class KitTotals:
    lines: list[dict[str, Any]] = field(default_factory=list)
    quantity: int = 0
    value: float = 0.0


# =============================================================================
# 2. UTILIDADES
# =============================================================================


def format_brl(value: float) -> str:
    """1234.5 -> 'R$ 1.234,50'."""
    formatted = f"{float(value):,.2f}"
    formatted = formatted.replace(",", "X").replace(".", ",").replace("X", ".")
    return f"R$ {formatted}"


def esc(value: Any) -> str:
    """Escapa texto do usuário antes de inseri-lo em HTML."""
    return html.escape(str(value or ""))


def chunked(seq: list[Any], size: int) -> Iterable[list[Any]]:
    for start in range(0, len(seq), size):
        yield seq[start : start + size]


def icon_text(item: Item) -> str:
    raw = str(item.get("icon") or "").strip().strip("[]").strip()
    if not raw or raw.upper() == "ITEM":
        raw = (str(item.get("name") or "").strip()[:1] or "?").upper()
    return raw


def icon_for_input(item: Item) -> str:
    raw = str(item.get("icon") or "").strip()
    return "" if raw.upper() == "[ITEM]" else raw


def _safe_filename(original_name: str) -> str:
    stem = Path(original_name).stem
    ext = Path(original_name).suffix.lower()
    if ext not in {f".{t}" for t in ALLOWED_IMAGE_TYPES}:
        ext = ".png"
    clean_stem = re.sub(r"[^a-zA-Z0-9_-]+", "_", stem).strip("_") or "brinde"
    return f"{clean_stem}_{uuid4().hex[:8]}{ext}"


def save_uploaded_image(uploaded_file: Any) -> str:
    if uploaded_file is None:
        return ""
    ASSETS_DIR.mkdir(parents=True, exist_ok=True)
    target = ASSETS_DIR / _safe_filename(uploaded_file.name)
    raw_bytes = uploaded_file.getvalue()
    try:
        img = Image.open(io.BytesIO(raw_bytes))
        img.thumbnail((600, 600))
        img.save(target)
    except Exception:
        target.write_bytes(raw_bytes)
    return str(target)


def suggest_next_kit_code(history: list[Item], offset: int = 0) -> str:
    max_num = 0
    for row in history:
        match = re.search(r"(\d+)", str(row.get("kit_code", "")))
        if match:
            max_num = max(max_num, int(match.group(1)))
    return f"Kit {max_num + 1 + max(0, int(offset))}"


def load_catalog_by_id() -> dict[int, Item]:
    return {int(item["id"]): item for item in get_catalog_items()}


def compute_totals(kit_items: dict[int, int], catalog_by_id: dict[int, Item]) -> KitTotals:
    totals = KitTotals()
    for item_id, qty in kit_items.items():
        item = catalog_by_id.get(int(item_id))
        if item is None or int(qty) <= 0:
            continue
        unit_price = float(item["unit_price"])
        subtotal = round(unit_price * int(qty), 2)
        totals.lines.append(
            {
                "item_id": int(item_id),
                "item": item,
                "quantity": int(qty),
                "unit_price": unit_price,
                "subtotal": subtotal,
            }
        )
        totals.quantity += int(qty)
        totals.value = round(totals.value + subtotal, 2)
    return totals


def single_kit_export_row(summary: dict[str, Any], money: Callable[[float], Any]) -> dict[str, Any]:
    """Linha padronizada com as 4 colunas obrigatórias para colar ou baixar em Excel."""
    return {
        "Itens e quantidades": summary.get("items_summary", ""),
        "Nome": summary.get("client_name", ""),
        "Endereço": summary.get("address_summary", ""),
        "Valor": money(float(summary.get("total_value", 0.0))),
    }


def build_batch_excel_bytes(summaries: list[dict[str, Any]]) -> bytes:
    """Gera um arquivo .xlsx com uma linha por kit finalizado contendo as 4 colunas obrigatórias."""
    rows = [single_kit_export_row(s, float) for s in summaries]
    df = pd.DataFrame(rows)
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="Kits_Finalizados")
        ws = writer.sheets["Kits_Finalizados"]
        if not df.empty:
            for r_idx in range(2, ws.max_row + 1):
                ws.cell(row=r_idx, column=4).number_format = "R$ #,##0.00"
            for idx, col in enumerate(df.columns, start=1):
                longest = max(len(str(col)), *(len(str(v)) for v in df[col]))
                ws.column_dimensions[get_column_letter(idx)].width = min(longest + 4, 65)
    return output.getvalue()


def build_batch_tsv(summaries: list[dict[str, Any]]) -> str:
    """Linhas separadas por TAB: ao copiar e colar no Excel, preenche 4 colunas por linha."""
    lines = []
    for s in summaries:
        row = single_kit_export_row(s, format_brl)
        lines.append("\t".join(str(v) for v in row.values()))
    return "\n".join(lines)


def _empty_recipient(kit_code: str = "", status: str = "") -> dict[str, str]:
    return {
        "kit_code": kit_code,
        "client_name": "",
        "status": status or STATUS_OPTIONS[0],
        "notes": "",
        "cep": "",
        "street": "",
        "address_number": "",
        "complement": "",
        "neighborhood": "",
        "city_state": "",
    }


# =============================================================================
# 3. ESTADO E AÇÕES (CALLBACKS)
# =============================================================================


def init_session_state() -> None:
    ss = st.session_state
    if "kit_items" in ss:
        return
    history = get_kits_history()
    ss.kit_items = {}
    ss.editing_kit_id = None
    ss.form_nonce = 0
    for name in FORM_FIELDS:
        ss[f"form_{name}"] = ""
    ss.form_kit_code = suggest_next_kit_code(history)
    ss.form_status = STATUS_OPTIONS[0]

    ss.batch_mode = False
    ss.batch_nonce = 0
    ss.batch_recipients = [
        _empty_recipient(suggest_next_kit_code(history, 0), STATUS_OPTIONS[0]),
        _empty_recipient(suggest_next_kit_code(history, 1), STATUS_OPTIONS[0]),
    ]

    ss.last_summary = None
    ss.last_batch_summaries = []
    ss.flash = None
    ss.nav_page = "montagem"


def set_form(**values: str) -> None:
    """Substitui todos os campos do formulário e recria os widgets."""
    ss = st.session_state
    for name in FORM_FIELDS:
        ss[f"form_{name}"] = str(values.get(name) or "")
    if ss.form_status not in STATUS_OPTIONS:
        ss.form_status = STATUS_OPTIONS[0]
    ss.form_nonce = int(ss.form_nonce) + 1


def current_form() -> dict[str, str]:
    """Valores atuais do formulário, inclusive o que foi digitado agora."""
    ss = st.session_state
    nonce = ss.form_nonce
    return {
        name: str(ss.get(f"w_{name}_{nonce}", ss[f"form_{name}"]) or "")
        for name in FORM_FIELDS
    }


def _sync_batch_from_widgets() -> list[dict[str, str]]:
    ss = st.session_state
    nonce = ss.batch_nonce
    synced: list[dict[str, str]] = []
    for idx, base in enumerate(ss.batch_recipients):
        row = dict(base)
        for name in FORM_FIELDS:
            widget_key = f"wb_{idx}_{name}_{nonce}"
            if widget_key in ss:
                row[name] = str(ss[widget_key] or "")
        if row.get("status") not in STATUS_OPTIONS:
            row["status"] = STATUS_OPTIONS[0]
        synced.append(row)
    ss.batch_recipients = synced
    return synced


def add_batch_recipient() -> None:
    ss = st.session_state
    rows = _sync_batch_from_widgets()
    if len(rows) >= MAX_BATCH_KITS:
        notify("warning", f"Limite de {MAX_BATCH_KITS} destinatários por lote atingido.")
        return
    next_code = suggest_next_kit_code(get_kits_history(), offset=len(rows))
    rows.append(_empty_recipient(next_code, STATUS_OPTIONS[0]))
    ss.batch_recipients = rows
    ss.batch_nonce = int(ss.batch_nonce) + 1


def remove_batch_recipient(idx: int) -> None:
    ss = st.session_state
    rows = _sync_batch_from_widgets()
    if len(rows) <= 1:
        notify("warning", "Mantenha ao menos 1 destinatário no lote.")
        return
    rows.pop(idx)
    ss.batch_recipients = rows
    ss.batch_nonce = int(ss.batch_nonce) + 1


def reset_batch_recipients() -> None:
    ss = st.session_state
    history = get_kits_history()
    ss.batch_recipients = [
        _empty_recipient(suggest_next_kit_code(history, 0), STATUS_OPTIONS[0]),
        _empty_recipient(suggest_next_kit_code(history, 1), STATUS_OPTIONS[0]),
    ]
    ss.batch_nonce = int(ss.batch_nonce) + 1


def toggle_batch_mode(enabled: bool) -> None:
    ss = st.session_state
    ss.batch_mode = bool(enabled)
    if enabled:
        ss.editing_kit_id = None
        history = get_kits_history()
        rows = _sync_batch_from_widgets()
        for idx, r in enumerate(rows):
            if not str(r.get("kit_code") or "").strip():
                r["kit_code"] = suggest_next_kit_code(history, offset=idx)
        ss.batch_recipients = rows
        ss.batch_nonce = int(ss.batch_nonce) + 1


def notify(kind: str, text: str) -> None:
    st.session_state.flash = (kind, text)


def show_flash() -> None:
    flash = st.session_state.get("flash")
    if flash:
        st.toast(flash[1])
        st.session_state.flash = None


def go_to(page: str) -> None:
    st.session_state.nav_page = page


def add_item(item_id: int, qty: int = 1) -> None:
    items = dict(st.session_state.kit_items)
    items[item_id] = int(items.get(item_id, 0)) + qty
    if items[item_id] <= 0:
        items.pop(item_id)
    st.session_state.kit_items = items
    st.session_state.last_summary = None
    st.session_state.last_batch_summaries = []


def change_qty(item_id: int, delta: int) -> None:
    items = dict(st.session_state.kit_items)
    new_qty = int(items.get(item_id, 0)) + delta
    if new_qty <= 0:
        items.pop(item_id, None)
    else:
        items[item_id] = new_qty
    st.session_state.kit_items = items


def remove_item(item_id: int) -> None:
    items = dict(st.session_state.kit_items)
    items.pop(item_id, None)
    st.session_state.kit_items = items


def apply_recipe(recipe_name: str, recipe_items: dict[int, int]) -> None:
    catalog_by_id = load_catalog_by_id()
    items = dict(st.session_state.kit_items)
    for item_id, qty in recipe_items.items():
        if int(item_id) in catalog_by_id and int(qty) > 0:
            items[int(item_id)] = int(items.get(int(item_id), 0)) + int(qty)
    st.session_state.kit_items = items
    st.session_state.last_summary = None
    st.session_state.last_batch_summaries = []
    notify("info", f"Receita “{recipe_name}” aplicada ao kit.")


def remove_recipe_direct(recipe_id: int) -> None:
    ok, msg = delete_recipe(recipe_id)
    notify("success" if ok else "warning", msg)


def reset_kit() -> None:
    """Esvazia o kit e o formulário (também cancela uma edição em andamento)."""
    st.session_state.kit_items = {}
    st.session_state.editing_kit_id = None
    set_form(kit_code=suggest_next_kit_code(get_kits_history()))
    reset_batch_recipients()


def cancel_edit() -> None:
    reset_kit()
    notify("info", "Edição cancelada. Nenhuma alteração foi salva.")


def finalize_kit() -> None:
    ss = st.session_state
    form = current_form()
    catalog_by_id = load_catalog_by_id()
    totals = compute_totals(ss.kit_items, catalog_by_id)
    editing_id = ss.editing_kit_id

    ok, msg, saved_id = save_kit(
        kit_code=form["kit_code"],
        client_name=form["client_name"],
        notes=form["notes"],
        status=form["status"],
        kit_items=ss.kit_items,
        catalog_by_id=catalog_by_id,
        cep=form["cep"],
        street=form["street"],
        address_number=form["address_number"],
        complement=form["complement"],
        neighborhood=form["neighborhood"],
        city_state=form["city_state"],
        kit_id=editing_id,
    )
    if not ok:
        notify("warning", msg)
        return

    items_summary = ", ".join(
        f"{line['quantity']}x {line['item']['name']}" for line in totals.lines
    )
    address_summary = format_address_summary(form)
    kit_code_clean = form["kit_code"].strip() or "Kit"

    ss.last_batch_summaries = []
    ss.last_summary = {
        "id": saved_id,
        "kit_code": kit_code_clean,
        "client_name": form["client_name"].strip(),
        "address_summary": address_summary,
        "items_summary": items_summary,
        "total_quantity": totals.quantity,
        "total_value": totals.value,
        "edited": editing_id is not None,
    }
    reset_kit()
    if editing_id is not None:
        go_to("historico")


def finalize_batch_kits() -> None:
    ss = st.session_state
    rows = _sync_batch_from_widgets()
    catalog_by_id = load_catalog_by_id()
    totals = compute_totals(ss.kit_items, catalog_by_id)

    if not totals.lines or totals.quantity <= 0:
        notify("warning", "Adicione ao menos 1 item à caixa de brindes antes de gerar os kits.")
        return

    filled_rows = [r for r in rows if str(r.get("client_name") or "").strip()]
    if not filled_rows:
        notify("warning", "Informe o nome de pelo menos 1 destinatário para gerar os kits.")
        return

    items_summary = ", ".join(
        f"{line['quantity']}x {line['item']['name']}" for line in totals.lines
    )

    saved_summaries: list[dict[str, Any]] = []
    for idx, r in enumerate(filled_rows):
        code_clean = str(r.get("kit_code") or "").strip() or suggest_next_kit_code(get_kits_history(), idx)
        client_clean = str(r.get("client_name") or "").strip()
        ok, msg, saved_id = save_kit(
            kit_code=code_clean,
            client_name=client_clean,
            notes=str(r.get("notes") or ""),
            status=str(r.get("status") or STATUS_OPTIONS[0]),
            kit_items=ss.kit_items,
            catalog_by_id=catalog_by_id,
            cep=str(r.get("cep") or ""),
            street=str(r.get("street") or ""),
            address_number=str(r.get("address_number") or ""),
            complement=str(r.get("complement") or ""),
            neighborhood=str(r.get("neighborhood") or ""),
            city_state=str(r.get("city_state") or ""),
            kit_id=None,
        )
        if not ok:
            notify("warning", f"{client_clean}: {msg}")
            return

        saved_summaries.append(
            {
                "id": saved_id,
                "kit_code": code_clean,
                "client_name": client_clean,
                "address_summary": format_address_summary(r),
                "items_summary": items_summary,
                "total_quantity": totals.quantity,
                "total_value": totals.value,
                "edited": False,
            }
        )

    ss.last_summary = None
    ss.last_batch_summaries = saved_summaries
    ss.kit_items = {}
    ss.editing_kit_id = None
    set_form(kit_code=suggest_next_kit_code(get_kits_history()))
    reset_batch_recipients()
    notify("success", f"{len(saved_summaries)} kits gerados e salvos no histórico.")


def load_kit_for_edit(kit: Item) -> None:
    ss = st.session_state
    catalog_by_id = load_catalog_by_id()
    loaded: dict[int, int] = {}
    missing: list[str] = []
    for it in kit["items"]:
        item_id = int(it["item_id"])
        if item_id in catalog_by_id:
            loaded[item_id] = int(it["quantity"])
        else:
            missing.append(str(it["item_name"]))

    ss.batch_mode = False
    ss.kit_items = loaded
    ss.editing_kit_id = int(kit["id"])
    ss.last_summary = None
    ss.last_batch_summaries = []
    set_form(**{name: kit.get(name) for name in FORM_FIELDS})
    go_to("montagem")

    msg = f"Editando “{kit['kit_code']}”."
    if missing:
        msg += f" Itens indisponíveis: {', '.join(missing)}."
    notify("info", msg)


def duplicate_kit(kit: Item) -> None:
    ss = st.session_state
    catalog_by_id = load_catalog_by_id()
    ss.batch_mode = False
    ss.kit_items = {
        int(it["item_id"]): int(it["quantity"])
        for it in kit["items"]
        if int(it["item_id"]) in catalog_by_id
    }
    ss.editing_kit_id = None
    ss.last_summary = None
    ss.last_batch_summaries = []
    values = {name: kit.get(name) for name in FORM_FIELDS}
    values["kit_code"] = suggest_next_kit_code(get_kits_history())
    values["status"] = STATUS_OPTIONS[0]
    set_form(**values)
    go_to("montagem")
    notify("info", f"“{kit['kit_code']}” duplicado. Ajuste os dados e finalize.")


def delete_kit_direct(kit_id: int) -> None:
    ok, msg = delete_kit(kit_id)
    if ok and st.session_state.editing_kit_id == kit_id:
        reset_kit()
    notify("success" if ok else "warning", msg)


# =============================================================================
# 4. ESTILOS
# =============================================================================

CSS = """
<style>
    .block-container { padding-top: 1.5rem; padding-bottom: 3rem; max-width: 1340px; }

    .hero {
        display: flex; justify-content: space-between; align-items: center; gap: 16px;
        background: linear-gradient(135deg, #1e1b4b 0%, #3730a3 100%);
        color: #fff; border-radius: 16px; padding: 18px 26px; margin-bottom: 14px;
    }
    .hero-title { font-size: 1.5rem; font-weight: 800; line-height: 1.2; }
    .hero-sub { margin-top: 3px; font-size: 0.88rem; color: rgba(255,255,255,0.75); }
    .hero-pill {
        background: rgba(255,255,255,0.14); border: 1px solid rgba(255,255,255,0.28);
        border-radius: 999px; padding: 6px 16px; font-size: 0.82rem; font-weight: 700;
        white-space: nowrap;
    }

    div[role="radiogroup"] { gap: 0.5rem; flex-wrap: wrap; }
    div[role="radiogroup"] > label {
        background: rgba(99,102,241,0.08); border: 1px solid rgba(99,102,241,0.25);
        border-radius: 999px; padding: 4px 18px; cursor: pointer; margin: 0;
    }
    div[role="radiogroup"] > label:has(input:checked) {
        background: #4f46e5; border-color: #4f46e5;
    }
    div[role="radiogroup"] > label:has(input:checked) * { color: #fff !important; }

    .total-card {
        background: linear-gradient(135deg, #0f172a 0%, #1e293b 100%);
        color: #fff; border-radius: 16px; padding: 16px 22px; margin-bottom: 12px;
        box-shadow: 0 6px 18px rgba(15,23,42,0.18);
    }
    .total-eyebrow {
        font-size: 0.74rem; font-weight: 700; letter-spacing: 0.08em;
        text-transform: uppercase; color: rgba(255,255,255,0.65);
    }
    .total-value { font-size: 2.3rem; font-weight: 800; line-height: 1.15; margin: 4px 0 2px; }
    .total-meta { font-size: 0.86rem; color: rgba(255,255,255,0.75); }

    .icon-tile {
        width: var(--s, 56px); height: var(--s, 56px);
        display: flex; align-items: center; justify-content: center;
        font-size: var(--fs, 20px); font-weight: 800; color: #6366f1;
        overflow: hidden; border-radius: 14px; background: rgba(99,102,241,0.12);
    }
    .chip {
        display: inline-block; font-size: 0.72rem; font-weight: 700;
        padding: 2px 10px; border-radius: 999px;
        background: rgba(99,102,241,0.12); color: #6366f1;
    }
    .chip-warn { background: rgba(245,158,11,0.16); color: #b45309; }
    .chip-info { background: rgba(59,130,246,0.16); color: #1d4ed8; }
    .chip-ok   { background: rgba(16,185,129,0.16); color: #047857; }
    .qty-badge {
        display: inline-block; background: #4f46e5; color: #fff;
        font-size: 0.78rem; font-weight: 700; padding: 1px 10px; border-radius: 999px;
    }
    .price { font-size: 1.02rem; font-weight: 700; }
    .kit-value { font-size: 1.5rem; font-weight: 800; text-align: right; line-height: 1.2; }
    .kit-meta { text-align: right; opacity: 0.65; font-size: 0.82rem; }
    .box-title { font-weight: 800; font-size: 1rem; }
    .tile-name { text-align: center; font-weight: 700; font-size: 0.88rem; margin-top: 4px; }
    .tile-qty { text-align: center; margin: 4px 0 2px; }
    .tile-sub { text-align: center; opacity: 0.65; font-size: 0.78rem; margin-bottom: 6px; }
</style>
"""


def render_styles() -> None:
    st.markdown(CSS, unsafe_allow_html=True)


# =============================================================================
# 5. COMPONENTES DE INTERFACE
# =============================================================================


def render_item_visual(item: Item, size: int = 56) -> None:
    img_path = item.get("image_path") or ""
    if img_path and Path(img_path).exists():
        st.image(img_path, width=size)
        return
    text = icon_text(item)
    factor = 0.5 if len(text) <= 1 else 0.34 if len(text) <= 3 else 0.24
    st.markdown(
        f"<div class='icon-tile' style='--s:{size}px;--fs:{int(size * factor)}px'>{esc(text)}</div>",
        unsafe_allow_html=True,
    )


def render_header(history_count: int) -> None:
    in_block = history_count % MAX_BATCH_KITS
    st.markdown(
        f"""
        <div class="hero">
            <div>
                <div class="hero-title">{APP_TITLE}</div>
                <div class="hero-sub">{APP_SUBTITLE}</div>
            </div>
            <div class="hero-pill">Bloco atual: {in_block}/{MAX_BATCH_KITS} kits</div>
        </div>
        """,
        unsafe_allow_html=True,
    )


def render_nav(history_count: int) -> str:
    labels = dict(PAGES)
    labels["historico"] = f"{PAGES['historico']} ({history_count})"
    return st.radio(
        "Navegação",
        options=list(labels),
        format_func=labels.get,
        key="nav_page",
        horizontal=True,
        label_visibility="collapsed",
    )


def render_last_summary() -> None:
    batch = st.session_state.get("last_batch_summaries") or []
    if batch:
        with st.container(border=True):
            total_lote = sum(float(s.get("total_value", 0.0)) for s in batch)
            st.success(
                f"**{len(batch)} kits gerados e salvos.** "
                f"Valor total do lote: {format_brl(total_lote)}."
            )
            st.caption(
                "Copie as linhas abaixo (ícone de copiar no canto direito) e cole direto no Excel "
                "(uma linha por pessoa com: Itens e quantidades | Nome | Endereço | Valor):"
            )
            st.code(build_batch_tsv(batch), language=None)

            c_dl, _ = st.columns([1.4, 2.6])
            c_dl.download_button(
                "Baixar planilha do lote (.xlsx)",
                data=build_batch_excel_bytes(batch),
                file_name="lote_kits_momento_uau.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                key="dl_batch_kits",
                **STRETCH,
            )
        return

    summary = st.session_state.last_summary
    if not summary:
        return

    with st.container(border=True):
        if summary.get("edited"):
            st.success(f"**{summary['kit_code']} atualizado.** Novo valor: {format_brl(summary['total_value'])}.")
        else:
            st.success(
                f"**{summary['kit_code']} finalizado.** "
                f"Registre {format_brl(summary['total_value'])} no formulário."
            )

        st.caption(
            "Copie a linha abaixo (ícone de copiar no canto direito) e cole direto no Excel "
            "(preenche automaticamente: Itens e quantidades | Nome | Endereço | Valor):"
        )
        st.code(build_batch_tsv([summary]), language=None)

        c_dl, _ = st.columns([1.3, 2.7])
        safe_code = re.sub(r"[^a-zA-Z0-9_-]+", "_", summary["kit_code"]).strip("_") or "kit"
        c_dl.download_button(
            "Baixar planilha deste kit (.xlsx)",
            data=build_batch_excel_bytes([summary]),
            file_name=f"{safe_code}_momento_uau.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            key=f"dl_single_kit_{summary.get('id', 'novo')}",
            **STRETCH,
        )


def render_recipes(recipes: list[Item], catalog: list[Item], catalog_by_id: dict[int, Item]) -> None:
    st.markdown("#### Receitas predefinidas")
    tab_use, tab_save, tab_edit = st.tabs(["Aplicar receita", "Salvar composição", "Editar / excluir"])

    with tab_use:
        if not recipes:
            st.caption("Nenhuma receita cadastrada.")
        for row in chunked(recipes, 3):
            cols = st.columns(3, gap="small")
            for col, rec in zip(cols, row):
                with col, st.container(border=True):
                    st.markdown(f"**{esc(rec['name'])}**")
                    parts = [
                        f"{qty}× {catalog_by_id[iid]['name']}"
                        for iid, qty in rec["items"].items()
                        if iid in catalog_by_id
                    ]
                    st.caption(", ".join(parts) if parts else "Sem itens")
                    st.button(
                        "Aplicar",
                        key=f"apply_recipe_{rec['id']}",
                        on_click=apply_recipe,
                        args=(rec["name"], rec["items"]),
                        **STRETCH,
                    )

    with tab_save:
        with st.form("save_recipe_form", clear_on_submit=True):
            c1, c2 = st.columns([1.2, 1.3])
            r_name = c1.text_input("Nome da receita", placeholder="Ex.: Kit D – VIP")
            r_desc = c2.text_input("Observação", placeholder="Opcional")
            if st.form_submit_button("Salvar como receita", **STRETCH):
                ok, msg = save_recipe(r_name, r_desc, st.session_state.kit_items)
                notify("success" if ok else "warning", msg)
                st.rerun()

    with tab_edit:
        if not recipes:
            st.caption("Nenhuma receita cadastrada.")
        for rec in recipes:
            rid = int(rec["id"])
            with st.expander(f"{rec['name']}"):
                with st.form(f"edit_recipe_form_{rid}"):
                    ed_name = st.text_input("Nome", value=rec["name"], key=f"rec_name_{rid}")
                    ed_desc = st.text_input("Descrição", value=rec.get("description") or "", key=f"rec_desc_{rid}")
                    st.markdown("**Quantidades por brinde**")
                    new_map: dict[int, int] = {}
                    for item in catalog:
                        iid = int(item["id"])
                        q_val = st.number_input(
                            f"{item['name']} ({format_brl(item['unit_price'])})",
                            min_value=0,
                            value=int(rec["items"].get(iid, 0)),
                            step=1,
                            key=f"rec_qty_{rid}_{iid}",
                        )
                        if q_val > 0:
                            new_map[iid] = int(q_val)
                    if st.form_submit_button("Salvar receita", **STRETCH):
                        ok, msg = update_recipe(rid, ed_name, ed_desc, new_map)
                        notify("success" if ok else "warning", msg)
                        st.rerun()

                st.button(
                    "Excluir receita",
                    key=f"del_recipe_{rid}",
                    on_click=remove_recipe_direct,
                    args=(rid,),
                    **STRETCH,
                )


def render_catalog_cards(catalog: list[Item]) -> None:
    st.markdown("#### Catálogo de brindes")

    categories = ["Todas"] + sorted({str(it.get("category") or "Geral") for it in catalog})
    f1, f2 = st.columns([1.4, 1.0])
    search = f1.text_input("Buscar", placeholder="Ex.: garrafa, blusa…", key="catalog_search")
    category = f2.selectbox("Categoria", categories, key="catalog_category_filter")

    term = search.strip().lower()
    filtered = [
        it
        for it in catalog
        if (category == "Todas" or (it.get("category") or "Geral") == category)
        and (not term or term in it["name"].lower())
    ]
    if not filtered:
        st.info("Nenhum brinde encontrado para o filtro atual.")
        return

    for row in chunked(filtered, 2):
        cols = st.columns(2, gap="medium")
        for col, item in zip(cols, row):
            item_id = int(item["id"])
            in_kit = int(st.session_state.kit_items.get(item_id, 0))
            with col, st.container(border=True):
                c_vis, c_info = st.columns([0.9, 2.1])
                with c_vis:
                    render_item_visual(item, size=60)
                with c_info:
                    st.markdown(
                        f"<span class='chip'>{esc(item.get('category') or 'Geral')}</span>",
                        unsafe_allow_html=True,
                    )
                    st.markdown(f"**{esc(item['name'])}**")
                    st.markdown(
                        f"<span class='price'>{format_brl(item['unit_price'])}</span>",
                        unsafe_allow_html=True,
                    )
                st.button(
                    f"Adicionar ({in_kit} no kit)" if in_kit else "Adicionar",
                    key=f"add_catalog_item_{item_id}",
                    type="secondary" if in_kit else "primary",
                    on_click=add_item,
                    args=(item_id,),
                    **STRETCH,
                )


def render_edit_banner() -> None:
    editing_id = st.session_state.editing_kit_id
    if editing_id is None:
        return
    with st.container(border=True):
        c1, c2 = st.columns([2.2, 1.0])
        c1.markdown(
            f"**Editando o registro #{editing_id}** — "
            "as alterações só valem depois de clicar em *Salvar alterações*."
        )
        c2.button("Cancelar edição", key="btn_cancel_edit", on_click=cancel_edit, **STRETCH)


def render_total_card(totals: KitTotals, recipients_count: int = 1) -> None:
    plural = "item" if totals.quantity == 1 else "itens"
    if recipients_count > 1:
        lote_total = round(totals.value * recipients_count, 2)
        st.markdown(
            f"""
            <div class="total-card">
                <div class="total-eyebrow">Valor por kit (lote de {recipients_count} pessoas)</div>
                <div class="total-value">{format_brl(totals.value)}</div>
                <div class="total-meta">{totals.quantity} {plural} por kit · Total do lote: {format_brl(lote_total)}</div>
            </div>
            """,
            unsafe_allow_html=True,
        )
        return

    st.markdown(
        f"""
        <div class="total-card">
            <div class="total-eyebrow">Valor para preencher no formulário</div>
            <div class="total-value">{format_brl(totals.value)}</div>
            <div class="total-meta">{totals.quantity} {plural} no kit</div>
        </div>
        """,
        unsafe_allow_html=True,
    )


def render_gift_box(totals: KitTotals) -> None:
    with st.container(border=True):
        st.markdown("<span class='box-title'>Caixa de brindes</span>", unsafe_allow_html=True)
        if not totals.lines:
            st.caption("Caixa vazia. Adicione itens do catálogo ou aplique uma receita.")
            return
        for row in chunked(totals.lines, 3):
            cols = st.columns(3, gap="small")
            for col, line in zip(cols, row):
                item_id = line["item_id"]
                with col, st.container(border=True):
                    render_item_visual(line["item"], size=48)
                    st.markdown(
                        f"<div class='tile-name'>{esc(line['item']['name'])}</div>"
                        f"<div class='tile-qty'><span class='qty-badge'>× {line['quantity']}</span></div>"
                        f"<div class='tile-sub'>{format_brl(line['subtotal'])}</div>",
                        unsafe_allow_html=True,
                    )
                    b1, b2 = st.columns(2)
                    b1.button("−", key=f"box_dec_{item_id}", on_click=change_qty, args=(item_id, -1), **STRETCH)
                    b2.button("+", key=f"box_inc_{item_id}", on_click=change_qty, args=(item_id, 1), **STRETCH)
                    st.button(
                        "Retirar",
                        key=f"box_remove_{item_id}",
                        on_click=remove_item,
                        args=(item_id,),
                        **STRETCH,
                    )


def render_kit_form() -> None:
    ss = st.session_state
    nonce = ss.form_nonce

    def text(name: str, label: str, placeholder: str = "") -> str:
        return st.text_input(
            label, value=ss[f"form_{name}"], key=f"w_{name}_{nonce}", placeholder=placeholder
        )

    with st.container(border=True):
        st.markdown("**Dados do kit**")
        c1, c2 = st.columns(2)
        with c1:
            ss.form_kit_code = text("kit_code", "Número ou nome do kit", "Ex.: Kit 1")
        with c2:
            ss.form_client_name = text("client_name", "Nome do cliente *", "Ex.: Cliente Novo")

        c3, c4 = st.columns([1.0, 1.4])
        with c3:
            status_idx = STATUS_OPTIONS.index(ss.form_status) if ss.form_status in STATUS_OPTIONS else 0
            ss.form_status = st.selectbox(
                "Status", STATUS_OPTIONS, index=status_idx, key=f"w_status_{nonce}"
            )
        with c4:
            ss.form_notes = text("notes", "Observações (opcional)", "Ex.: Entregar na recepção")

        has_address = any(ss[f"form_{name}"] for name in ADDRESS_FIELDS)
        with st.expander("Endereço de entrega (opcional)", expanded=has_address):
            a1, a2, a3 = st.columns([1.0, 2.0, 0.8])
            with a1:
                ss.form_cep = text("cep", "CEP", "00000-000")
            with a2:
                ss.form_street = text("street", "Rua / Logradouro", "Ex.: Av. Paulista")
            with a3:
                ss.form_address_number = text("address_number", "Número", "1000")
            a4, a5, a6 = st.columns([1.0, 1.2, 1.4])
            with a4:
                ss.form_complement = text("complement", "Complemento", "Sala / Andar")
            with a5:
                ss.form_neighborhood = text("neighborhood", "Bairro", "Ex.: Bela Vista")
            with a6:
                ss.form_city_state = text("city_state", "Cidade / UF", "Ex.: São Paulo - SP")


def render_batch_form() -> int:
    ss = st.session_state
    nonce = ss.batch_nonce
    rows = ss.batch_recipients
    active_count = 0

    for idx, row in enumerate(rows):
        with st.container(border=True):
            h1, h2 = st.columns([3.0, 1.0])
            h1.markdown(f"**Destinatário {idx + 1}**")
            if len(rows) > 1:
                h2.button(
                    "Remover",
                    key=f"btn_rm_batch_{idx}_{nonce}",
                    on_click=remove_batch_recipient,
                    args=(idx,),
                    **STRETCH,
                )

            c1, c2, c3 = st.columns([1.0, 1.6, 1.0])
            code_val = c1.text_input(
                "Kit",
                value=row.get("kit_code", ""),
                key=f"wb_{idx}_kit_code_{nonce}",
                placeholder=f"Kit {idx + 1}",
            )
            name_val = c2.text_input(
                "Nome *",
                value=row.get("client_name", ""),
                key=f"wb_{idx}_client_name_{nonce}",
                placeholder="Nome da pessoa",
            )
            st_curr = row.get("status") if row.get("status") in STATUS_OPTIONS else STATUS_OPTIONS[0]
            status_val = c3.selectbox(
                "Status",
                STATUS_OPTIONS,
                index=STATUS_OPTIONS.index(st_curr),
                key=f"wb_{idx}_status_{nonce}",
            )

            notes_val = st.text_input(
                "Observações (opcional)",
                value=row.get("notes", ""),
                key=f"wb_{idx}_notes_{nonce}",
                placeholder="Ex.: Entregar na recepção",
            )

            has_addr = any(row.get(f) for f in ADDRESS_FIELDS)
            with st.expander("Endereço de entrega (opcional)", expanded=has_addr):
                a1, a2, a3 = st.columns([1.0, 2.0, 0.8])
                cep_val = a1.text_input("CEP", value=row.get("cep", ""), key=f"wb_{idx}_cep_{nonce}", placeholder="00000-000")
                street_val = a2.text_input("Rua / Logradouro", value=row.get("street", ""), key=f"wb_{idx}_street_{nonce}", placeholder="Ex.: Av. Paulista")
                num_val = a3.text_input("Número", value=row.get("address_number", ""), key=f"wb_{idx}_address_number_{nonce}", placeholder="1000")

                a4, a5, a6 = st.columns([1.0, 1.2, 1.4])
                comp_val = a4.text_input("Complemento", value=row.get("complement", ""), key=f"wb_{idx}_complement_{nonce}", placeholder="Sala / Andar")
                neigh_val = a5.text_input("Bairro", value=row.get("neighborhood", ""), key=f"wb_{idx}_neighborhood_{nonce}", placeholder="Ex.: Bela Vista")
                city_val = a6.text_input("Cidade / UF", value=row.get("city_state", ""), key=f"wb_{idx}_city_state_{nonce}", placeholder="Ex.: São Paulo - SP")

            rows[idx] = {
                "kit_code": code_val,
                "client_name": name_val,
                "status": status_val,
                "notes": notes_val,
                "cep": cep_val,
                "street": street_val,
                "address_number": num_val,
                "complement": comp_val,
                "neighborhood": neigh_val,
                "city_state": city_val,
            }
            if name_val.strip():
                active_count += 1

    st.button(
        "+ Adicionar pessoa ao lote",
        key=f"btn_add_batch_{nonce}",
        disabled=len(rows) >= MAX_BATCH_KITS,
        on_click=add_batch_recipient,
        **STRETCH,
    )
    return max(1, active_count)


def render_kit_panel(catalog_by_id: dict[int, Item]) -> None:
    ss = st.session_state
    totals = compute_totals(ss.kit_items, catalog_by_id)
    editing = ss.editing_kit_id is not None

    if not editing:
        m1, m2 = st.columns(2)
        m1.button(
            "1 pessoa",
            key="mode_single",
            type="primary" if not ss.batch_mode else "secondary",
            on_click=toggle_batch_mode,
            args=(False,),
            **STRETCH,
        )
        m2.button(
            "Várias pessoas (lote)",
            key="mode_batch",
            type="primary" if ss.batch_mode else "secondary",
            on_click=toggle_batch_mode,
            args=(True,),
            **STRETCH,
        )

    render_edit_banner()

    if ss.batch_mode and not editing:
        active_people = render_batch_form()
        render_total_card(totals, recipients_count=active_people)
        render_gift_box(totals)
        st.write("")
        a1, a2 = st.columns([1.6, 1.0])
        a1.button(
            "Gerar kits do lote",
            key="btn_finalize_batch",
            type="primary",
            disabled=not totals.lines,
            on_click=finalize_batch_kits,
            **STRETCH,
        )
        a2.button("Limpar", key="btn_clear_batch", on_click=reset_kit, **STRETCH)
        return

    render_total_card(totals, recipients_count=1)
    render_gift_box(totals)
    render_kit_form()

    st.write("")
    a1, a2 = st.columns([1.6, 1.0])
    a1.button(
        "Salvar alterações" if editing else "Finalizar kit",
        key="btn_finalize_kit",
        type="primary",
        disabled=not totals.lines,
        on_click=finalize_kit,
        **STRETCH,
    )
    a2.button("Limpar", key="btn_clear_kit", on_click=reset_kit, **STRETCH)


def render_new_item_form() -> None:
    st.markdown("#### Cadastrar novo brinde")
    with st.form("add_catalog_form", clear_on_submit=True):
        c1, c2, c3, c4 = st.columns([2.0, 1.2, 1.0, 1.5])
        name = c1.text_input("Nome do brinde *", placeholder="Ex.: Mochila")
        price = c2.number_input("Valor unitário (R$) *", min_value=0.0, value=20.0, step=1.0, format="%.2f")
        icon = c3.text_input("Sigla (opcional)", placeholder="Ex.: MOC", max_chars=12)
        category = c4.text_input("Categoria (opcional)", placeholder="Ex.: Acessórios")
        image = st.file_uploader("Imagem do brinde (opcional)", type=ALLOWED_IMAGE_TYPES, key="new_item_uploader")

        if st.form_submit_button("Cadastrar brinde", type="primary"):
            image_path = save_uploaded_image(image) if image else ""
            ok, msg = add_catalog_item(
                name=name, unit_price=price, icon=icon, category=category, image_path=image_path
            )
            notify("success" if ok else "warning", msg)
            st.rerun()


def render_edit_item(item: Item) -> None:
    item_id = int(item["id"])
    label = f"{item['name']}  ·  {format_brl(item['unit_price'])}  ·  {item.get('category') or 'Geral'}"
    with st.expander(label):
        c_img, c_form = st.columns([1.0, 3.2], gap="medium")
        with c_img:
            render_item_visual(item, size=92)

        with c_form:
            with st.form(f"edit_catalog_form_{item_id}"):
                f1, f2, f3, f4 = st.columns([2.0, 1.2, 1.0, 1.5])
                name = f1.text_input("Nome", value=item["name"], key=f"ed_name_{item_id}")
                price = f2.number_input(
                    "Valor unitário (R$)", min_value=0.0, value=float(item["unit_price"]),
                    step=1.0, format="%.2f", key=f"ed_price_{item_id}",
                )
                icon = f3.text_input(
                    "Sigla", value=icon_for_input(item), max_chars=12, key=f"ed_icon_{item_id}"
                )
                category = f4.text_input(
                    "Categoria", value=item.get("category") or "Geral", key=f"ed_cat_{item_id}"
                )
                upload = st.file_uploader(
                    "Substituir imagem (opcional)", type=ALLOWED_IMAGE_TYPES, key=f"ed_upload_{item_id}"
                )
                if st.form_submit_button("Salvar alterações"):
                    image_path = item.get("image_path") or ""
                    if upload is not None:
                        image_path = save_uploaded_image(upload)
                    ok, msg = update_catalog_item(
                        item_id=item_id, name=name, unit_price=price,
                        icon=icon, category=category, image_path=image_path,
                    )
                    notify("success" if ok else "warning", msg)
                    st.rerun()

            if st.button("Excluir brinde", key=f"del_catalog_item_{item_id}"):
                ok, msg = delete_catalog_item(
                    item_id=item_id, active_kit_items=st.session_state.kit_items
                )
                notify("success" if ok else "warning", msg)
                st.rerun()


def history_rows(history: list[Item], money: Callable[[float], Any]) -> list[dict[str, Any]]:
    return [
        {
            "Kit": k["kit_code"],
            "Cliente": k["client_name"],
            "Endereço": k.get("address_summary") or format_address_summary(k),
            "Itens": k["items_summary"],
            "Quantidade total": k["total_quantity"],
            "Valor": money(k["total_value"]),
            "Status": k["status"],
            "Data": k["created_at"],
            "Observações": k.get("notes") or "",
        }
        for k in history
    ]


def build_excel_bytes(history: list[Item]) -> bytes:
    df = pd.DataFrame(history_rows(history, float))
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="Historico_Kits")
        if not df.empty:
            ws = writer.sheets["Historico_Kits"]
            value_col = list(df.columns).index("Valor") + 1
            for row in range(2, ws.max_row + 1):
                ws.cell(row=row, column=value_col).number_format = "R$ #,##0.00"
            for idx, col in enumerate(df.columns, start=1):
                longest = max(len(str(col)), *(len(str(v)) for v in df[col]))
                ws.column_dimensions[get_column_letter(idx)].width = min(longest + 2, 60)
    return output.getvalue()


def build_csv_bytes(history: list[Item]) -> bytes:
    df = pd.DataFrame(history_rows(history, format_brl))
    return df.to_csv(index=False, sep=";").encode("utf-8-sig")


def render_kit_card(kit: Item) -> None:
    kid = int(kit["id"])
    address = kit.get("address_summary") or format_address_summary(kit)
    chip = STATUS_CHIP.get(kit["status"], "chip")

    with st.container(border=True):
        left, right = st.columns([3.2, 1.2])
        with left:
            st.markdown(
                f"**{esc(kit['kit_code'])}** · {esc(kit['client_name'])} "
                f"<span class='chip {chip}'>{esc(kit['status'])}</span>",
                unsafe_allow_html=True,
            )
            st.caption(f"Itens: {kit['items_summary']}")
            details = [f"Data: {kit['created_at']}"]
            if address:
                details.append(f"Endereço: {address}")
            if kit.get("notes"):
                details.append(f"Obs.: {kit['notes']}")
            st.caption("  |  ".join(details))
        with right:
            st.markdown(
                f"<div class='kit-value'>{format_brl(kit['total_value'])}</div>"
                f"<div class='kit-meta'>{kit['total_quantity']} itens</div>",
                unsafe_allow_html=True,
            )

        b1, b2, b3, _ = st.columns([1.0, 1.0, 1.0, 2.2])
        b1.button(
            "Editar", key=f"edit_kit_{kid}", type="primary",
            on_click=load_kit_for_edit, args=(kit,), **STRETCH,
        )
        b2.button("Duplicar", key=f"dup_kit_{kid}", on_click=duplicate_kit, args=(kit,), **STRETCH)
        b3.button("Excluir", key=f"del_kit_{kid}", on_click=delete_kit_direct, args=(kid,), **STRETCH)


def render_history(history: list[Item]) -> None:
    st.markdown("#### Histórico de kits")
    if not history:
        st.info("Nenhum kit registrado ainda.")
        return

    m1, m2, m3 = st.columns(3)
    m1.metric("Kits registrados", len(history))
    m2.metric("Itens no total", sum(int(k["total_quantity"]) for k in history))
    m3.metric("Valor acumulado", format_brl(sum(float(k["total_value"]) for k in history)))

    f1, f2, f3, f4 = st.columns([2.0, 1.2, 1.0, 1.0])
    search = f1.text_input("Buscar kit ou cliente", placeholder="Ex.: Kit 3, nome do cliente…", key="hist_search")
    status = f2.selectbox("Status", ["Todos"] + STATUS_OPTIONS, key="hist_status")
    f3.write("")
    f4.write("")
    f3.download_button(
        "Exportar Excel",
        data=build_excel_bytes(history),
        file_name="historico_kits_momento_uau.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        **STRETCH,
    )
    f4.download_button(
        "Exportar CSV",
        data=build_csv_bytes(history),
        file_name="historico_kits_momento_uau.csv",
        mime="text/csv",
        **STRETCH,
    )

    term = search.strip().lower()
    filtered = [
        k
        for k in history
        if (status == "Todos" or k["status"] == status)
        and (
            not term
            or term in str(k["kit_code"]).lower()
            or term in str(k["client_name"]).lower()
            or term in str(k["items_summary"]).lower()
        )
    ]

    with st.expander("Ver como tabela"):
        table = pd.DataFrame(history_rows(filtered, format_brl)).drop(columns=["Observações"], errors="ignore")
        st.dataframe(table, hide_index=True, **STRETCH)

    st.caption(f"{len(filtered)} de {len(history)} kits")
    if not filtered:
        st.info("Nenhum kit encontrado para o filtro atual.")
    for kit in filtered:
        render_kit_card(kit)


# =============================================================================
# 6. PÁGINAS
# =============================================================================


def page_assembly(catalog: list[Item], catalog_by_id: dict[int, Item], recipes: list[Item]) -> None:
    col_left, col_right = st.columns([1.2, 1.0], gap="large")
    with col_left:
        render_recipes(recipes, catalog, catalog_by_id)
        st.divider()
        render_catalog_cards(catalog)
    with col_right:
        st.markdown("#### Kit atual")
        render_kit_panel(catalog_by_id)


def page_catalog(catalog: list[Item]) -> None:
    render_new_item_form()
    st.divider()
    st.markdown("#### Editar ou excluir brindes")
    if not catalog:
        st.info("Nenhum brinde cadastrado.")
    for item in catalog:
        render_edit_item(item)


# =============================================================================
# 7. FUNÇÃO PRINCIPAL
# =============================================================================


def main() -> None:
    st.set_page_config(page_title=f"{APP_TITLE} – Momento Uau", layout="wide")
    render_styles()
    init_db()
    init_session_state()

    catalog = get_catalog_items()
    catalog_by_id = {int(item["id"]): item for item in catalog}
    recipes = get_recipes()
    history = get_kits_history()

    st.session_state.kit_items = {
        int(k): int(v)
        for k, v in st.session_state.kit_items.items()
        if int(k) in catalog_by_id and int(v) > 0
    }

    show_flash()
    render_header(len(history))
    page = render_nav(len(history))
    render_last_summary()

    if page == "montagem":
        page_assembly(catalog, catalog_by_id, recipes)
    elif page == "catalogo":
        page_catalog(catalog)
    else:
        render_history(history)


if __name__ == "__main__":
    main()