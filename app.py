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
    adjust_inventory_stock,
    delete_catalog_item,
    delete_kit,
    delete_recipe,
    format_address_summary,
    get_catalog_items,
    get_inventory_items,
    get_inventory_movements,
    get_kits_history,
    get_recipes,
    init_db,
    remove_from_inventory,
    save_kit,
    save_recipe,
    update_catalog_item,
    update_recipe,
)

# =============================================================================
# 1. CONFIGURAÇÃO E CONSTANTES
# =============================================================================

APP_TITLE = "Kits Momento Uau"
APP_SUBTITLE = "Montagem visual, controle de inventário e cálculo automático"

STATUS_OPTIONS = ["Em montagem", "Conferido", "Finalizado"]
STATUS_CHIP = {"Em montagem": "chip-warn", "Conferido": "chip-blue", "Finalizado": "chip-ok"}
MAX_BATCH_KITS = 10
ALLOWED_IMAGE_TYPES = ["png", "jpg", "jpeg", "webp"]

PAGES = {
    "montagem": "Montagem do kit",
    "inventario": "Inventário",
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


def md_money(value: float) -> str:
    """Como format_brl, mas escapa o '$' para o Markdown não ler como fórmula."""
    return format_brl(value).replace("$", "\\$")


def esc(value: Any) -> str:
    """Escapa texto do usuário antes de inseri-lo em HTML."""
    return html.escape(str(value or ""))


def render_html(raw: str) -> None:
    """Renderiza HTML no Streamlit sem virar bloco de código.

    O Markdown trata linhas com 4+ espaços de indentação (ou separadas por
    linhas em branco) como código, então removemos ambos antes de enviar.
    """
    compact = "\n".join(line.strip() for line in raw.splitlines() if line.strip())
    st.markdown(compact, unsafe_allow_html=True)


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


def inventory_status_label(stock_qty: int, min_stock: int) -> tuple[str, str]:
    if stock_qty <= 0:
        return "Sem estoque", "chip-danger"
    if stock_qty <= min_stock:
        return "Estoque baixo", "chip-warn"
    return "Disponível", "chip-ok"


def inventory_rows(items: list[Item], money: Callable[[float], Any]) -> list[dict[str, Any]]:
    rows = []
    for it in items:
        qty = int(it.get("stock_qty", 0))
        min_q = int(it.get("min_stock", 0))
        unit = float(it.get("unit_price", 0.0))
        status_text, _ = inventory_status_label(qty, min_q)
        rows.append(
            {
                "Item": it["name"],
                "Categoria": it.get("category") or "Geral",
                "Quantidade em estoque": qty,
                "Estoque mínimo": min_q,
                "Situação": status_text,
                "Valor unitário": money(unit),
                "Valor total em estoque": money(round(qty * unit, 2)),
            }
        )
    return rows


def build_inventory_excel_bytes(items: list[Item]) -> bytes:
    df = pd.DataFrame(inventory_rows(items, float))
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="Inventario_Brindes")
        ws = writer.sheets["Inventario_Brindes"]
        if not df.empty:
            unit_col = list(df.columns).index("Valor unitário") + 1
            tot_col = list(df.columns).index("Valor total em estoque") + 1
            for r_idx in range(2, ws.max_row + 1):
                ws.cell(row=r_idx, column=unit_col).number_format = "R$ #,##0.00"
                ws.cell(row=r_idx, column=tot_col).number_format = "R$ #,##0.00"
            for idx, col in enumerate(df.columns, start=1):
                longest = max(len(str(col)), *(len(str(v)) for v in df[col]))
                ws.column_dimensions[get_column_letter(idx)].width = min(longest + 3, 55)
    return output.getvalue()


def build_inventory_csv_bytes(items: list[Item]) -> bytes:
    df = pd.DataFrame(inventory_rows(items, format_brl))
    return df.to_csv(index=False, sep=";").encode("utf-8-sig")


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


@st.cache_resource(show_spinner=False)
def _init_db_once() -> bool:
    """Cria/verifica as tabelas só uma vez por deploy (evita consultas a cada clique)."""
    init_db()
    return True


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
    ss = st.session_state
    for name in FORM_FIELDS:
        ss[f"form_{name}"] = str(values.get(name) or "")
    if ss.form_status not in STATUS_OPTIONS:
        ss.form_status = STATUS_OPTIONS[0]
    ss.form_nonce = int(ss.form_nonce) + 1


def current_form() -> dict[str, str]:
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

    count_kits = len(filled_rows)
    for line in totals.lines:
        cat_item = line["item"]
        needed_total = int(line["quantity"]) * count_kits
        avail = int(cat_item.get("stock_qty", 0))
        if avail < needed_total:
            notify(
                "warning",
                f"Estoque insuficiente de '{cat_item['name']}' para {count_kits} kits "
                f"(disponível: {avail}, necessário: {needed_total}).",
            )
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
    notify("success", f"{len(saved_summaries)} kits gerados e baixados do inventário.")


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


def inv_quick_delta(item_id: int, delta: int, reason: str) -> None:
    ok, msg = adjust_inventory_stock(item_id=item_id, delta=delta, reason=reason)
    notify("success" if ok else "warning", msg)


def inv_remove_direct(item_id: int) -> None:
    ok, msg = remove_from_inventory(item_id=item_id, active_kit_items=st.session_state.kit_items)
    notify("success" if ok else "warning", msg)


# =============================================================================
# 4. ESTILOS (DASHBOARD MODERNO COM ACENTUAÇÕES EM AZUL)
# =============================================================================

CSS = """
<style>
    :root {
        --uau-blue: #2563eb;
        --uau-blue-dark: #1d4ed8;
        --uau-blue-light: #3b82f6;
        --uau-sky: #60a5fa;
        --uau-blue-soft: rgba(37, 99, 235, 0.12);
        --uau-blue-border: rgba(37, 99, 235, 0.34);
        --uau-navy: #0f172a;
        --uau-navy-2: #1e293b;
        --uau-border: rgba(148, 163, 184, 0.26);
    }

    .block-container { padding-top: 1.4rem; padding-bottom: 3rem; max-width: 1360px; }

    /* Cards (containers com borda) e inputs mais arredondados */
    div[data-testid="stVerticalBlockBorderWrapper"] { border-radius: 18px; }
    div[data-baseweb="input"], div[data-baseweb="select"] > div, div[data-baseweb="textarea"] {
        border-radius: 10px !important;
    }

    /* ---------- Cabeçalho escuro com detalhes em azul ---------- */
    .top-bar {
        display: flex; justify-content: space-between; align-items: center; flex-wrap: wrap; gap: 14px;
        background: linear-gradient(135deg, #0f172a 0%, #1e293b 100%);
        color: #ffffff;
        border-radius: 22px; padding: 16px 24px; margin-bottom: 14px;
        box-shadow: 0 8px 24px rgba(15, 23, 42, 0.20);
    }
    .brand-wrap { display: flex; align-items: center; gap: 14px; }
    .brand-badge {
        width: 44px; height: 44px; border-radius: 14px;
        background: linear-gradient(135deg, #2563eb 0%, #3b82f6 100%);
        color: #ffffff; font-weight: 800; font-size: 0.92rem; letter-spacing: 0.04em;
        display: flex; align-items: center; justify-content: center;
        box-shadow: 0 6px 16px rgba(37, 99, 235, 0.45);
    }
    .hero-title { font-size: 1.38rem; font-weight: 800; line-height: 1.2; letter-spacing: -0.01em; color: #ffffff; }
    .hero-sub { margin-top: 2px; font-size: 0.84rem; color: rgba(255, 255, 255, 0.66); }
    .header-pills { display: flex; align-items: center; gap: 8px; flex-wrap: wrap; }
    .hero-pill {
        background: rgba(96, 165, 250, 0.14); border: 1px solid rgba(96, 165, 250, 0.36);
        color: #93c5fd; border-radius: 999px;
        padding: 6px 14px; font-size: 0.8rem; font-weight: 700; white-space: nowrap;
    }
    .hero-pill-solid {
        background: var(--uau-blue); color: #ffffff; border-radius: 999px;
        padding: 6px 14px; font-size: 0.8rem; font-weight: 700; white-space: nowrap;
    }

    /* ---------- Navegação em pílulas ---------- */
    div[role="radiogroup"] {
        display: inline-flex !important; flex-wrap: wrap; gap: 6px;
        background: var(--uau-navy);
        border-radius: 999px; padding: 6px; margin-bottom: 10px;
    }
    div[role="radiogroup"] > label {
        background: transparent; border: 1px solid transparent;
        border-radius: 999px; padding: 6px 18px; cursor: pointer; margin: 0;
        transition: all 0.16s ease;
    }
    div[role="radiogroup"] > label > div:first-child { display: none; }
    div[role="radiogroup"] > label * { color: rgba(255, 255, 255, 0.80) !important; }
    div[role="radiogroup"] > label:hover { background: rgba(255, 255, 255, 0.08); }
    div[role="radiogroup"] > label:has(input:checked) {
        background: linear-gradient(135deg, #2563eb 0%, #3b82f6 100%);
        box-shadow: 0 4px 12px rgba(37, 99, 235, 0.40);
    }
    div[role="radiogroup"] > label:has(input:checked) * { color: #ffffff !important; font-weight: 700; }

    /* ---------- KPIs ---------- */
    .kpi-grid {
        display: grid; grid-template-columns: repeat(4, minmax(0, 1fr)); gap: 12px;
        margin-bottom: 16px;
    }
    @media (max-width: 900px) { .kpi-grid { grid-template-columns: repeat(2, minmax(0, 1fr)); } }
    .kpi-card {
        background: var(--secondary-background-color);
        border: 1px solid var(--uau-border);
        border-radius: 20px; padding: 14px 18px;
        transition: border-color 0.15s ease;
    }
    .kpi-card:hover { border-color: var(--uau-blue-border); }
    .kpi-card-accent {
        background: linear-gradient(135deg, #0f172a 0%, #1e3a8a 100%);
        border: 1px solid #1e3a8a; color: #ffffff;
    }
    .kpi-card-accent .kpi-num { color: var(--uau-sky); }
    .kpi-label { font-size: 0.75rem; font-weight: 700; text-transform: uppercase; letter-spacing: 0.05em; opacity: 0.68; }
    .kpi-num { font-size: 1.48rem; font-weight: 800; line-height: 1.2; margin-top: 4px; }
    .kpi-sub { font-size: 0.78rem; opacity: 0.68; margin-top: 2px; }

    /* ---------- Resumo do kit ---------- */
    .summary-header {
        display: flex; justify-content: space-between; align-items: center;
        padding-bottom: 10px; margin-bottom: 12px;
        border-bottom: 1px solid var(--uau-border);
    }
    .summary-title { font-size: 1.05rem; font-weight: 800; }

    /* ---------- Card de resumo de valor (estilo "Order Summary") ---------- */
    .order-card {
        background: linear-gradient(135deg, #0f172a 0%, #1e3a8a 100%);
        color: #ffffff; border-radius: 20px; padding: 18px 22px; margin: 10px 0 12px;
        box-shadow: 0 10px 26px rgba(30, 58, 138, 0.28);
    }
    .order-title { font-weight: 800; font-size: 1rem; margin-bottom: 10px; }
    .order-row {
        display: flex; justify-content: space-between; align-items: center;
        font-size: 0.88rem; padding: 4px 0; color: rgba(255, 255, 255, 0.72);
    }
    .order-row b { color: #ffffff; font-weight: 700; }
    .order-divider { border-top: 1px solid rgba(255, 255, 255, 0.16); margin: 10px 0; }
    .order-total { display: flex; justify-content: space-between; align-items: baseline; font-weight: 700; }
    .order-total-value { font-size: 1.95rem; font-weight: 800; color: var(--uau-sky); }
    .order-hint { margin-top: 6px; font-size: 0.76rem; color: rgba(255, 255, 255, 0.58); }

    /* ---------- Ícones, chips e badges ---------- */
    .icon-tile {
        width: var(--s, 54px); height: var(--s, 54px);
        display: flex; align-items: center; justify-content: center;
        font-size: var(--fs, 19px); font-weight: 800; color: var(--uau-blue);
        overflow: hidden; border-radius: 14px;
        background: var(--uau-blue-soft);
        border: 1px solid var(--uau-blue-border);
    }
    .chip {
        display: inline-block; font-size: 0.72rem; font-weight: 700;
        padding: 2px 10px; border-radius: 999px;
        background: var(--uau-blue-soft); color: var(--uau-blue-light);
    }
    .chip-blue   { background: rgba(37, 99, 235, 0.15); color: #2563eb; }
    .chip-warn   { background: rgba(245, 158, 11, 0.16); color: #d97706; }
    .chip-ok     { background: rgba(16, 185, 129, 0.16); color: #059669; }
    .chip-danger { background: rgba(239, 68, 68, 0.16); color: #dc2626; }

    .qty-badge {
        display: inline-block; background: var(--uau-blue); color: #ffffff;
        font-size: 0.76rem; font-weight: 700; padding: 2px 10px; border-radius: 999px;
    }
    .price { font-size: 1.02rem; font-weight: 800; color: var(--uau-blue-light); }
    .kit-value { font-size: 1.45rem; font-weight: 800; text-align: right; line-height: 1.2; color: var(--uau-blue-light); }
    .kit-meta { text-align: right; opacity: 0.65; font-size: 0.82rem; }
    .box-title { font-weight: 800; font-size: 0.98rem; }
    .tile-name { text-align: center; font-weight: 700; font-size: 0.88rem; margin-top: 4px; }
    .tile-qty { text-align: center; margin: 4px 0 2px; }
    .tile-sub { text-align: center; opacity: 0.72; font-size: 0.8rem; font-weight: 600; margin-bottom: 6px; }

    /* ---------- Botões primários em azul ---------- */
    button[kind="primary"],
    [data-testid^="stBaseButton-primary"] {
        background: linear-gradient(135deg, #2563eb 0%, #3b82f6 100%) !important;
        border: 1px solid #2563eb !important;
        color: #ffffff !important;
        border-radius: 12px !important;
        font-weight: 700 !important;
        box-shadow: 0 4px 12px rgba(37, 99, 235, 0.26);
    }
    button[kind="primary"]:hover,
    [data-testid^="stBaseButton-primary"]:hover {
        background: linear-gradient(135deg, #1d4ed8 0%, #2563eb 100%) !important;
        border-color: #1d4ed8 !important;
    }
    button[kind="secondary"],
    [data-testid^="stBaseButton-secondary"] { border-radius: 12px !important; }
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


def render_header(history_count: int, total_stock: int) -> None:
    in_block = history_count % MAX_BATCH_KITS
    render_html(
        f"""
        <div class="top-bar">
            <div class="brand-wrap">
                <div class="brand-badge">UAU</div>
                <div>
                    <div class="hero-title">{APP_TITLE}</div>
                    <div class="hero-sub">{APP_SUBTITLE}</div>
                </div>
            </div>
            <div class="header-pills">
                <span class="hero-pill">Estoque total: {total_stock} un.</span>
                <span class="hero-pill-solid">Bloco: {in_block}/{MAX_BATCH_KITS} kits</span>
            </div>
        </div>
        """
    )


def render_nav(history_count: int, total_stock: int) -> str:
    labels = dict(PAGES)
    labels["inventario"] = f"{PAGES['inventario']} ({total_stock})"
    labels["historico"] = f"{PAGES['historico']} ({history_count})"
    return st.radio(
        "Navegação",
        options=list(labels),
        format_func=labels.get,
        key="nav_page",
        horizontal=True,
        label_visibility="collapsed",
    )


def render_top_kpis(
    totals: KitTotals,
    total_stock: int,
    low_stock_count: int,
    history: list[Item],
) -> None:
    accumulated_val = sum(float(k.get("total_value", 0.0)) for k in history)
    alert_text = f"{low_stock_count} item(ns) em alerta" if low_stock_count else "Níveis normais"
    average = format_brl(accumulated_val / len(history)) if history else "R$ 0,00"
    render_html(
        f"""
        <div class="kpi-grid">
            <div class="kpi-card kpi-card-accent">
                <div class="kpi-label">Kit em montagem</div>
                <div class="kpi-num">{format_brl(totals.value)}</div>
                <div class="kpi-sub">{totals.quantity} item(ns) na caixa</div>
            </div>
            <div class="kpi-card">
                <div class="kpi-label">Kits finalizados</div>
                <div class="kpi-num">{len(history)}</div>
                <div class="kpi-sub">Total acumulado: {format_brl(accumulated_val)}</div>
            </div>
            <div class="kpi-card">
                <div class="kpi-label">Peças em estoque</div>
                <div class="kpi-num">{total_stock} un.</div>
                <div class="kpi-sub">{alert_text}</div>
            </div>
            <div class="kpi-card">
                <div class="kpi-label">Média por kit</div>
                <div class="kpi-num">{average}</div>
                <div class="kpi-sub">Baseado no histórico</div>
            </div>
        </div>
        """
    )


def render_last_summary() -> None:
    batch = st.session_state.get("last_batch_summaries") or []
    if batch:
        with st.container(border=True):
            total_lote = sum(float(s.get("total_value", 0.0)) for s in batch)
            st.success(
                f"**{len(batch)} kits gerados e baixados do inventário.** "
                f"Valor total do lote: {md_money(total_lote)}."
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
            st.success(f"**{summary['kit_code']} atualizado.** Novo valor: {md_money(summary['total_value'])}.")
        else:
            st.success(
                f"**{summary['kit_code']} finalizado e baixado do inventário.** "
                f"Registre {md_money(summary['total_value'])} no formulário."
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
    with st.container(border=True):
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
                                f"{item['name']} ({md_money(item['unit_price'])})",
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
    with st.container(border=True):
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
                stock_qty = int(item.get("stock_qty", 0))
                min_stock = int(item.get("min_stock", 0))
                _, stock_chip = inventory_status_label(stock_qty, min_stock)

                with col, st.container(border=True):
                    c_vis, c_info = st.columns([0.9, 2.1])
                    with c_vis:
                        render_item_visual(item, size=58)
                    with c_info:
                        st.markdown(
                            f"<span class='chip chip-blue'>{esc(item.get('category') or 'Geral')}</span> "
                            f"<span class='chip {stock_chip}'>Estoque: {stock_qty}</span>",
                            unsafe_allow_html=True,
                        )
                        st.markdown(f"**{esc(item['name'])}**")
                        st.markdown(
                            f"<span class='price'>{format_brl(item['unit_price'])}</span>",
                            unsafe_allow_html=True,
                        )
                    st.button(
                        f"Selecionado ({in_kit} no kit)" if in_kit else "Adicionar ao kit",
                        key=f"add_catalog_item_{item_id}",
                        type="primary" if in_kit else "secondary",
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
    people = max(1, int(recipients_count))
    lote_total = round(totals.value * people, 2)
    hint = (
        "Valor por pessoa para preencher no formulário"
        if people > 1
        else "Valor para preencher no formulário"
    )
    render_html(
        f"""
        <div class="order-card">
            <div class="order-title">Resumo do valor</div>
            <div class="order-row"><span>Itens por kit</span><b>{totals.quantity} {plural}</b></div>
            <div class="order-row"><span>Valor por kit</span><b>{format_brl(totals.value)}</b></div>
            <div class="order-row"><span>Destinatários</span><b>{people}</b></div>
            <div class="order-divider"></div>
            <div class="order-total">
                <span>Total</span>
                <span class="order-total-value">{format_brl(lote_total)}</span>
            </div>
            <div class="order-hint">{hint}</div>
        </div>
        """
    )


def render_gift_box(totals: KitTotals, multiplier: int = 1) -> None:
    with st.container(border=True):
        st.markdown(
            "<div class='summary-header'>"
            "<span class='summary-title'>Resumo do kit</span>"
            f"<span class='chip chip-blue'>{totals.quantity} un.</span>"
            "</div>",
            unsafe_allow_html=True,
        )
        if not totals.lines:
            st.caption("Nenhum item selecionado. Clique em um brinde ao lado ou aplique uma receita.")
            return

        low_warnings = []
        for line in totals.lines:
            avail = int(line["item"].get("stock_qty", 0))
            needed = int(line["quantity"]) * max(1, multiplier)
            if avail < needed:
                low_warnings.append(f"{line['item']['name']} (disponível: {avail}, necessário: {needed})")
        if low_warnings:
            st.warning("Estoque insuficiente para finalizar: " + ", ".join(low_warnings))

        for row in chunked(totals.lines, 3):
            cols = st.columns(3, gap="small")
            for col, line in zip(cols, row):
                item_id = line["item_id"]
                with col, st.container(border=True):
                    render_item_visual(line["item"], size=46)
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
        st.markdown("**Dados do destinatário**")
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
        synced_rows = _sync_batch_from_widgets()
        filled = len([r for r in synced_rows if str(r.get("client_name") or "").strip()])
        render_gift_box(totals, multiplier=max(1, filled))
        active_people = render_batch_form()
        render_total_card(totals, recipients_count=active_people)
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

    render_gift_box(totals, multiplier=1)
    render_kit_form()
    render_total_card(totals, recipients_count=1)

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


# ---- Inventário -------------------------------------------------------------


def page_inventory(inv_items: list[Item], all_catalog: list[Item]) -> None:
    st.markdown("#### Controle de inventário")
    st.caption(
        "A cada kit finalizado (individual ou em lote), o estoque é reduzido automaticamente. "
        "Se um kit for editado ou excluído no histórico, o saldo é recalculado."
    )

    total_units = sum(int(it.get("stock_qty", 0)) for it in inv_items)
    total_stock_value = sum(
        int(it.get("stock_qty", 0)) * float(it.get("unit_price", 0.0)) for it in inv_items
    )
    low_count = sum(
        1 for it in inv_items if int(it.get("stock_qty", 0)) <= int(it.get("min_stock", 0))
    )

    m1, m2, m3 = st.columns(3)
    m1.metric("Brindes no inventário", len(inv_items))
    m2.metric("Unidades disponíveis", total_units)
    m3.metric(
        "Valor total em estoque",
        format_brl(total_stock_value),
        delta=f"{low_count} em alerta" if low_count else None,
        delta_color="inverse" if low_count else "off",
    )

    st.write("")
    exp1, exp2, _ = st.columns([1.1, 1.1, 2.2])
    exp1.download_button(
        "Exportar inventário Excel (.xlsx)",
        data=build_inventory_excel_bytes(inv_items),
        file_name="inventario_brindes_momento_uau.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        **STRETCH,
    )
    exp2.download_button(
        "Exportar inventário CSV (.csv)",
        data=build_inventory_csv_bytes(inv_items),
        file_name="inventario_brindes_momento_uau.csv",
        mime="text/csv",
        **STRETCH,
    )

    with st.expander("Adicionar novo item ao inventário", expanded=False):
        with st.form("inv_add_item_form", clear_on_submit=True):
            c1, c2, c3, c4, c5 = st.columns([2.0, 1.1, 1.0, 1.0, 1.3])
            n_name = c1.text_input("Nome do item *", placeholder="Ex.: Squeeze")
            n_price = c2.number_input("Valor unit. (R$) *", min_value=0.0, value=25.0, step=1.0, format="%.2f")
            n_qty = c3.number_input("Qtd. inicial *", min_value=0, value=30, step=1)
            n_min = c4.number_input("Estoque mín.", min_value=0, value=5, step=1)
            n_cat = c5.text_input("Categoria", placeholder="Ex.: Brindes")
            n_icon = st.text_input("Sigla (opcional)", placeholder="Ex.: SQZ", max_chars=12)

            if st.form_submit_button("Adicionar ao inventário", type="primary"):
                ok, msg = add_catalog_item(
                    name=n_name,
                    unit_price=n_price,
                    icon=n_icon,
                    category=n_cat,
                    stock_qty=int(n_qty),
                    min_stock=int(n_min),
                    in_inventory=1,
                )
                notify("success" if ok else "warning", msg)
                st.rerun()

        hidden_items = [it for it in all_catalog if int(it.get("in_inventory", 1)) == 0]
        if hidden_items:
            st.markdown("**Reativar item existente do catálogo**")
            r1, r2, r3 = st.columns([2.0, 1.0, 1.0])
            sel_name = r1.selectbox(
                "Item",
                options=[it["name"] for it in hidden_items],
                key="inv_reactivate_select",
            )
            sel_qty = r2.number_input("Quantidade", min_value=1, value=10, step=1, key="inv_reactivate_qty")
            if r3.button("Reativar", key="inv_reactivate_btn", **STRETCH):
                target = next(it for it in hidden_items if it["name"] == sel_name)
                ok, msg = adjust_inventory_stock(
                    item_id=int(target["id"]),
                    exact_qty=int(sel_qty),
                    reason="Reativado no inventario",
                )
                notify("success" if ok else "warning", msg)
                st.rerun()

    st.divider()
    f1, f2 = st.columns([1.8, 1.2])
    search_inv = f1.text_input("Buscar no inventário", placeholder="Nome ou categoria...", key="inv_search")
    filter_inv = f2.selectbox(
        "Filtrar situação",
        ["Todos", "Disponível", "Estoque baixo", "Sem estoque"],
        key="inv_filter_status",
    )

    term = search_inv.strip().lower()
    filtered_inv = []
    for it in inv_items:
        qty = int(it.get("stock_qty", 0))
        min_q = int(it.get("min_stock", 0))
        st_lbl, _ = inventory_status_label(qty, min_q)
        if filter_inv != "Todos" and st_lbl != filter_inv:
            continue
        if term and term not in it["name"].lower() and term not in str(it.get("category") or "").lower():
            continue
        filtered_inv.append(it)

    with st.expander("Ver inventário em tabela"):
        st.dataframe(pd.DataFrame(inventory_rows(filtered_inv, format_brl)), hide_index=True, **STRETCH)

    if not filtered_inv:
        st.info("Nenhum item encontrado no inventário para o filtro atual.")
    else:
        for it in filtered_inv:
            iid = int(it["id"])
            qty = int(it.get("stock_qty", 0))
            min_q = int(it.get("min_stock", 0))
            unit = float(it.get("unit_price", 0.0))
            st_lbl, chip_cls = inventory_status_label(qty, min_q)

            with st.container(border=True):
                c_vis, c_info, c_qty, c_actions = st.columns([0.65, 2.2, 1.3, 1.6], gap="medium")
                with c_vis:
                    render_item_visual(it, size=50)
                with c_info:
                    st.markdown(
                        f"**{esc(it['name'])}** "
                        f"<span class='chip chip-blue'>{esc(it.get('category') or 'Geral')}</span> "
                        f"<span class='chip {chip_cls}'>{st_lbl}</span>",
                        unsafe_allow_html=True,
                    )
                    st.caption(
                        f"Unitário: {md_money(unit)} · Total em estoque: {md_money(qty * unit)} · Mínimo: {min_q} un."
                    )
                with c_qty:
                    st.markdown(
                        f"<div style='font-size:1.35rem;font-weight:800;'>{qty} un.</div>",
                        unsafe_allow_html=True,
                    )
                    q1, q2 = st.columns(2)
                    q1.button(
                        "-1",
                        key=f"inv_dec_{iid}",
                        disabled=qty <= 0,
                        on_click=inv_quick_delta,
                        args=(iid, -1, "Saida rapida (-1)"),
                        **STRETCH,
                    )
                    q2.button(
                        "+1",
                        key=f"inv_inc_{iid}",
                        on_click=inv_quick_delta,
                        args=(iid, 1, "Entrada rapida (+1)"),
                        **STRETCH,
                    )
                with c_actions:
                    with st.popover("Alterar / Repor", **STRETCH):
                        with st.form(f"inv_edit_form_{iid}"):
                            ed_name = st.text_input("Nome", value=it["name"], key=f"inv_ed_name_{iid}")
                            ed_price = st.number_input(
                                "Valor unitário (R$)", min_value=0.0, value=unit, step=1.0, format="%.2f", key=f"inv_ed_price_{iid}"
                            )
                            ed_cat = st.text_input(
                                "Categoria", value=it.get("category") or "Geral", key=f"inv_ed_cat_{iid}"
                            )
                            ed_qty = st.number_input(
                                "Quantidade exata em estoque", min_value=0, value=qty, step=1, key=f"inv_ed_qty_{iid}"
                            )
                            ed_min = st.number_input(
                                "Estoque mínimo para alerta", min_value=0, value=min_q, step=1, key=f"inv_ed_min_{iid}"
                            )
                            if st.form_submit_button("Salvar alterações", type="primary", **STRETCH):
                                ok, msg = update_catalog_item(
                                    item_id=iid,
                                    name=ed_name,
                                    unit_price=ed_price,
                                    icon=it.get("icon") or "[ITEM]",
                                    category=ed_cat,
                                    image_path=it.get("image_path") or "",
                                    stock_qty=int(ed_qty),
                                    min_stock=int(ed_min),
                                    in_inventory=1,
                                )
                                notify("success" if ok else "warning", msg)
                                st.rerun()

                    st.button(
                        "Remover",
                        key=f"inv_del_{iid}",
                        on_click=inv_remove_direct,
                        args=(iid,),
                        **STRETCH,
                    )

    movements = get_inventory_movements(limit=30)
    if movements:
        with st.expander("Histórico recente de movimentações de estoque"):
            mov_df = pd.DataFrame(
                [
                    {
                        "Data": m["created_at"],
                        "Item": m["item_name"],
                        "Movimento": f"+{m['delta']}" if int(m["delta"]) > 0 else str(m["delta"]),
                        "Saldo após": m["balance_after"],
                        "Motivo": m["reason"],
                    }
                    for m in movements
                ]
            )
            st.dataframe(mov_df, hide_index=True, **STRETCH)


# ---- Catálogo: cadastro e edição -------------------------------------------


def render_new_item_form() -> None:
    st.markdown("#### Cadastrar novo brinde")
    with st.form("add_catalog_form", clear_on_submit=True):
        c1, c2, c3, c4, c5 = st.columns([1.8, 1.1, 0.9, 1.0, 1.3])
        name = c1.text_input("Nome do brinde *", placeholder="Ex.: Mochila")
        price = c2.number_input("Valor unitário (R$) *", min_value=0.0, value=20.0, step=1.0, format="%.2f")
        stock = c3.number_input("Estoque inicial", min_value=0, value=20, step=1)
        icon = c4.text_input("Sigla", placeholder="Ex.: MOC", max_chars=12)
        category = c5.text_input("Categoria", placeholder="Ex.: Acessórios")
        image = st.file_uploader("Imagem do brinde (opcional)", type=ALLOWED_IMAGE_TYPES, key="new_item_uploader")

        if st.form_submit_button("Cadastrar brinde", type="primary"):
            image_path = save_uploaded_image(image) if image else ""
            ok, msg = add_catalog_item(
                name=name,
                unit_price=price,
                icon=icon,
                category=category,
                image_path=image_path,
                stock_qty=int(stock),
                in_inventory=1,
            )
            notify("success" if ok else "warning", msg)
            st.rerun()


def render_edit_item(item: Item) -> None:
    item_id = int(item["id"])
    stock_qty = int(item.get("stock_qty", 0))
    label = (
        f"{item['name']}  ·  {md_money(item['unit_price'])}  ·  "
        f"{item.get('category') or 'Geral'}  ·  Estoque: {stock_qty} un."
    )
    with st.expander(label):
        c_img, c_form = st.columns([1.0, 3.2], gap="medium")
        with c_img:
            render_item_visual(item, size=92)

        with c_form:
            with st.form(f"edit_catalog_form_{item_id}"):
                f1, f2, f3, f4, f5 = st.columns([1.8, 1.1, 0.9, 1.0, 1.3])
                name = f1.text_input("Nome", value=item["name"], key=f"ed_name_{item_id}")
                price = f2.number_input(
                    "Valor unitário (R$)", min_value=0.0, value=float(item["unit_price"]),
                    step=1.0, format="%.2f", key=f"ed_price_{item_id}",
                )
                stock = f3.number_input(
                    "Estoque (un.)", min_value=0, value=stock_qty, step=1, key=f"ed_stock_{item_id}",
                )
                icon = f4.text_input(
                    "Sigla", value=icon_for_input(item), max_chars=12, key=f"ed_icon_{item_id}"
                )
                category = f5.text_input(
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
                        item_id=item_id,
                        name=name,
                        unit_price=price,
                        icon=icon,
                        category=category,
                        image_path=image_path,
                        stock_qty=int(stock),
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
    col_left, col_right = st.columns([1.22, 1.0], gap="large")
    with col_left:
        render_recipes(recipes, catalog, catalog_by_id)
        st.write("")
        render_catalog_cards(catalog)
    with col_right:
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
    st.set_page_config(page_title=APP_TITLE, layout="wide")
    render_styles()
    _init_db_once()
    init_session_state()

    catalog = get_catalog_items()
    catalog_by_id = {int(item["id"]): item for item in catalog}
    inv_items = get_inventory_items()
    recipes = get_recipes()
    history = get_kits_history()

    st.session_state.kit_items = {
        int(k): int(v)
        for k, v in st.session_state.kit_items.items()
        if int(k) in catalog_by_id and int(v) > 0
    }

    total_stock = sum(int(it.get("stock_qty", 0)) for it in inv_items)
    low_stock_count = sum(
        1 for it in inv_items if int(it.get("stock_qty", 0)) <= int(it.get("min_stock", 0))
    )
    current_totals = compute_totals(st.session_state.kit_items, catalog_by_id)

    show_flash()
    render_header(len(history), total_stock)
    page = render_nav(len(history), total_stock)
    render_top_kpis(current_totals, total_stock, low_stock_count, history)
    render_last_summary()

    if page == "montagem":
        page_assembly(catalog, catalog_by_id, recipes)
    elif page == "inventario":
        page_inventory(inv_items, catalog)
    elif page == "catalogo":
        page_catalog(catalog)
    else:
        render_history(history)


if __name__ == "__main__":
    main()