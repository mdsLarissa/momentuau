import base64
import io
import os
import uuid
from datetime import datetime
from typing import Any, Dict, List

import pandas as pd
import streamlit as st
from PIL import Image

from database import (
    DEFAULT_DB_PATH,
    add_item,
    adjust_item_stock,
    clear_all_kits,
    create_kit,
    delete_item,
    delete_kit,
    delete_recipe,
    generate_code_from_name,
    init_db,
    is_postgres_enabled,
    list_items,
    list_kits,
    list_recipes,
    list_stock_movements,
    save_recipe,
    update_item,
    update_kit,
)

ASSETS_DIR = "assets"
os.makedirs(ASSETS_DIR, exist_ok=True)


# ==============================================================================
# UTILITÁRIOS DE FORMATAÇÃO E IMAGEM
# ==============================================================================

def format_brl(value: float) -> str:
    formatted = f"{float(value):,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
    return f"R$ {formatted}"


def save_uploaded_image(uploaded_file) -> str:
    if uploaded_file is None:
        return ""
    os.makedirs(ASSETS_DIR, exist_ok=True)
    ext = os.path.splitext(uploaded_file.name)[1].lower() or ".png"
    filename = f"{uuid.uuid4().hex[:12]}{ext}"
    dest_path = os.path.join(ASSETS_DIR, filename)
    img = Image.open(uploaded_file)
    img.thumbnail((360, 360))
    img.save(dest_path)
    return dest_path


def get_image_base64(image_path: str) -> str:
    if not image_path or not os.path.exists(image_path):
        return ""
    try:
        with open(image_path, "rb") as f:
            raw = f.read()
        ext = os.path.splitext(image_path)[1].lower().replace(".", "") or "png"
        if ext == "jpg":
            ext = "jpeg"
        b64 = base64.b64encode(raw).decode("utf-8")
        return f"data:image/{ext};base64,{b64}"
    except Exception:
        return ""


def format_items_summary(kit_items: List[Dict[str, Any]]) -> str:
    parts = [f"{int(i['quantity'])}x {i['item_name']}" for i in kit_items if int(i["quantity"]) > 0]
    return ", ".join(parts)


# ==============================================================================
# EXPORTAÇÃO PADRONIZADA (4 COLUNAS) E INVENTÁRIO
# ==============================================================================

def build_kits_dataframe(kits: List[Dict[str, Any]]) -> pd.DataFrame:
    rows = []
    for k in kits:
        rows.append({
            "Itens e quantidades": format_items_summary(k.get("items", [])),
            "Nome": k.get("recipient_name", ""),
            "Endereço": k.get("recipient_address", ""),
            "Valor": format_brl(float(k.get("total_value", 0.0))),
        })
    if not rows:
        return pd.DataFrame(columns=["Itens e quantidades", "Nome", "Endereço", "Valor"])
    return pd.DataFrame(rows)


def dataframe_to_excel_bytes(df: pd.DataFrame, sheet_name: str = "Kits") -> bytes:
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name=sheet_name)
        ws = writer.sheets[sheet_name]
        for col_idx, col_name in enumerate(df.columns, start=1):
            max_len = max(
                len(str(col_name)),
                *(len(str(val)) for val in df[col_name].fillna("")),
            ) if not df.empty else len(str(col_name))
            ws.column_dimensions[ws.cell(row=1, column=col_idx).column_letter].width = min(max_len + 5, 65)
    return output.getvalue()


def build_inventory_dataframe(items: List[Dict[str, Any]]) -> pd.DataFrame:
    rows = []
    for it in items:
        qty = int(it.get("stock_qty", 0))
        min_s = int(it.get("min_stock", 5))
        unit = float(it.get("unit_price", 0.0))
        status = "Esgotado" if qty == 0 else ("Estoque Baixo" if qty <= min_s else "Disponível")
        rows.append({
            "Código": it.get("code", ""),
            "Brinde": it.get("name", ""),
            "Estoque Atual": qty,
            "Estoque Mínimo": min_s,
            "Status": status,
            "Valor Unitário": format_brl(unit),
            "Valor Total em Estoque": format_brl(qty * unit),
        })
    if not rows:
        return pd.DataFrame(
            columns=["Código", "Brinde", "Estoque Atual", "Estoque Mínimo", "Status", "Valor Unitário", "Valor Total em Estoque"]
        )
    return pd.DataFrame(rows)


# ==============================================================================
# ESTILO VISUAL MODERNO COM ACENTOS EM AZUL (ADAPTATIVO CLARO / ESCURO)
# ==============================================================================

def inject_modern_blue_css() -> None:
    st.markdown(
        """
        <style>
        @import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700;800&display=swap');

        :root {
            --brand-blue: #2563EB;
            --brand-blue-hover: #1D4ED8;
            --brand-blue-soft: rgba(37, 99, 235, 0.12);
            --brand-blue-border: rgba(37, 99, 235, 0.35);
            --card-border: rgba(148, 163, 184, 0.22);
            --radius-xl: 20px;
            --radius-lg: 14px;
            --radius-pill: 999px;
        }

        html, body, [class*="css"] {
            font-family: 'Inter', -apple-system, BlinkMacSystemFont, sans-serif;
        }

        .block-container {
            padding-top: 1.5rem !important;
            padding-bottom: 3rem !important;
            max-width: 1320px;
        }

        /* Cabeçalho Hero SaaS */
        .uau-topbar {
            display: flex;
            align-items: center;
            justify-content: space-between;
            padding: 1rem 1.4rem;
            border-radius: var(--radius-xl);
            background: var(--secondary-background-color);
            border: 1px solid var(--card-border);
            margin-bottom: 1.1rem;
        }
        .uau-brand-group {
            display: flex;
            align-items: center;
            gap: 0.85rem;
        }
        .uau-logo-badge {
            width: 42px;
            height: 42px;
            border-radius: 12px;
            background: linear-gradient(135deg, #2563EB 0%, #3B82F6 100%);
            color: #FFFFFF;
            font-weight: 800;
            font-size: 1rem;
            display: flex;
            align-items: center;
            justify-content: center;
            box-shadow: 0 6px 16px rgba(37, 99, 235, 0.28);
        }
        .uau-title {
            font-size: 1.25rem;
            font-weight: 800;
            margin: 0;
            line-height: 1.2;
            letter-spacing: -0.02em;
        }
        .uau-subtitle {
            font-size: 0.82rem;
            opacity: 0.72;
            margin: 0.1rem 0 0 0;
        }
        .uau-status-pill {
            display: inline-flex;
            align-items: center;
            gap: 0.4rem;
            padding: 0.35rem 0.85rem;
            border-radius: var(--radius-pill);
            font-size: 0.78rem;
            font-weight: 600;
            background: var(--brand-blue-soft);
            color: #3B82F6;
            border: 1px solid var(--brand-blue-border);
        }

        /* Navegação estilo Pill Bar */
        div[data-testid="stRadio"] > div[role="radiogroup"] {
            background: var(--secondary-background-color);
            border: 1px solid var(--card-border);
            border-radius: var(--radius-pill);
            padding: 0.35rem;
            display: inline-flex;
            flex-wrap: wrap;
            gap: 0.3rem;
            margin-bottom: 0.75rem;
        }
        div[data-testid="stRadio"] > div[role="radiogroup"] > label {
            border-radius: var(--radius-pill) !important;
            padding: 0.42rem 1.05rem !important;
            margin: 0 !important;
            border: 1px solid transparent !important;
            transition: all 0.18s ease;
            cursor: pointer;
        }
        div[data-testid="stRadio"] > div[role="radiogroup"] > label > div:first-child {
            display: none !important;
        }
        div[data-testid="stRadio"] > div[role="radiogroup"] > label[data-checked="true"],
        div[data-testid="stRadio"] > div[role="radiogroup"] > label:has(input:checked) {
            background: var(--brand-blue) !important;
            color: #FFFFFF !important;
            box-shadow: 0 4px 12px rgba(37, 99, 235, 0.30);
        }
        div[data-testid="stRadio"] > div[role="radiogroup"] > label:has(input:checked) p {
            color: #FFFFFF !important;
            font-weight: 600 !important;
        }

        /* Cartões de KPIs arredondados */
        .kpi-card {
            background: var(--secondary-background-color);
            border: 1px solid var(--card-border);
            border-radius: var(--radius-xl);
            padding: 1.05rem 1.2rem;
            transition: border-color 0.2s ease, transform 0.15s ease;
        }
        .kpi-card:hover {
            border-color: var(--brand-blue-border);
        }
        .kpi-card-accent {
            background: linear-gradient(135deg, rgba(37, 99, 235, 0.14) 0%, rgba(59, 130, 246, 0.05) 100%);
            border: 1px solid var(--brand-blue-border);
        }
        .kpi-label {
            font-size: 0.78rem;
            font-weight: 600;
            opacity: 0.72;
            margin-bottom: 0.35rem;
            text-transform: uppercase;
            letter-spacing: 0.03em;
        }
        .kpi-value {
            font-size: 1.55rem;
            font-weight: 800;
            letter-spacing: -0.02em;
            line-height: 1.15;
        }
        .kpi-value-blue {
            color: #3B82F6;
        }
        .kpi-foot {
            font-size: 0.75rem;
            opacity: 0.68;
            margin-top: 0.3rem;
        }

        /* Cartões de seção estilo Checkout */
        .section-card {
            background: var(--secondary-background-color);
            border: 1px solid var(--card-border);
            border-radius: var(--radius-xl);
            padding: 1.25rem 1.35rem;
            margin-bottom: 1rem;
        }
        .section-card-title {
            font-size: 1rem;
            font-weight: 700;
            margin: 0 0 0.2rem 0;
            display: flex;
            align-items: center;
            gap: 0.45rem;
        }
        .section-card-desc {
            font-size: 0.8rem;
            opacity: 0.7;
            margin: 0 0 0.85rem 0;
        }

        /* Card de Brinde na Grade */
        .item-tile {
            border: 1px solid var(--card-border);
            border-radius: var(--radius-lg);
            padding: 0.85rem;
            background: var(--background-color);
            margin-bottom: 0.45rem;
            transition: all 0.18s ease;
        }
        .item-tile-active {
            border: 1.5px solid var(--brand-blue) !important;
            background: var(--brand-blue-soft) !important;
        }
        .item-code-pill {
            display: inline-block;
            padding: 0.18rem 0.55rem;
            border-radius: 8px;
            font-size: 0.72rem;
            font-weight: 700;
            background: var(--brand-blue-soft);
            color: #3B82F6;
            border: 1px solid var(--brand-blue-border);
        }
        .stock-badge-ok {
            font-size: 0.72rem;
            font-weight: 600;
            color: #10B981;
        }
        .stock-badge-low {
            font-size: 0.72rem;
            font-weight: 600;
            color: #F59E0B;
        }
        .stock-badge-out {
            font-size: 0.72rem;
            font-weight: 600;
            color: #EF4444;
        }

        /* Order Summary lateral */
        .summary-row {
            display: flex;
            justify-content: space-between;
            align-items: center;
            padding: 0.48rem 0;
            border-bottom: 1px dashed var(--card-border);
            font-size: 0.88rem;
        }
        .summary-total-box {
            margin-top: 0.85rem;
            padding: 0.9rem 1rem;
            border-radius: var(--radius-lg);
            background: var(--brand-blue-soft);
            border: 1px solid var(--brand-blue-border);
            display: flex;
            justify-content: space-between;
            align-items: center;
        }

        /* Botões primários em Azul */
        div.stButton > button[kind="primary"],
        div.stFormSubmitButton > button[kind="primary"] {
            background-color: var(--brand-blue) !important;
            border-color: var(--brand-blue) !important;
            color: #FFFFFF !important;
            border-radius: 12px !important;
            font-weight: 600 !important;
            transition: all 0.18s ease !important;
        }
        div.stButton > button[kind="primary"]:hover,
        div.stFormSubmitButton > button[kind="primary"]:hover {
            background-color: var(--brand-blue-hover) !important;
            border-color: var(--brand-blue-hover) !important;
            box-shadow: 0 6px 16px rgba(37, 99, 235, 0.32) !important;
        }
        div.stButton > button:not([kind="primary"]),
        div.stDownloadButton > button {
            border-radius: 12px !important;
            font-weight: 500 !important;
        }
        </style>
        """,
        unsafe_allow_html=True,
    )


# ==============================================================================
# ESTADO DE SESSÃO E HELPERS DE MONTAGEM
# ==============================================================================

def ensure_session_state(items: List[Dict[str, Any]]) -> None:
    if "kit_quantities" not in st.session_state:
        st.session_state["kit_quantities"] = {int(i["id"]): 0 for i in items}
    else:
        for i in items:
            st.session_state["kit_quantities"].setdefault(int(i["id"]), 0)

    st.session_state.setdefault("batch_count", 1)
    st.session_state.setdefault("editing_kit_id", None)
    st.session_state.setdefault("active_nav", "Montagem do Kit")


def reset_builder_form(items: List[Dict[str, Any]]) -> None:
    for i in items:
        iid = int(i["id"])
        st.session_state["kit_quantities"][iid] = 0
        st.session_state[f"qty_input_{iid}"] = 0
    for idx in range(10):
        st.session_state[f"rec_name_{idx}"] = ""
        st.session_state[f"rec_addr_{idx}"] = ""
    st.session_state["kit_notes_input"] = ""
    st.session_state["batch_count"] = 1
    st.session_state["editing_kit_id"] = None


def apply_recipe_to_builder(recipe: Dict[str, Any], items: List[Dict[str, Any]]) -> None:
    for i in items:
        iid = int(i["id"])
        st.session_state["kit_quantities"][iid] = 0
        st.session_state[f"qty_input_{iid}"] = 0
    for r_item in recipe.get("items", []):
        iid = int(r_item["item_id"])
        qty = int(r_item["quantity"])
        st.session_state["kit_quantities"][iid] = qty
        st.session_state[f"qty_input_{iid}"] = qty


def load_kit_for_editing(kit: Dict[str, Any], items: List[Dict[str, Any]]) -> None:
    reset_builder_form(items)
    st.session_state["editing_kit_id"] = int(kit["id"])
    st.session_state["batch_count"] = 1
    st.session_state["rec_name_0"] = kit.get("recipient_name", "")
    st.session_state["rec_addr_0"] = kit.get("recipient_address", "")
    st.session_state["kit_notes_input"] = kit.get("notes", "")

    name_to_id = {i["name"].lower(): int(i["id"]) for i in items}
    for ki in kit.get("items", []):
        iid = ki.get("item_id")
        if iid is None or int(iid) not in st.session_state["kit_quantities"]:
            iid = name_to_id.get(ki["item_name"].lower())
        if iid is not None:
            qty = int(ki["quantity"])
            st.session_state["kit_quantities"][int(iid)] = qty
            st.session_state[f"qty_input_{int(iid)}"] = qty
    st.session_state["active_nav"] = "Montagem do Kit"


# ==============================================================================
# ABA 1: MONTAGEM DO KIT (CHECKOUT 2 COLUNAS + LOTE + RECEITAS)
# ==============================================================================

def render_builder_tab(items: List[Dict[str, Any]], recipes: List[Dict[str, Any]]) -> None:
    editing_id = st.session_state.get("editing_kit_id")
    if editing_id is not None:
        c_edit_msg, c_edit_cancel = st.columns([4, 1])
        with c_edit_msg:
            st.info(f"Editando Kit #{editing_id} — ao salvar, o estoque será recalculado automaticamente.")
        with c_edit_cancel:
            if st.button("Cancelar edição", use_container_width=True):
                reset_builder_form(items)
                st.rerun()

    col_left, col_right = st.columns([1.65, 1.0], gap="large")

    # ---------------- COLUNA ESQUERDA: DESTINATÁRIOS + RECEITAS + BRINDES ----------------
    with col_left:
        st.markdown(
            """
            <div class="section-card">
                <div class="section-card-title">1. Destinatário(s) e Entrega</div>
                <div class="section-card-desc">Preencha um destinatário ou monte o mesmo kit para até 10 pessoas de uma só vez.</div>
            </div>
            """,
            unsafe_allow_html=True,
        )

        if editing_id is None:
            c_b1, c_b2, _ = st.columns([1.2, 1.2, 2.2])
            with c_b1:
                if st.button("+ Adicionar destinatário", use_container_width=True, disabled=st.session_state["batch_count"] >= 10):
                    st.session_state["batch_count"] += 1
                    st.rerun()
            with c_b2:
                if st.session_state["batch_count"] > 1:
                    if st.button("- Remover último", use_container_width=True):
                        idx = st.session_state["batch_count"] - 1
                        st.session_state[f"rec_name_{idx}"] = ""
                        st.session_state[f"rec_addr_{idx}"] = ""
                        st.session_state["batch_count"] -= 1
                        st.rerun()

        batch_total = 1 if editing_id is not None else int(st.session_state["batch_count"])
        for idx in range(batch_total):
            label_prefix = f"Destinatário #{idx + 1}" if batch_total > 1 else "Destinatário"
            c_n, c_a = st.columns([1.05, 1.35])
            with c_n:
                st.text_input(
                    f"Nome ({label_prefix})",
                    key=f"rec_name_{idx}",
                    placeholder="Ex.: Ana Paula Souza",
                )
            with c_a:
                st.text_input(
                    f"Endereço ({label_prefix})",
                    key=f"rec_addr_{idx}",
                    placeholder="Rua, número, complemento, cidade/UF e CEP",
                )

        st.text_input(
            "Observação interna (opcional)",
            key="kit_notes_input",
            placeholder="Ex.: Boas-vindas time comercial / Entrega urgente",
        )

        # Receitas rápidas (Kits Padrão)
        if recipes:
            st.markdown(
                """
                <div class="section-card" style="margin-top:0.85rem;">
                    <div class="section-card-title">2. Kits Padrão (Receitas Rápidas)</div>
                    <div class="section-card-desc">Aplique uma combinação pronta em 1 clique e ajuste se necessário.</div>
                </div>
                """,
                unsafe_allow_html=True,
            )
            rec_cols = st.columns(min(len(recipes), 3))
            for r_idx, rec in enumerate(recipes):
                with rec_cols[r_idx % 3]:
                    rec_total = sum(float(x["unit_price"]) * int(x["quantity"]) for x in rec.get("items", []))
                    if st.button(
                        f"{rec['name']} ({format_brl(rec_total)})",
                        key=f"apply_rec_{rec['id']}",
                        use_container_width=True,
                    ):
                        apply_recipe_to_builder(rec, items)
                        st.rerun()

        # Grade visual de Brindes
        st.markdown(
            """
            <div class="section-card" style="margin-top:0.85rem;">
                <div class="section-card-title">3. Seleção de Brindes</div>
                <div class="section-card-desc">Defina as quantidades por kit. O saldo disponível em estoque é exibido em cada item.</div>
            </div>
            """,
            unsafe_allow_html=True,
        )

        if not items:
            st.warning("Nenhum brinde ativo no catálogo. Cadastre itens na aba 'Catálogo e Brindes'.")
        else:
            grid_cols = st.columns(3)
            for idx, item in enumerate(items):
                iid = int(item["id"])
                stock = int(item.get("stock_qty", 0))
                min_s = int(item.get("min_stock", 5))
                cur_qty = int(st.session_state["kit_quantities"].get(iid, 0))
                active_cls = "item-tile item-tile-active" if cur_qty > 0 else "item-tile"

                if stock == 0:
                    stock_html = '<span class="stock-badge-out">Sem estoque (0)</span>'
                elif stock <= min_s:
                    stock_html = f'<span class="stock-badge-low">Baixo: {stock} un.</span>'
                else:
                    stock_html = f'<span class="stock-badge-ok">Estoque: {stock} un.</span>'

                img_b64 = get_image_base64(item.get("image_path", ""))
                thumb_html = (
                    f'<img src="{img_b64}" style="width:42px;height:42px;object-fit:cover;border-radius:10px;border:1px solid rgba(148,163,184,0.25);" />'
                    if img_b64
                    else f'<div style="width:42px;height:42px;border-radius:10px;background:rgba(37,99,235,0.14);color:#3B82F6;font-weight:800;font-size:0.82rem;display:flex;align-items:center;justify-content:center;">{item["code"]}</div>'
                )

                with grid_cols[idx % 3]:
                    st.markdown(
                        f"""
                        <div class="{active_cls}">
                            <div style="display:flex;align-items:center;gap:0.6rem;margin-bottom:0.45rem;">
                                {thumb_html}
                                <div style="min-width:0;">
                                    <div style="font-weight:700;font-size:0.9rem;white-space:nowrap;overflow:hidden;text-overflow:ellipsis;">{item['name']}</div>
                                    <div style="font-size:0.82rem;font-weight:600;color:#3B82F6;">{format_brl(item['unit_price'])}</div>
                                </div>
                            </div>
                            <div style="display:flex;justify-content:space-between;align-items:center;">
                                <span class="item-code-pill">[{item['code']}]</span>
                                {stock_html}
                            </div>
                        </div>
                        """,
                        unsafe_allow_html=True,
                    )
                    new_qty = st.number_input(
                        f"Qtd {item['name']}",
                        min_value=0,
                        max_value=999,
                        value=cur_qty,
                        step=1,
                        key=f"qty_input_{iid}",
                        label_visibility="collapsed",
                    )
                    if int(new_qty) != cur_qty:
                        st.session_state["kit_quantities"][iid] = int(new_qty)
                        st.rerun()

    # ---------------- COLUNA DIREITA: ORDER SUMMARY (RESUMO DO KIT) ----------------
    with col_right:
        selected_list = []
        unit_kit_total = 0.0
        total_pieces_per_kit = 0

        for item in items:
            iid = int(item["id"])
            qty = int(st.session_state["kit_quantities"].get(iid, 0))
            if qty > 0:
                sub = round(float(item["unit_price"]) * qty, 2)
                unit_kit_total = round(unit_kit_total + sub, 2)
                total_pieces_per_kit += qty
                selected_list.append({
                    "item_id": iid,
                    "item_name": item["name"],
                    "item_code": item["code"],
                    "unit_price": float(item["unit_price"]),
                    "stock_qty": int(item.get("stock_qty", 0)),
                    "quantity": qty,
                    "subtotal": sub,
                })

        recipients_filled = []
        for idx in range(batch_total):
            r_name = st.session_state.get(f"rec_name_{idx}", "").strip()
            r_addr = st.session_state.get(f"rec_addr_{idx}", "").strip()
            if r_name:
                recipients_filled.append((r_name, r_addr))

        multiplier = max(1, len(recipients_filled)) if editing_id is None else 1
        grand_total = round(unit_kit_total * multiplier, 2)

        st.markdown(
            """
            <div class="section-card">
                <div class="section-card-title">Resumo do Kit</div>
                <div class="section-card-desc">Conferência em tempo real para o formulário e baixa automática no inventário.</div>
            """,
            unsafe_allow_html=True,
        )

        if not selected_list:
            st.markdown(
                '<div style="padding:1.2rem 0;text-align:center;opacity:0.65;font-size:0.88rem;">Nenhum brinde selecionado ainda. Escolha as quantidades ao lado.</div>',
                unsafe_allow_html=True,
            )
        else:
            rows_html = ""
            for s in selected_list:
                rows_html += (
                    f'<div class="summary-row">'
                    f'<span><strong>{s["quantity"]}x</strong> {s["item_name"]} <span style="opacity:0.65;">[{s["item_code"]}]</span></span>'
                    f'<span style="font-weight:600;">{format_brl(s["subtotal"])}</span>'
                    f'</div>'
                )
            st.markdown(rows_html, unsafe_allow_html=True)

        st.markdown(
            f"""
                <div style="margin-top:0.75rem;font-size:0.83rem;opacity:0.78;display:flex;justify-content:space-between;">
                    <span>Peças por kit</span>
                    <strong>{total_pieces_per_kit} un.</strong>
                </div>
                <div style="margin-top:0.3rem;font-size:0.83rem;opacity:0.78;display:flex;justify-content:space-between;">
                    <span>Destinatários preenchidos</span>
                    <strong>{len(recipients_filled)}</strong>
                </div>
                <div class="summary-total-box">
                    <div>
                        <div style="font-size:0.76rem;font-weight:600;opacity:0.78;">VALOR DO KIT (COPIAR P/ FORMULÁRIO)</div>
                        <div style="font-size:1.35rem;font-weight:800;color:#3B82F6;">{format_brl(unit_kit_total)}</div>
                    </div>
                    <div style="text-align:right;">
                        <div style="font-size:0.74rem;opacity:0.72;">Total do lote ({multiplier}x)</div>
                        <div style="font-size:1rem;font-weight:700;">{format_brl(grand_total)}</div>
                    </div>
                </div>
            </div>
            """,
            unsafe_allow_html=True,
        )

        # Texto pronto para copiar rápido
        if selected_list:
            st.caption("Texto resumido dos itens (pronto para copiar):")
            st.code(format_items_summary(selected_list), language=None)

        # Checagem de estoque suficiente
        stock_warnings = []
        for s in selected_list:
            needed = s["quantity"] * multiplier
            if needed > s["stock_qty"]:
                stock_warnings.append(
                    f"{s['item_name']}: precisa de {needed} un., mas há {s['stock_qty']} un. no inventário."
                )
        for w in stock_warnings:
            st.warning(w)

        c_act1, c_act2 = st.columns([2, 1])
        with c_act1:
            btn_label = (
                f"Salvar alterações do Kit #{editing_id}"
                if editing_id is not None
                else (f"Finalizar {len(recipients_filled)} Kits" if len(recipients_filled) > 1 else "Finalizar Kit")
            )
            if st.button(btn_label, type="primary", use_container_width=True):
                if not selected_list:
                    st.error("Selecione pelo menos 1 brinde antes de finalizar.")
                elif not recipients_filled:
                    st.error("Informe o nome de pelo menos 1 destinatário.")
                else:
                    notes_val = st.session_state.get("kit_notes_input", "")
                    if editing_id is not None:
                        r_name, r_addr = recipients_filled[0]
                        update_kit(int(editing_id), r_name, r_addr, notes_val, selected_list)
                        reset_builder_form(items)
                        st.success(f"Kit #{editing_id} atualizado e estoque recalculado!")
                        st.rerun()
                    else:
                        created_ids = []
                        for r_name, r_addr in recipients_filled:
                            kid = create_kit(r_name, r_addr, notes_val, selected_list, deduct_stock=True)
                            created_ids.append(kid)
                        reset_builder_form(items)
                        if len(created_ids) == 1:
                            st.success(f"Kit #{created_ids[0]} finalizado e descontado do inventário!")
                        else:
                            st.success(f"{len(created_ids)} kits finalizados e descontados do inventário!")
                        st.rerun()

        with c_act2:
            if st.button("Limpar", use_container_width=True):
                reset_builder_form(items)
                st.rerun()

        # Salvar montagem atual como Kit Padrão (Receita)
        with st.expander("Salvar esta seleção como Kit Padrão"):
            new_rec_name = st.text_input("Nome do Kit Padrão", placeholder="Ex.: Kit Onboarding Executivo")
            new_rec_desc = st.text_input("Descrição curta", placeholder="Ex.: Garrafa + Caderno + Caneca")
            if st.button("Salvar Kit Padrão", use_container_width=True):
                if not new_rec_name.strip():
                    st.error("Digite um nome para o Kit Padrão.")
                elif not selected_list:
                    st.error("Selecione ao menos um brinde na grade.")
                else:
                    qty_map = {s["item_id"]: s["quantity"] for s in selected_list}
                    save_recipe(new_rec_name, new_rec_desc, qty_map)
                    st.success(f"Kit Padrão '{new_rec_name}' salvo!")
                    st.rerun()


# ==============================================================================
# ABA 2: INVENTÁRIO COMPLETO (CONTROLE DE ESTOQUE, REPOSIÇÃO E EXPORTAÇÃO)
# ==============================================================================

def render_inventory_tab(items: List[Dict[str, Any]]) -> None:
    st.markdown(
        """
        <div class="section-card">
            <div class="section-card-title">Controle de Inventário e Estoque</div>
            <div class="section-card-desc">Adicione novos brindes, faça reposição rápida ou ajuste quantidades. Cada kit finalizado dá baixa automática aqui.</div>
        </div>
        """,
        unsafe_allow_html=True,
    )

    # Exportação da planilha de inventário
    inv_df = build_inventory_dataframe(items)
    c_exp1, c_exp2, _ = st.columns([1.1, 1.1, 2.5])
    with c_exp1:
        st.download_button(
            "Baixar Inventário (.xlsx)",
            data=dataframe_to_excel_bytes(inv_df, sheet_name="Inventario"),
            file_name=f"inventario_momento_uau_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
        )
    with c_exp2:
        st.download_button(
            "Baixar Inventário (.csv)",
            data=inv_df.to_csv(index=False, sep=";").encode("utf-8-sig"),
            file_name=f"inventario_momento_uau_{datetime.now().strftime('%Y%m%d_%H%M')}.csv",
            mime="text/csv",
            use_container_width=True,
        )

    # Adicionar novo item ao inventário
    with st.expander("Adicionar novo brinde ao inventário", expanded=False):
        with st.form("form_add_inventory_item", clear_on_submit=True):
            c1, c2, c3, c4, c5 = st.columns([2.0, 0.9, 1.1, 1.0, 1.0])
            with c1:
                inv_name = st.text_input("Nome do brinde *", placeholder="Ex.: Mochila Executiva")
            with c2:
                inv_code = st.text_input("Sigla (3 letras)", max_chars=5, placeholder="MOC")
            with c3:
                inv_price = st.number_input("Valor unitário (R$)", min_value=0.0, value=0.0, step=1.0, format="%.2f")
            with c4:
                inv_stock = st.number_input("Estoque inicial", min_value=0, value=10, step=1)
            with c5:
                inv_min = st.number_input("Alerta mínimo", min_value=0, value=5, step=1)
            inv_photo = st.file_uploader("Foto do brinde (opcional)", type=["png", "jpg", "jpeg", "webp"])
            submitted_inv = st.form_submit_button("Cadastrar no Inventário", type="primary")
            if submitted_inv:
                if not inv_name.strip():
                    st.error("Informe o nome do brinde.")
                else:
                    saved_path = save_uploaded_image(inv_photo)
                    add_item(inv_name, inv_code, inv_price, int(inv_stock), int(inv_min), saved_path)
                    st.success(f"Brinde '{inv_name}' adicionado ao inventário!")
                    st.rerun()

    st.markdown("### Saldo Atual por Brinde")
    if not items:
        st.info("Nenhum brinde cadastrado.")
    else:
        for item in items:
            iid = int(item["id"])
            qty = int(item.get("stock_qty", 0))
            min_s = int(item.get("min_stock", 5))
            unit = float(item.get("unit_price", 0.0))
            status_label = "Esgotado" if qty == 0 else ("Estoque Baixo" if qty <= min_s else "OK")

            with st.container(border=True):
                c1, c2, c3, c4, c5, c6 = st.columns([2.1, 1.0, 1.1, 1.1, 0.85, 0.75])
                with c1:
                    st.markdown(f"**{item['name']}** `[{item['code']}]`")
                    st.caption(f"Unitário: {format_brl(unit)} • Total parado: {format_brl(qty * unit)} • Status: **{status_label}**")
                with c2:
                    st.metric("Saldo", f"{qty} un.")
                with c3:
                    add_delta = st.number_input(
                        "Repor/Retirar",
                        min_value=-999,
                        max_value=9999,
                        value=0,
                        step=1,
                        key=f"delta_stock_{iid}",
                        label_visibility="collapsed",
                    )
                with c4:
                    if st.button("Aplicar ajuste", key=f"btn_adj_{iid}", use_container_width=True):
                        if int(add_delta) != 0:
                            adjust_item_stock(iid, int(add_delta), reason="Ajuste rápido na aba Inventário")
                            st.rerun()
                with c5:
                    new_exact = st.number_input(
                        "Definir saldo",
                        min_value=0,
                        max_value=99999,
                        value=qty,
                        step=1,
                        key=f"exact_stock_{iid}",
                        label_visibility="collapsed",
                    )
                    if int(new_exact) != qty:
                        update_item(iid, item["name"], item["code"], unit, int(new_exact), min_s)
                        st.rerun()
                with c6:
                    if st.button("Excluir", key=f"del_inv_{iid}", use_container_width=True):
                        delete_item(iid, hard_delete=False)
                        st.rerun()

    st.markdown("### Últimas Movimentações de Estoque")
    movements = list_stock_movements(limit=40)
    if not movements:
        st.caption("Nenhuma movimentação registrada ainda.")
    else:
        mov_df = pd.DataFrame([
            {
                "Data/Hora": m["created_at"],
                "Brinde": m["item_name"],
                "Tipo": m["movement_type"],
                "Variação": f"+{m['qty_change']}" if int(m["qty_change"]) > 0 else str(m["qty_change"]),
                "Saldo Após": m["balance_after"],
                "Referência": m["reference_info"],
            }
            for m in movements
        ])
        st.dataframe(mov_df, use_container_width=True, hide_index=True)


# ==============================================================================
# ABA 3: CATÁLOGO DE BRINDES E RECEITAS EDITÁVEIS
# ==============================================================================

def render_catalog_tab(items: List[Dict[str, Any]], recipes: List[Dict[str, Any]]) -> None:
    col_cat, col_rec = st.columns([1.25, 1.0], gap="large")

    with col_cat:
        st.markdown(
            """
            <div class="section-card">
                <div class="section-card-title">Catálogo de Brindes</div>
                <div class="section-card-desc">Edite nomes, siglas, preços unitários e fotos dos brindes.</div>
            </div>
            """,
            unsafe_allow_html=True,
        )
        for item in items:
            iid = int(item["id"])
            with st.expander(f"{item['name']} [{item['code']}] — {format_brl(item['unit_price'])}"):
                with st.form(f"edit_item_form_{iid}"):
                    c1, c2 = st.columns([2, 1])
                    with c1:
                        e_name = st.text_input("Nome", value=item["name"])
                    with c2:
                        e_code = st.text_input("Sigla", value=item["code"], max_chars=5)
                    c3, c4, c5 = st.columns(3)
                    with c3:
                        e_price = st.number_input("Preço (R$)", min_value=0.0, value=float(item["unit_price"]), step=0.5, format="%.2f")
                    with c4:
                        e_stock = st.number_input("Estoque", min_value=0, value=int(item.get("stock_qty", 0)), step=1)
                    with c5:
                        e_min = st.number_input("Mínimo", min_value=0, value=int(item.get("min_stock", 5)), step=1)
                    e_photo = st.file_uploader("Trocar foto (opcional)", type=["png", "jpg", "jpeg", "webp"], key=f"up_cat_{iid}")
                    cs1, cs2 = st.columns(2)
                    with cs1:
                        if st.form_submit_button("Salvar alterações", type="primary", use_container_width=True):
                            new_img = save_uploaded_image(e_photo) if e_photo is not None else None
                            update_item(
                                iid,
                                e_name or item["name"],
                                e_code or generate_code_from_name(e_name or item["name"]),
                                e_price,
                                int(e_stock),
                                int(e_min),
                                image_path=new_img,
                            )
                            st.success("Brinde atualizado!")
                            st.rerun()
                    with cs2:
                        if st.form_submit_button("Remover do catálogo", use_container_width=True):
                            delete_item(iid, hard_delete=False)
                            st.rerun()

    with col_rec:
        st.markdown(
            """
            <div class="section-card">
                <div class="section-card-title">Receitas (Kits Padrão)</div>
                <div class="section-card-desc">Edite as quantidades dos kits pré-definidos ou exclua receitas antigas.</div>
            </div>
            """,
            unsafe_allow_html=True,
        )
        if not recipes:
            st.info("Nenhum Kit Padrão salvo ainda. Você pode criar um na aba 'Montagem do Kit'.")
        else:
            for rec in recipes:
                rid = int(rec["id"])
                rec_map = {int(x["item_id"]): int(x["quantity"]) for x in rec.get("items", [])}
                with st.expander(f"{rec['name']}"):
                    with st.form(f"edit_recipe_form_{rid}"):
                        r_name = st.text_input("Nome da receita", value=rec["name"])
                        r_desc = st.text_input("Descrição", value=rec.get("description", ""))
                        updated_map: Dict[int, int] = {}
                        for it in items:
                            it_id = int(it["id"])
                            updated_map[it_id] = st.number_input(
                                f"{it['name']} ({format_brl(it['unit_price'])})",
                                min_value=0,
                                max_value=999,
                                value=rec_map.get(it_id, 0),
                                step=1,
                                key=f"rec_item_{rid}_{it_id}",
                            )
                        cr1, cr2 = st.columns(2)
                        with cr1:
                            if st.form_submit_button("Salvar receita", type="primary", use_container_width=True):
                                if r_name.strip():
                                    save_recipe(r_name, r_desc, updated_map, recipe_id=rid)
                                    st.success("Receita atualizada!")
                                    st.rerun()
                        with cr2:
                            if st.form_submit_button("Excluir receita", use_container_width=True):
                                delete_recipe(rid)
                                st.rerun()


# ==============================================================================
# ABA 4: HISTÓRICO DE KITS FINALIZADOS E EXPORTAÇÃO (4 COLUNAS)
# ==============================================================================

def render_history_tab(items: List[Dict[str, Any]]) -> None:
    st.markdown(
        """
        <div class="section-card">
            <div class="section-card-title">Histórico de Kits Finalizados</div>
            <div class="section-card-desc">Exporte a planilha pronta (Itens e quantidades, Nome, Endereço, Valor), edite ou exclua kits com estorno automático de estoque.</div>
        </div>
        """,
        unsafe_allow_html=True,
    )

    c_search, c_xlsx, c_csv, c_clear = st.columns([2.1, 1.05, 1.05, 0.95])
    with c_search:
        q = st.text_input(
            "Buscar no histórico",
            placeholder="Filtrar por nome, endereço, observação ou brinde...",
            label_visibility="collapsed",
        )

    kits = list_kits(search_query=q)
    df_export = build_kits_dataframe(kits)

    with c_xlsx:
        st.download_button(
            "Baixar Excel (.xlsx)",
            data=dataframe_to_excel_bytes(df_export, sheet_name="Kits_Momento_Uau"),
            file_name=f"kits_momento_uau_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            type="primary",
            use_container_width=True,
        )
    with c_csv:
        st.download_button(
            "Baixar CSV (.csv)",
            data=df_export.to_csv(index=False, sep=";").encode("utf-8-sig"),
            file_name=f"kits_momento_uau_{datetime.now().strftime('%Y%m%d_%H%M')}.csv",
            mime="text/csv",
            use_container_width=True,
        )
    with c_clear:
        if st.button("Limpar tudo", use_container_width=True, disabled=len(kits) == 0):
            clear_all_kits(restore_stock=True)
            st.rerun()

    if not kits:
        st.info("Nenhum kit encontrado no histórico.")
        return

    st.markdown("#### Prévia da Planilha de Exportação")
    st.dataframe(df_export, use_container_width=True, hide_index=True)

    st.markdown("#### Kits Registrados (Editar ou Excluir Direto)")
    for kit in kits:
        kid = int(kit["id"])
        summary_txt = format_items_summary(kit.get("items", []))
        with st.container(border=True):
            c1, c2, c3, c4 = st.columns([2.3, 1.1, 0.75, 0.75])
            with c1:
                st.markdown(f"**#{kid} — {kit['recipient_name']}**")
                st.caption(f"Endereço: {kit.get('recipient_address') or 'Não informado'}")
                st.markdown(f"<span style='font-size:0.86rem;'>{summary_txt}</span>", unsafe_allow_html=True)
                if kit.get("notes"):
                    st.caption(f"Obs.: {kit['notes']}")
            with c2:
                st.markdown(f"<div style='font-size:1.15rem;font-weight:800;color:#3B82F6;'>{format_brl(kit['total_value'])}</div>", unsafe_allow_html=True)
                st.caption(kit.get("created_at", ""))
            with c3:
                if st.button("Editar", key=f"edit_kit_btn_{kid}", use_container_width=True):
                    load_kit_for_editing(kit, items)
                    st.rerun()
            with c4:
                if st.button("Excluir", key=f"del_kit_btn_{kid}", use_container_width=True):
                    delete_kit(kid, restore_stock=True)
                    st.rerun()


# ==============================================================================
# APLICAÇÃO PRINCIPAL
# ==============================================================================

def main() -> None:
    st.set_page_config(
        page_title="Montador Visual de Kits – Momento Uau",
        page_icon="🎁",
        layout="wide",
        initial_sidebar_state="collapsed",
    )

    init_db(DEFAULT_DB_PATH)
    inject_modern_blue_css()

    items = list_items(active_only=True)
    recipes = list_recipes()
    kits_all = list_kits()
    ensure_session_state(items)

    db_badge = "PostgreSQL Cloud Conectado" if is_postgres_enabled() else "Banco Local SQLite"

    # Topbar SaaS
    st.markdown(
        f"""
        <div class="uau-topbar">
            <div class="uau-brand-group">
                <div class="uau-logo-badge">UAU</div>
                <div>
                    <h1 class="uau-title">Montador Visual de Kits – Momento Uau</h1>
                    <p class="uau-subtitle">Montagem visual, cálculo automático em R$, baixa inteligente de inventário e exportação pronta</p>
                </div>
            </div>
            <div class="uau-status-pill">● {db_badge}</div>
        </div>
        """,
        unsafe_allow_html=True,
    )

    # Navegação em Pílula
    nav_options = ["Montagem do Kit", "Inventário", "Catálogo e Brindes", "Histórico"]
    selected_nav = st.radio(
        "Navegação",
        options=nav_options,
        horizontal=True,
        key="active_nav",
        label_visibility="collapsed",
    )

    # Faixa superior de KPIs arredondados
    total_kits_count = len(kits_all)
    total_kits_value = sum(float(k.get("total_value", 0.0)) for k in kits_all)
    total_stock_units = sum(int(i.get("stock_qty", 0)) for i in items)
    low_stock_count = sum(1 for i in items if int(i.get("stock_qty", 0)) <= int(i.get("min_stock", 5)))
    current_kit_val = sum(
        float(i["unit_price"]) * int(st.session_state["kit_quantities"].get(int(i["id"]), 0))
        for i in items
    )

    k1, k2, k3, k4 = st.columns(4)
    with k1:
        st.markdown(
            f"""
            <div class="kpi-card kpi-card-accent">
                <div class="kpi-label">Kit Atual em Montagem</div>
                <div class="kpi-value kpi-value-blue">{format_brl(current_kit_val)}</div>
                <div class="kpi-foot">Atualizado automaticamente</div>
            </div>
            """,
            unsafe_allow_html=True,
        )
    with k2:
        st.markdown(
            f"""
            <div class="kpi-card">
                <div class="kpi-label">Kits Finalizados</div>
                <div class="kpi-value">{total_kits_count}</div>
                <div class="kpi-foot">Prontos para exportação</div>
            </div>
            """,
            unsafe_allow_html=True,
        )
    with k3:
        st.markdown(
            f"""
            <div class="kpi-card">
                <div class="kpi-label">Acumulado em Kits</div>
                <div class="kpi-value">{format_brl(total_kits_value)}</div>
                <div class="kpi-foot">Soma do histórico salvo</div>
            </div>
            """,
            unsafe_allow_html=True,
        )
    with k4:
        st.markdown(
            f"""
            <div class="kpi-card">
                <div class="kpi-label">Brindes em Inventário</div>
                <div class="kpi-value">{total_stock_units} un.</div>
                <div class="kpi-foot">{low_stock_count} item(ns) em alerta mínimo</div>
            </div>
            """,
            unsafe_allow_html=True,
        )

    st.markdown("<div style='height:0.65rem;'></div>", unsafe_allow_html=True)

    if selected_nav == "Montagem do Kit":
        render_builder_tab(items, recipes)
    elif selected_nav == "Inventário":
        render_inventory_tab(items)
    elif selected_nav == "Catálogo e Brindes":
        render_catalog_tab(items, recipes)
    elif selected_nav == "Histórico":
        render_history_tab(items)


if __name__ == "__main__":
    main()