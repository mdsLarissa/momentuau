from __future__ import annotations

import json
import sqlite3
from datetime import datetime
from pathlib import Path
from typing import Any

BASE_DIR = Path(__file__).resolve().parent
DATA_DIR = BASE_DIR / "data"
ASSETS_DIR = BASE_DIR / "assets"
DB_PATH = DATA_DIR / "momento_uau.db"

DEFAULT_ITEMS: list[dict[str, Any]] = [
    {"name": "Garrafa", "unit_price": 35.00, "icon": "[GAR]", "category": "Hidratacao", "image_path": "", "stock_qty": 50, "min_stock": 10},
    {"name": "Blusa", "unit_price": 60.00, "icon": "[BLU]", "category": "Vestuario", "image_path": "", "stock_qty": 40, "min_stock": 10},
    {"name": "Chaveiro", "unit_price": 8.00, "icon": "[CHV]", "category": "Acessorios", "image_path": "", "stock_qty": 100, "min_stock": 20},
    {"name": "Caneca", "unit_price": 25.00, "icon": "[CAN]", "category": "Escritorio", "image_path": "", "stock_qty": 50, "min_stock": 10},
    {"name": "Caderno", "unit_price": 18.00, "icon": "[CAD]", "category": "Papelaria", "image_path": "", "stock_qty": 60, "min_stock": 15},
]

DEFAULT_RECIPES: list[dict[str, Any]] = [
    {
        "name": "Kit A - Essencial",
        "description": "Garrafa e chaveiro para boas-vindas rapidas.",
        "items": {"Garrafa": 1, "Chaveiro": 1},
    },
    {
        "name": "Kit B - Completo",
        "description": "Garrafa, blusa e chaveiro.",
        "items": {"Garrafa": 1, "Blusa": 1, "Chaveiro": 1},
    },
    {
        "name": "Kit C - Especial",
        "description": "Garrafa, blusa, chaveiro e caneca (R$ 128,00).",
        "items": {"Garrafa": 1, "Blusa": 1, "Chaveiro": 1, "Caneca": 1},
    },
]

ADDRESS_COLUMNS: dict[str, str] = {
    "cep": "TEXT NOT NULL DEFAULT ''",
    "street": "TEXT NOT NULL DEFAULT ''",
    "address_number": "TEXT NOT NULL DEFAULT ''",
    "complement": "TEXT NOT NULL DEFAULT ''",
    "neighborhood": "TEXT NOT NULL DEFAULT ''",
    "city_state": "TEXT NOT NULL DEFAULT ''",
}

CATALOG_INVENTORY_COLUMNS: dict[str, str] = {
    "stock_qty": "INTEGER NOT NULL DEFAULT 0",
    "min_stock": "INTEGER NOT NULL DEFAULT 5",
    "in_inventory": "INTEGER NOT NULL DEFAULT 1",
}


def get_connection(db_path: Path = DB_PATH) -> sqlite3.Connection:
    db_path.parent.mkdir(parents=True, exist_ok=True)
    ASSETS_DIR.mkdir(parents=True, exist_ok=True)
    conn = sqlite3.connect(str(db_path), check_same_thread=False)
    conn.row_factory = sqlite3.Row
    conn.execute("PRAGMA foreign_keys = ON;")
    return conn


def _ensure_table_columns(conn: sqlite3.Connection, table_name: str, columns: dict[str, str]) -> None:
    existing_cols = {
        row["name"] for row in conn.execute(f"PRAGMA table_info({table_name})").fetchall()
    }
    for col_name, col_def in columns.items():
        if col_name not in existing_cols:
            conn.execute(f"ALTER TABLE {table_name} ADD COLUMN {col_name} {col_def}")


def init_db(db_path: Path = DB_PATH) -> None:
    with get_connection(db_path) as conn:
        conn.executescript(
            """
            CREATE TABLE IF NOT EXISTS catalog_items (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                name TEXT NOT NULL UNIQUE,
                unit_price REAL NOT NULL CHECK (unit_price >= 0),
                icon TEXT NOT NULL DEFAULT '[ITEM]',
                category TEXT NOT NULL DEFAULT 'Geral',
                image_path TEXT NOT NULL DEFAULT '',
                stock_qty INTEGER NOT NULL DEFAULT 0,
                min_stock INTEGER NOT NULL DEFAULT 5,
                in_inventory INTEGER NOT NULL DEFAULT 1,
                created_at TEXT NOT NULL
            );

            CREATE TABLE IF NOT EXISTS recipes (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                name TEXT NOT NULL UNIQUE,
                description TEXT NOT NULL DEFAULT '',
                items_json TEXT NOT NULL,
                created_at TEXT NOT NULL
            );

            CREATE TABLE IF NOT EXISTS kits (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                kit_code TEXT NOT NULL,
                client_name TEXT NOT NULL,
                cep TEXT NOT NULL DEFAULT '',
                street TEXT NOT NULL DEFAULT '',
                address_number TEXT NOT NULL DEFAULT '',
                complement TEXT NOT NULL DEFAULT '',
                neighborhood TEXT NOT NULL DEFAULT '',
                city_state TEXT NOT NULL DEFAULT '',
                notes TEXT NOT NULL DEFAULT '',
                status TEXT NOT NULL DEFAULT 'Finalizado',
                total_quantity INTEGER NOT NULL DEFAULT 0,
                total_value REAL NOT NULL DEFAULT 0.0,
                created_at TEXT NOT NULL,
                updated_at TEXT NOT NULL
            );

            CREATE TABLE IF NOT EXISTS kit_items (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                kit_id INTEGER NOT NULL,
                item_id INTEGER NOT NULL,
                item_name TEXT NOT NULL,
                unit_price REAL NOT NULL,
                quantity INTEGER NOT NULL CHECK (quantity > 0),
                subtotal REAL NOT NULL,
                FOREIGN KEY (kit_id) REFERENCES kits(id) ON DELETE CASCADE,
                FOREIGN KEY (item_id) REFERENCES catalog_items(id) ON DELETE RESTRICT
            );

            CREATE TABLE IF NOT EXISTS inventory_movements (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                item_id INTEGER NOT NULL,
                item_name TEXT NOT NULL,
                delta INTEGER NOT NULL,
                balance_after INTEGER NOT NULL,
                reason TEXT NOT NULL DEFAULT '',
                created_at TEXT NOT NULL
            );
            """
        )
        _ensure_table_columns(conn, "kits", ADDRESS_COLUMNS)
        _ensure_table_columns(conn, "catalog_items", CATALOG_INVENTORY_COLUMNS)

        now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

        count_items = conn.execute("SELECT COUNT(*) AS cnt FROM catalog_items").fetchone()["cnt"]
        if count_items == 0:
            for item in DEFAULT_ITEMS:
                cur = conn.execute(
                    """
                    INSERT INTO catalog_items (
                        name, unit_price, icon, category, image_path,
                        stock_qty, min_stock, in_inventory, created_at
                    )
                    VALUES (?, ?, ?, ?, ?, ?, ?, 1, ?)
                    """,
                    (
                        item["name"],
                        float(item["unit_price"]),
                        item["icon"],
                        item["category"],
                        item["image_path"],
                        int(item.get("stock_qty", 0)),
                        int(item.get("min_stock", 5)),
                        now,
                    ),
                )
                if int(item.get("stock_qty", 0)) > 0:
                    conn.execute(
                        """
                        INSERT INTO inventory_movements (item_id, item_name, delta, balance_after, reason, created_at)
                        VALUES (?, ?, ?, ?, ?, ?)
                        """,
                        (
                            int(cur.lastrowid),
                            item["name"],
                            int(item["stock_qty"]),
                            int(item["stock_qty"]),
                            "Saldo inicial",
                            now,
                        ),
                    )

        name_to_id = {
            row["name"]: row["id"]
            for row in conn.execute("SELECT id, name FROM catalog_items").fetchall()
        }

        count_recipes = conn.execute("SELECT COUNT(*) AS cnt FROM recipes").fetchone()["cnt"]
        if count_recipes == 0:
            for recipe in DEFAULT_RECIPES:
                mapped_items: dict[str, int] = {}
                for item_name, qty in recipe["items"].items():
                    if item_name in name_to_id and int(qty) > 0:
                        mapped_items[str(name_to_id[item_name])] = int(qty)
                if mapped_items:
                    conn.execute(
                        """
                        INSERT INTO recipes (name, description, items_json, created_at)
                        VALUES (?, ?, ?, ?)
                        """,
                        (
                            recipe["name"],
                            recipe["description"],
                            json.dumps(mapped_items, ensure_ascii=False),
                            now,
                        ),
                    )


def _log_inventory_movement(
    conn: sqlite3.Connection,
    item_id: int,
    item_name: str,
    delta: int,
    balance_after: int,
    reason: str,
    now: str,
) -> None:
    if delta == 0:
        return
    conn.execute(
        """
        INSERT INTO inventory_movements (item_id, item_name, delta, balance_after, reason, created_at)
        VALUES (?, ?, ?, ?, ?, ?)
        """,
        (item_id, item_name, int(delta), int(balance_after), reason.strip(), now),
    )


def get_catalog_items(db_path: Path = DB_PATH) -> list[dict[str, Any]]:
    with get_connection(db_path) as conn:
        rows = conn.execute(
            """
            SELECT id, name, unit_price, icon, category, image_path,
                   stock_qty, min_stock, in_inventory, created_at
            FROM catalog_items
            ORDER BY id ASC
            """
        ).fetchall()
        return [dict(r) for r in rows]


def get_inventory_items(db_path: Path = DB_PATH) -> list[dict[str, Any]]:
    with get_connection(db_path) as conn:
        rows = conn.execute(
            """
            SELECT id, name, unit_price, icon, category, image_path,
                   stock_qty, min_stock, in_inventory, created_at
            FROM catalog_items
            WHERE in_inventory = 1
            ORDER BY name ASC
            """
        ).fetchall()
        return [dict(r) for r in rows]


def get_inventory_movements(limit: int = 40, db_path: Path = DB_PATH) -> list[dict[str, Any]]:
    with get_connection(db_path) as conn:
        rows = conn.execute(
            """
            SELECT id, item_id, item_name, delta, balance_after, reason, created_at
            FROM inventory_movements
            ORDER BY id DESC
            LIMIT ?
            """,
            (int(limit),),
        ).fetchall()
        return [dict(r) for r in rows]


def add_catalog_item(
    name: str,
    unit_price: float,
    icon: str = "[ITEM]",
    category: str = "Geral",
    image_path: str = "",
    stock_qty: int = 0,
    min_stock: int = 5,
    in_inventory: int = 1,
    db_path: Path = DB_PATH,
) -> tuple[bool, str]:
    clean_name = name.strip()
    if not clean_name:
        return False, "Informe o nome do brinde."
    if unit_price < 0:
        return False, "O valor unitario nao pode ser negativo."
    if int(stock_qty) < 0 or int(min_stock) < 0:
        return False, "As quantidades de estoque nao podem ser negativas."

    clean_icon = icon.strip() or "[ITEM]"
    clean_cat = category.strip() or "Geral"
    now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    qty_init = int(stock_qty)

    try:
        with get_connection(db_path) as conn:
            cur = conn.execute(
                """
                INSERT INTO catalog_items (
                    name, unit_price, icon, category, image_path,
                    stock_qty, min_stock, in_inventory, created_at
                )
                VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)
                """,
                (
                    clean_name,
                    float(unit_price),
                    clean_icon,
                    clean_cat,
                    image_path.strip(),
                    qty_init,
                    int(min_stock),
                    1 if in_inventory else 0,
                    now,
                ),
            )
            if qty_init > 0:
                _log_inventory_movement(
                    conn,
                    int(cur.lastrowid),
                    clean_name,
                    qty_init,
                    qty_init,
                    "Cadastro inicial",
                    now,
                )
        return True, f"Brinde '{clean_name}' cadastrado com sucesso."
    except sqlite3.IntegrityError:
        return False, f"Ja existe um brinde chamado '{clean_name}' no catalogo."


def update_catalog_item(
    item_id: int,
    name: str,
    unit_price: float,
    icon: str,
    category: str,
    image_path: str,
    stock_qty: int | None = None,
    min_stock: int | None = None,
    in_inventory: int | None = None,
    db_path: Path = DB_PATH,
) -> tuple[bool, str]:
    clean_name = name.strip()
    if not clean_name:
        return False, "O nome do brinde nao pode ficar vazio."
    if unit_price < 0:
        return False, "O valor unitario nao pode ser negativo."

    clean_icon = icon.strip() or "[ITEM]"
    clean_cat = category.strip() or "Geral"
    now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

    try:
        with get_connection(db_path) as conn:
            existing = conn.execute(
                "SELECT stock_qty, min_stock, in_inventory FROM catalog_items WHERE id = ?",
                (item_id,),
            ).fetchone()
            if not existing:
                return False, "Item nao encontrado no catalogo."

            old_qty = int(existing["stock_qty"])
            new_qty = old_qty if stock_qty is None else max(0, int(stock_qty))
            new_min = int(existing["min_stock"]) if min_stock is None else max(0, int(min_stock))
            new_inv = int(existing["in_inventory"]) if in_inventory is None else (1 if in_inventory else 0)

            conn.execute(
                """
                UPDATE catalog_items
                SET name = ?, unit_price = ?, icon = ?, category = ?, image_path = ?,
                    stock_qty = ?, min_stock = ?, in_inventory = ?
                WHERE id = ?
                """,
                (
                    clean_name,
                    float(unit_price),
                    clean_icon,
                    clean_cat,
                    image_path.strip(),
                    new_qty,
                    new_min,
                    new_inv,
                    item_id,
                ),
            )
            if new_qty != old_qty:
                _log_inventory_movement(
                    conn,
                    item_id,
                    clean_name,
                    new_qty - old_qty,
                    new_qty,
                    "Ajuste manual de inventario",
                    now,
                )
        return True, f"Brinde '{clean_name}' atualizado."
    except sqlite3.IntegrityError:
        return False, f"Ja existe outro brinde chamado '{clean_name}'."


def adjust_inventory_stock(
    item_id: int,
    delta: int | None = None,
    exact_qty: int | None = None,
    min_stock: int | None = None,
    reason: str = "Movimentacao manual",
    db_path: Path = DB_PATH,
) -> tuple[bool, str]:
    now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    with get_connection(db_path) as conn:
        row = conn.execute(
            "SELECT id, name, stock_qty, min_stock FROM catalog_items WHERE id = ?",
            (item_id,),
        ).fetchone()
        if not row:
            return False, "Item nao encontrado."

        old_qty = int(row["stock_qty"])
        if exact_qty is not None:
            new_qty = max(0, int(exact_qty))
        else:
            new_qty = max(0, old_qty + int(delta or 0))

        new_min = int(row["min_stock"]) if min_stock is None else max(0, int(min_stock))
        conn.execute(
            """
            UPDATE catalog_items
            SET stock_qty = ?, min_stock = ?, in_inventory = 1
            WHERE id = ?
            """,
            (new_qty, new_min, item_id),
        )
        diff = new_qty - old_qty
        if diff != 0:
            _log_inventory_movement(conn, item_id, row["name"], diff, new_qty, reason, now)

    return True, f"Estoque de '{row['name']}': {new_qty} un."


def remove_from_inventory(
    item_id: int,
    active_kit_items: dict[int, int] | None = None,
    db_path: Path = DB_PATH,
) -> tuple[bool, str]:
    """Remove o item completamente se nao estiver em kits salvos; caso contrario, remove da visao de inventario e zera o saldo."""
    active_map = active_kit_items or {}
    if item_id in active_map and active_map[item_id] > 0:
        return False, "Este item esta no kit atual. Remova-o da caixa antes de excluir."

    now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    if not is_item_in_saved_kits(item_id, db_path):
        return delete_catalog_item(item_id, active_kit_items=active_map, db_path=db_path)

    with get_connection(db_path) as conn:
        row = conn.execute(
            "SELECT name, stock_qty FROM catalog_items WHERE id = ?",
            (item_id,),
        ).fetchone()
        if not row:
            return False, "Item nao encontrado."
        old_qty = int(row["stock_qty"])
        conn.execute(
            "UPDATE catalog_items SET stock_qty = 0, in_inventory = 0 WHERE id = ?",
            (item_id,),
        )
        if old_qty > 0:
            _log_inventory_movement(
                conn, item_id, row["name"], -old_qty, 0, "Removido do inventario", now
            )
    return True, f"'{row['name']}' removido do inventario (mantido apenas no historico antigo)."


def is_item_in_saved_kits(item_id: int, db_path: Path = DB_PATH) -> bool:
    with get_connection(db_path) as conn:
        row = conn.execute(
            "SELECT COUNT(*) AS cnt FROM kit_items WHERE item_id = ?",
            (item_id,),
        ).fetchone()
        return bool(row and row["cnt"] > 0)


def delete_catalog_item(
    item_id: int,
    active_kit_items: dict[int, int] | None = None,
    db_path: Path = DB_PATH,
) -> tuple[bool, str]:
    active_map = active_kit_items or {}
    if item_id in active_map and active_map[item_id] > 0:
        return False, "Este item esta no kit atual. Remova-o do kit antes de excluir do catalogo."

    if is_item_in_saved_kits(item_id, db_path):
        return (
            False,
            "Este item faz parte de kits salvos no historico e nao pode ser excluido para manter a integridade dos registros.",
        )

    with get_connection(db_path) as conn:
        row = conn.execute("SELECT name, image_path FROM catalog_items WHERE id = ?", (item_id,)).fetchone()
        if not row:
            return False, "Item nao encontrado."
        img_str = row["image_path"]
        conn.execute("DELETE FROM catalog_items WHERE id = ?", (item_id,))

    if img_str:
        try:
            img_path = Path(img_str)
            if img_path.exists() and ASSETS_DIR.resolve() in img_path.resolve().parents:
                img_path.unlink(missing_ok=True)
        except OSError:
            pass

    return True, f"Brinde '{row['name']}' excluido."


def get_recipes(db_path: Path = DB_PATH) -> list[dict[str, Any]]:
    with get_connection(db_path) as conn:
        rows = conn.execute(
            "SELECT id, name, description, items_json, created_at FROM recipes ORDER BY id ASC"
        ).fetchall()
        result: list[dict[str, Any]] = []
        for r in rows:
            try:
                raw_map = json.loads(r["items_json"])
            except json.JSONDecodeError:
                raw_map = {}
            parsed_items: dict[int, int] = {}
            for k, v in raw_map.items():
                try:
                    iid = int(k)
                    qty = int(v)
                    if qty > 0:
                        parsed_items[iid] = qty
                except (ValueError, TypeError):
                    continue
            result.append(
                {
                    "id": r["id"],
                    "name": r["name"],
                    "description": r["description"],
                    "items": parsed_items,
                    "created_at": r["created_at"],
                }
            )
        return result


def save_recipe(
    name: str,
    description: str,
    kit_items: dict[int, int],
    db_path: Path = DB_PATH,
) -> tuple[bool, str]:
    clean_name = name.strip()
    if not clean_name:
        return False, "Informe um nome para a receita."
    valid_items = {str(int(k)): int(v) for k, v in kit_items.items() if int(v) > 0}
    if not valid_items:
        return False, "Adicione ao menos 1 item ao kit antes de salvar como receita."

    now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    try:
        with get_connection(db_path) as conn:
            conn.execute(
                """
                INSERT INTO recipes (name, description, items_json, created_at)
                VALUES (?, ?, ?, ?)
                """,
                (clean_name, description.strip(), json.dumps(valid_items, ensure_ascii=False), now),
            )
        return True, f"Receita '{clean_name}' salva com sucesso."
    except sqlite3.IntegrityError:
        return False, f"Ja existe uma receita chamada '{clean_name}'."


def update_recipe(
    recipe_id: int,
    name: str,
    description: str,
    kit_items: dict[int, int],
    db_path: Path = DB_PATH,
) -> tuple[bool, str]:
    clean_name = name.strip()
    if not clean_name:
        return False, "Informe um nome para a receita."
    valid_items = {str(int(k)): int(v) for k, v in kit_items.items() if int(v) > 0}
    if not valid_items:
        return False, "A receita precisa ter pelo menos 1 item."

    try:
        with get_connection(db_path) as conn:
            cur = conn.execute(
                """
                UPDATE recipes
                SET name = ?, description = ?, items_json = ?
                WHERE id = ?
                """,
                (clean_name, description.strip(), json.dumps(valid_items, ensure_ascii=False), recipe_id),
            )
            if cur.rowcount == 0:
                return False, "Receita nao encontrada."
        return True, f"Receita '{clean_name}' atualizada."
    except sqlite3.IntegrityError:
        return False, f"Ja existe outra receita chamada '{clean_name}'."


def delete_recipe(recipe_id: int, db_path: Path = DB_PATH) -> tuple[bool, str]:
    with get_connection(db_path) as conn:
        row = conn.execute("SELECT name FROM recipes WHERE id = ?", (recipe_id,)).fetchone()
        if not row:
            return False, "Receita nao encontrada."
        conn.execute("DELETE FROM recipes WHERE id = ?", (recipe_id,))
    return True, f"Receita '{row['name']}' excluida."


def _compute_kit_totals(
    kit_items: dict[int, int],
    catalog_by_id: dict[int, dict[str, Any]],
) -> tuple[int, float, list[dict[str, Any]]]:
    total_qty = 0
    total_val = 0.0
    lines: list[dict[str, Any]] = []

    for item_id, qty in kit_items.items():
        q = int(qty)
        if q <= 0 or item_id not in catalog_by_id:
            continue
        cat_item = catalog_by_id[item_id]
        unit_price = float(cat_item["unit_price"])
        subtotal = round(unit_price * q, 2)
        total_qty += q
        total_val = round(total_val + subtotal, 2)
        lines.append(
            {
                "item_id": item_id,
                "item_name": cat_item["name"],
                "unit_price": unit_price,
                "quantity": q,
                "subtotal": subtotal,
            }
        )
    return total_qty, total_val, lines


def format_address_summary(kit_data: dict[str, Any]) -> str:
    street = str(kit_data.get("street") or "").strip()
    number = str(kit_data.get("address_number") or "").strip()
    complement = str(kit_data.get("complement") or "").strip()
    neighborhood = str(kit_data.get("neighborhood") or "").strip()
    city_state = str(kit_data.get("city_state") or "").strip()
    cep = str(kit_data.get("cep") or "").strip()

    main_line = street
    if number:
        main_line = f"{main_line}, {number}" if main_line else number
    if complement:
        main_line = f"{main_line} ({complement})" if main_line else complement

    parts = [p for p in [main_line, neighborhood, city_state, f"CEP {cep}" if cep else ""] if p]
    return " - ".join(parts)


def save_kit(
    kit_code: str,
    client_name: str,
    notes: str,
    status: str,
    kit_items: dict[int, int],
    catalog_by_id: dict[int, dict[str, Any]],
    cep: str = "",
    street: str = "",
    address_number: str = "",
    complement: str = "",
    neighborhood: str = "",
    city_state: str = "",
    kit_id: int | None = None,
    db_path: Path = DB_PATH,
) -> tuple[bool, str, int | None]:
    clean_code = kit_code.strip() or "Kit sem numero"
    clean_client = client_name.strip()
    if not clean_client:
        return False, "Informe o nome do cliente antes de finalizar ou salvar o kit.", None

    total_qty, total_val, lines = _compute_kit_totals(kit_items, catalog_by_id)
    if not lines or total_qty <= 0:
        return False, "O kit precisa ter pelo menos 1 item para ser finalizado.", None

    now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    clean_cep = cep.strip()
    clean_street = street.strip()
    clean_number = address_number.strip()
    clean_comp = complement.strip()
    clean_neigh = neighborhood.strip()
    clean_city = city_state.strip()

    with get_connection(db_path) as conn:
        old_quantities: dict[int, int] = {}
        if kit_id is not None:
            existing = conn.execute("SELECT id FROM kits WHERE id = ?", (kit_id,)).fetchone()
            if not existing:
                return False, "Kit nao encontrado para atualizacao.", None

            for old_row in conn.execute(
                "SELECT item_id, quantity FROM kit_items WHERE kit_id = ?", (kit_id,)
            ).fetchall():
                oid = int(old_row["item_id"])
                old_quantities[oid] = old_quantities.get(oid, 0) + int(old_row["quantity"])

        new_quantities: dict[int, int] = {
            int(line["item_id"]): int(line["quantity"]) for line in lines
        }

        # Valida saldo disponivel no inventario antes de efetivar baixa
        all_item_ids = set(old_quantities) | set(new_quantities)
        for iid in all_item_ids:
            net_needed = new_quantities.get(iid, 0) - old_quantities.get(iid, 0)
            if net_needed > 0:
                inv_row = conn.execute(
                    "SELECT name, stock_qty FROM catalog_items WHERE id = ?", (iid,)
                ).fetchone()
                if inv_row and int(inv_row["stock_qty"]) < net_needed:
                    return (
                        False,
                        f"Estoque insuficiente de '{inv_row['name']}' "
                        f"(disponivel: {inv_row['stock_qty']}, necessario: {net_needed}).",
                        None,
                    )

        if kit_id is not None:
            conn.execute(
                """
                UPDATE kits
                SET kit_code = ?, client_name = ?,
                    cep = ?, street = ?, address_number = ?, complement = ?,
                    neighborhood = ?, city_state = ?,
                    notes = ?, status = ?,
                    total_quantity = ?, total_value = ?, updated_at = ?
                WHERE id = ?
                """,
                (
                    clean_code,
                    clean_client,
                    clean_cep,
                    clean_street,
                    clean_number,
                    clean_comp,
                    clean_neigh,
                    clean_city,
                    notes.strip(),
                    status,
                    total_qty,
                    total_val,
                    now,
                    kit_id,
                ),
            )
            conn.execute("DELETE FROM kit_items WHERE kit_id = ?", (kit_id,))
            target_id = kit_id
        else:
            cur = conn.execute(
                """
                INSERT INTO kits (
                    kit_code, client_name,
                    cep, street, address_number, complement,
                    neighborhood, city_state,
                    notes, status,
                    total_quantity, total_value, created_at, updated_at
                )
                VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                """,
                (
                    clean_code,
                    clean_client,
                    clean_cep,
                    clean_street,
                    clean_number,
                    clean_comp,
                    clean_neigh,
                    clean_city,
                    notes.strip(),
                    status,
                    total_qty,
                    total_val,
                    now,
                    now,
                ),
            )
            target_id = int(cur.lastrowid)

        for line in lines:
            conn.execute(
                """
                INSERT INTO kit_items (
                    kit_id, item_id, item_name, unit_price, quantity, subtotal
                )
                VALUES (?, ?, ?, ?, ?, ?)
                """,
                (
                    target_id,
                    line["item_id"],
                    line["item_name"],
                    line["unit_price"],
                    line["quantity"],
                    line["subtotal"],
                ),
            )

        # Atualiza o saldo no inventario (baixa automatica ou diferenca de edicao)
        for iid in all_item_ids:
            net_used = new_quantities.get(iid, 0) - old_quantities.get(iid, 0)
            if net_used == 0:
                continue
            inv_row = conn.execute(
                "SELECT name, stock_qty FROM catalog_items WHERE id = ?", (iid,)
            ).fetchone()
            if inv_row:
                new_balance = max(0, int(inv_row["stock_qty"]) - net_used)
                conn.execute(
                    "UPDATE catalog_items SET stock_qty = ? WHERE id = ?",
                    (new_balance, iid),
                )
                reason = (
                    f"Saida no {clean_code} ({clean_client})"
                    if net_used > 0
                    else f"Devolucao por edicao do {clean_code}"
                )
                _log_inventory_movement(
                    conn, iid, inv_row["name"], -net_used, new_balance, reason, now
                )

    return True, f"Kit '{clean_code}' salvo com sucesso.", target_id


def get_kits_history(db_path: Path = DB_PATH) -> list[dict[str, Any]]:
    with get_connection(db_path) as conn:
        kit_rows = conn.execute(
            """
            SELECT id, kit_code, client_name,
                   cep, street, address_number, complement, neighborhood, city_state,
                   notes, status, total_quantity, total_value, created_at, updated_at
            FROM kits
            ORDER BY id DESC
            """
        ).fetchall()

        item_rows = conn.execute(
            """
            SELECT kit_id, item_id, item_name, unit_price, quantity, subtotal
            FROM kit_items
            ORDER BY id ASC
            """
        ).fetchall()

    items_by_kit: dict[int, list[dict[str, Any]]] = {}
    for ir in item_rows:
        items_by_kit.setdefault(ir["kit_id"], []).append(dict(ir))

    history: list[dict[str, Any]] = []
    for kr in kit_rows:
        k_dict = dict(kr)
        k_items = items_by_kit.get(kr["id"], [])
        k_dict["items"] = k_items
        k_dict["items_summary"] = ", ".join(
            f"{it['quantity']}x {it['item_name']}" for it in k_items
        )
        k_dict["address_summary"] = format_address_summary(k_dict)
        history.append(k_dict)
    return history


def delete_kit(kit_id: int, db_path: Path = DB_PATH) -> tuple[bool, str]:
    now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    with get_connection(db_path) as conn:
        row = conn.execute(
            "SELECT kit_code, client_name FROM kits WHERE id = ?", (kit_id,)
        ).fetchone()
        if not row:
            return False, "Kit nao encontrado."

        # Devolve os itens do kit excluido ao inventario
        items_to_restore = conn.execute(
            "SELECT item_id, quantity FROM kit_items WHERE kit_id = ?", (kit_id,)
        ).fetchall()
        for it in items_to_restore:
            iid = int(it["item_id"])
            qty = int(it["quantity"])
            inv_row = conn.execute(
                "SELECT name, stock_qty FROM catalog_items WHERE id = ?", (iid,)
            ).fetchone()
            if inv_row and qty > 0:
                new_balance = int(inv_row["stock_qty"]) + qty
                conn.execute(
                    "UPDATE catalog_items SET stock_qty = ? WHERE id = ?",
                    (new_balance, iid),
                )
                _log_inventory_movement(
                    conn,
                    iid,
                    inv_row["name"],
                    qty,
                    new_balance,
                    f"Estorno por exclusao do {row['kit_code']}",
                    now,
                )

        conn.execute("DELETE FROM kits WHERE id = ?", (kit_id,))
    return True, f"Kit '{row['kit_code']}' excluido e itens devolvidos ao inventario."