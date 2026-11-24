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
    {"name": "Garrafa", "unit_price": 35.00, "icon": "[GAR]", "category": "Hidratacao", "image_path": ""},
    {"name": "Blusa", "unit_price": 60.00, "icon": "[BLU]", "category": "Vestuario", "image_path": ""},
    {"name": "Chaveiro", "unit_price": 8.00, "icon": "[CHV]", "category": "Acessorios", "image_path": ""},
    {"name": "Caneca", "unit_price": 25.00, "icon": "[CAN]", "category": "Escritorio", "image_path": ""},
    {"name": "Caderno", "unit_price": 18.00, "icon": "[CAD]", "category": "Papelaria", "image_path": ""},
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


def get_connection(db_path: Path = DB_PATH) -> sqlite3.Connection:
    db_path.parent.mkdir(parents=True, exist_ok=True)
    ASSETS_DIR.mkdir(parents=True, exist_ok=True)
    conn = sqlite3.connect(str(db_path), check_same_thread=False)
    conn.row_factory = sqlite3.Row
    conn.execute("PRAGMA foreign_keys = ON;")
    return conn


def _ensure_kits_address_columns(conn: sqlite3.Connection) -> None:
    existing_cols = {
        row["name"] for row in conn.execute("PRAGMA table_info(kits)").fetchall()
    }
    for col_name, col_def in ADDRESS_COLUMNS.items():
        if col_name not in existing_cols:
            conn.execute(f"ALTER TABLE kits ADD COLUMN {col_name} {col_def}")


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
            """
        )
        _ensure_kits_address_columns(conn)

        now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

        count_items = conn.execute("SELECT COUNT(*) AS cnt FROM catalog_items").fetchone()["cnt"]
        if count_items == 0:
            for item in DEFAULT_ITEMS:
                conn.execute(
                    """
                    INSERT INTO catalog_items (name, unit_price, icon, category, image_path, created_at)
                    VALUES (?, ?, ?, ?, ?, ?)
                    """,
                    (
                        item["name"],
                        float(item["unit_price"]),
                        item["icon"],
                        item["category"],
                        item["image_path"],
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


def get_catalog_items(db_path: Path = DB_PATH) -> list[dict[str, Any]]:
    with get_connection(db_path) as conn:
        rows = conn.execute(
            """
            SELECT id, name, unit_price, icon, category, image_path, created_at
            FROM catalog_items
            ORDER BY id ASC
            """
        ).fetchall()
        return [dict(r) for r in rows]


def add_catalog_item(
    name: str,
    unit_price: float,
    icon: str = "[ITEM]",
    category: str = "Geral",
    image_path: str = "",
    db_path: Path = DB_PATH,
) -> tuple[bool, str]:
    clean_name = name.strip()
    if not clean_name:
        return False, "Informe o nome do brinde."
    if unit_price < 0:
        return False, "O valor unitario nao pode ser negativo."

    clean_icon = icon.strip() or "[ITEM]"
    clean_cat = category.strip() or "Geral"
    now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

    try:
        with get_connection(db_path) as conn:
            conn.execute(
                """
                INSERT INTO catalog_items (name, unit_price, icon, category, image_path, created_at)
                VALUES (?, ?, ?, ?, ?, ?)
                """,
                (clean_name, float(unit_price), clean_icon, clean_cat, image_path.strip(), now),
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
    db_path: Path = DB_PATH,
) -> tuple[bool, str]:
    clean_name = name.strip()
    if not clean_name:
        return False, "O nome do brinde nao pode ficar vazio."
    if unit_price < 0:
        return False, "O valor unitario nao pode ser negativo."

    clean_icon = icon.strip() or "[ITEM]"
    clean_cat = category.strip() or "Geral"

    try:
        with get_connection(db_path) as conn:
            cur = conn.execute(
                """
                UPDATE catalog_items
                SET name = ?, unit_price = ?, icon = ?, category = ?, image_path = ?
                WHERE id = ?
                """,
                (clean_name, float(unit_price), clean_icon, clean_cat, image_path.strip(), item_id),
            )
            if cur.rowcount == 0:
                return False, "Item nao encontrado no catalogo."
        return True, f"Brinde '{clean_name}' atualizado."
    except sqlite3.IntegrityError:
        return False, f"Ja existe outro brinde chamado '{clean_name}'."


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

    return True, f"Brinde '{row['name']}' excluido do catalogo."


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
        if kit_id is not None:
            existing = conn.execute("SELECT id, created_at FROM kits WHERE id = ?", (kit_id,)).fetchone()
            if not existing:
                return False, "Kit nao encontrado para atualizacao.", None
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
    with get_connection(db_path) as conn:
        row = conn.execute("SELECT kit_code FROM kits WHERE id = ?", (kit_id,)).fetchone()
        if not row:
            return False, "Kit nao encontrado."
        conn.execute("DELETE FROM kits WHERE id = ?", (kit_id,))
    return True, f"Kit '{row['kit_code']}' excluido do historico."
