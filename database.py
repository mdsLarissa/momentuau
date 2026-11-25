import os
import re
import sqlite3
from contextlib import contextmanager
from datetime import datetime
from typing import Any, Dict, List, Optional, Tuple
from urllib.parse import parse_qsl, urlencode, urlparse, urlunparse

try:
    import streamlit as st
except Exception:
    st = None

try:
    import psycopg2
    from psycopg2.extras import RealDictCursor
except Exception:
    psycopg2 = None
    RealDictCursor = None

DEFAULT_DB_PATH = os.path.join("data", "momento_uau.db")


def _get_database_url() -> str:
    """Obtém a URL do PostgreSQL dos Secrets do Streamlit ou variáveis de ambiente."""
    url = ""
    if st is not None:
        try:
            if "DATABASE_URL" in st.secrets:
                url = str(st.secrets["DATABASE_URL"]).strip()
        except Exception:
            pass
    if not url:
        url = os.getenv("DATABASE_URL", "").strip()
    return url.strip('"').strip("'").strip()


def _normalize_postgres_url(raw_url: str) -> str:
    """Garante sslmode=require e connect_timeout=10 na string de conexão PostgreSQL."""
    cleaned = raw_url.strip().strip('"').strip("'")
    if cleaned.startswith("postgres://"):
        cleaned = "postgresql://" + cleaned[len("postgres://"):]

    parsed = urlparse(cleaned)
    query_params = dict(parse_qsl(parsed.query, keep_blank_values=True))
    if "sslmode" not in query_params:
        query_params["sslmode"] = "require"
    if "connect_timeout" not in query_params:
        query_params["connect_timeout"] = "10"

    new_query = urlencode(query_params)
    return urlunparse((
        parsed.scheme,
        parsed.netloc,
        parsed.path,
        parsed.params,
        new_query,
        parsed.fragment,
    ))


def is_postgres_enabled() -> bool:
    url = _get_database_url()
    return bool(url and (url.startswith("postgres://") or url.startswith("postgresql://")))


def _adapt_sql_for_postgres(sql: str) -> str:
    """Converte placeholders '?' do SQLite para '%s' do PostgreSQL."""
    return sql.replace("?", "%s")


class DBWrapper:
    """Interface unificada entre PostgreSQL (Supabase) e SQLite local."""

    def __init__(self, conn: Any, dialect: str):
        self.conn = conn
        self.dialect = dialect

    def execute(self, sql: str, params: Tuple[Any, ...] = ()) -> Any:
        if self.dialect == "postgres":
            pg_sql = _adapt_sql_for_postgres(sql)
            cur = self.conn.cursor(cursor_factory=RealDictCursor)
            cur.execute(pg_sql, params)
            return cur
        else:
            cur = self.conn.cursor()
            cur.execute(sql, params)
            return cur

    def fetchone(self, sql: str, params: Tuple[Any, ...] = ()) -> Optional[Dict[str, Any]]:
        cur = self.execute(sql, params)
        row = cur.fetchone()
        if row is None:
            return None
        return dict(row)

    def fetchall(self, sql: str, params: Tuple[Any, ...] = ()) -> List[Dict[str, Any]]:
        cur = self.execute(sql, params)
        rows = cur.fetchall()
        return [dict(r) for r in rows]

    def insert_returning_id(self, sql: str, params: Tuple[Any, ...] = ()) -> int:
        if self.dialect == "postgres":
            pg_sql = _adapt_sql_for_postgres(sql).rstrip().rstrip(";") + " RETURNING id"
            cur = self.conn.cursor(cursor_factory=RealDictCursor)
            cur.execute(pg_sql, params)
            row = cur.fetchone()
            return int(row["id"])
        else:
            cur = self.execute(sql, params)
            return int(cur.lastrowid)

    def commit(self) -> None:
        self.conn.commit()

    def rollback(self) -> None:
        self.conn.rollback()


@contextmanager
def get_connection(db_path: str = DEFAULT_DB_PATH):
    """Abre conexão com PostgreSQL (se DATABASE_URL configurada) ou SQLite local."""
    if is_postgres_enabled():
        if psycopg2 is None:
            raise RuntimeError(
                "DATABASE_URL configurada, mas 'psycopg2-binary' não está instalado. "
                "Adicione psycopg2-binary ao requirements.txt."
            )
        db_url = _normalize_postgres_url(_get_database_url())
        raw_conn = psycopg2.connect(db_url)
        wrapper = DBWrapper(raw_conn, "postgres")
        try:
            yield wrapper
            wrapper.commit()
        except Exception:
            wrapper.rollback()
            raise
        finally:
            raw_conn.close()
    else:
        os.makedirs(os.path.dirname(db_path) or ".", exist_ok=True)
        raw_conn = sqlite3.connect(db_path, check_same_thread=False)
        raw_conn.row_factory = sqlite3.Row
        raw_conn.execute("PRAGMA foreign_keys = ON;")
        wrapper = DBWrapper(raw_conn, "sqlite")
        try:
            yield wrapper
            wrapper.commit()
        except Exception:
            wrapper.rollback()
            raise
        finally:
            raw_conn.close()


def _column_exists(db: DBWrapper, table_name: str, column_name: str) -> bool:
    if db.dialect == "postgres":
        row = db.fetchone(
            """
            SELECT 1 AS ok
            FROM information_schema.columns
            WHERE table_schema = 'public'
              AND table_name = ?
              AND column_name = ?
            """,
            (table_name, column_name),
        )
        return row is not None
    else:
        rows = db.fetchall(f"PRAGMA table_info({table_name})")
        return any(r["name"] == column_name for r in rows)


def init_db(db_path: str = DEFAULT_DB_PATH) -> None:
    """Cria tabelas, aplica migrações e insere brindes padrão se o catálogo estiver vazio."""
    with get_connection(db_path) as db:
        if db.dialect == "postgres":
            db.execute("""
                CREATE TABLE IF NOT EXISTS items (
                    id SERIAL PRIMARY KEY,
                    name TEXT NOT NULL UNIQUE,
                    code TEXT NOT NULL,
                    unit_price DOUBLE PRECISION NOT NULL DEFAULT 0.0,
                    stock_qty INTEGER NOT NULL DEFAULT 0,
                    min_stock INTEGER NOT NULL DEFAULT 5,
                    image_path TEXT DEFAULT '',
                    active INTEGER NOT NULL DEFAULT 1,
                    created_at TEXT NOT NULL
                )
            """)
            db.execute("""
                CREATE TABLE IF NOT EXISTS recipes (
                    id SERIAL PRIMARY KEY,
                    name TEXT NOT NULL UNIQUE,
                    description TEXT DEFAULT '',
                    created_at TEXT NOT NULL
                )
            """)
            db.execute("""
                CREATE TABLE IF NOT EXISTS recipe_items (
                    id SERIAL PRIMARY KEY,
                    recipe_id INTEGER NOT NULL REFERENCES recipes(id) ON DELETE CASCADE,
                    item_id INTEGER NOT NULL REFERENCES items(id) ON DELETE CASCADE,
                    quantity INTEGER NOT NULL DEFAULT 1
                )
            """)
            db.execute("""
                CREATE TABLE IF NOT EXISTS kits (
                    id SERIAL PRIMARY KEY,
                    recipient_name TEXT NOT NULL,
                    recipient_address TEXT NOT NULL DEFAULT '',
                    notes TEXT DEFAULT '',
                    total_value DOUBLE PRECISION NOT NULL DEFAULT 0.0,
                    created_at TEXT NOT NULL
                )
            """)
            db.execute("""
                CREATE TABLE IF NOT EXISTS kit_items (
                    id SERIAL PRIMARY KEY,
                    kit_id INTEGER NOT NULL REFERENCES kits(id) ON DELETE CASCADE,
                    item_id INTEGER REFERENCES items(id) ON DELETE SET NULL,
                    item_name TEXT NOT NULL,
                    item_code TEXT NOT NULL,
                    unit_price DOUBLE PRECISION NOT NULL DEFAULT 0.0,
                    quantity INTEGER NOT NULL DEFAULT 1
                )
            """)
            db.execute("""
                CREATE TABLE IF NOT EXISTS stock_movements (
                    id SERIAL PRIMARY KEY,
                    item_id INTEGER REFERENCES items(id) ON DELETE SET NULL,
                    item_name TEXT NOT NULL,
                    movement_type TEXT NOT NULL,
                    qty_change INTEGER NOT NULL,
                    balance_after INTEGER NOT NULL,
                    reference_info TEXT DEFAULT '',
                    created_at TEXT NOT NULL
                )
            """)
        else:
            db.execute("""
                CREATE TABLE IF NOT EXISTS items (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    name TEXT NOT NULL UNIQUE,
                    code TEXT NOT NULL,
                    unit_price REAL NOT NULL DEFAULT 0.0,
                    stock_qty INTEGER NOT NULL DEFAULT 0,
                    min_stock INTEGER NOT NULL DEFAULT 5,
                    image_path TEXT DEFAULT '',
                    active INTEGER NOT NULL DEFAULT 1,
                    created_at TEXT NOT NULL
                )
            """)
            db.execute("""
                CREATE TABLE IF NOT EXISTS recipes (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    name TEXT NOT NULL UNIQUE,
                    description TEXT DEFAULT '',
                    created_at TEXT NOT NULL
                )
            """)
            db.execute("""
                CREATE TABLE IF NOT EXISTS recipe_items (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    recipe_id INTEGER NOT NULL,
                    item_id INTEGER NOT NULL,
                    quantity INTEGER NOT NULL DEFAULT 1,
                    FOREIGN KEY (recipe_id) REFERENCES recipes(id) ON DELETE CASCADE,
                    FOREIGN KEY (item_id) REFERENCES items(id) ON DELETE CASCADE
                )
            """)
            db.execute("""
                CREATE TABLE IF NOT EXISTS kits (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    recipient_name TEXT NOT NULL,
                    recipient_address TEXT NOT NULL DEFAULT '',
                    notes TEXT DEFAULT '',
                    total_value REAL NOT NULL DEFAULT 0.0,
                    created_at TEXT NOT NULL
                )
            """)
            db.execute("""
                CREATE TABLE IF NOT EXISTS kit_items (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    kit_id INTEGER NOT NULL,
                    item_id INTEGER,
                    item_name TEXT NOT NULL,
                    item_code TEXT NOT NULL,
                    unit_price REAL NOT NULL DEFAULT 0.0,
                    quantity INTEGER NOT NULL DEFAULT 1,
                    FOREIGN KEY (kit_id) REFERENCES kits(id) ON DELETE CASCADE,
                    FOREIGN KEY (item_id) REFERENCES items(id) ON DELETE SET NULL
                )
            """)
            db.execute("""
                CREATE TABLE IF NOT EXISTS stock_movements (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    item_id INTEGER,
                    item_name TEXT NOT NULL,
                    movement_type TEXT NOT NULL,
                    qty_change INTEGER NOT NULL,
                    balance_after INTEGER NOT NULL,
                    reference_info TEXT DEFAULT '',
                    created_at TEXT NOT NULL,
                    FOREIGN KEY (item_id) REFERENCES items(id) ON DELETE SET NULL
                )
            """)

        # Migrações incrementais seguras caso tabelas antigas existam
        if not _column_exists(db, "items", "stock_qty"):
            db.execute("ALTER TABLE items ADD COLUMN stock_qty INTEGER NOT NULL DEFAULT 0")
        if not _column_exists(db, "items", "min_stock"):
            db.execute("ALTER TABLE items ADD COLUMN min_stock INTEGER NOT NULL DEFAULT 5")
        if not _column_exists(db, "kits", "recipient_address"):
            db.execute("ALTER TABLE kits ADD COLUMN recipient_address TEXT NOT NULL DEFAULT ''")

        # Seed inicial caso não haja brindes cadastrados
        count_row = db.fetchone("SELECT COUNT(*) AS total FROM items")
        if count_row and int(count_row["total"]) == 0:
            now_str = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            default_items = [
                ("Garrafa Térmica", "GAR", 45.00, 50, 10, ""),
                ("Blusa Corporativa", "BLU", 68.90, 40, 8, ""),
                ("Chaveiro Metálico", "CHV", 12.50, 100, 15, ""),
                ("Caneca Cerâmica", "CAN", 29.90, 60, 12, ""),
                ("Caderno Capa Dura", "CAD", 34.00, 45, 10, ""),
            ]
            for name, code, price, stock, min_s, img in default_items:
                item_id = db.insert_returning_id(
                    """
                    INSERT INTO items (name, code, unit_price, stock_qty, min_stock, image_path, active, created_at)
                    VALUES (?, ?, ?, ?, ?, ?, 1, ?)
                    """,
                    (name, code, price, stock, min_s, img, now_str),
                )
                db.execute(
                    """
                    INSERT INTO stock_movements
                    (item_id, item_name, movement_type, qty_change, balance_after, reference_info, created_at)
                    VALUES (?, ?, 'ENTRADA_INICIAL', ?, ?, 'Carga inicial do catálogo', ?)
                    """,
                    (item_id, name, stock, stock, now_str),
                )


def generate_code_from_name(name: str) -> str:
    clean = re.sub(r"[^A-Za-zÀ-ÿ0-9]", "", name.strip()).upper()
    return (clean[:3] if len(clean) >= 3 else clean.ljust(3, "X"))


# ==============================================================================
# CRUD DE BRINDES E INVENTÁRIO
# ==============================================================================

def list_items(active_only: bool = True, db_path: str = DEFAULT_DB_PATH) -> List[Dict[str, Any]]:
    with get_connection(db_path) as db:
        if active_only:
            return db.fetchall("SELECT * FROM items WHERE active = 1 ORDER BY name ASC")
        return db.fetchall("SELECT * FROM items ORDER BY active DESC, name ASC")


def get_item(item_id: int, db_path: str = DEFAULT_DB_PATH) -> Optional[Dict[str, Any]]:
    with get_connection(db_path) as db:
        return db.fetchone("SELECT * FROM items WHERE id = ?", (item_id,))


def add_item(
    name: str,
    code: str,
    unit_price: float,
    stock_qty: int = 0,
    min_stock: int = 5,
    image_path: str = "",
    db_path: str = DEFAULT_DB_PATH,
) -> int:
    clean_name = name.strip()
    clean_code = (code.strip().upper() or generate_code_from_name(clean_name))[:6]
    now_str = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    with get_connection(db_path) as db:
        existing = db.fetchone("SELECT id FROM items WHERE LOWER(name) = LOWER(?)", (clean_name,))
        if existing:
            db.execute(
                """
                UPDATE items
                SET code = ?, unit_price = ?, stock_qty = ?, min_stock = ?,
                    image_path = CASE WHEN ? <> '' THEN ? ELSE image_path END,
                    active = 1
                WHERE id = ?
                """,
                (clean_code, float(unit_price), int(stock_qty), int(min_stock), image_path, image_path, existing["id"]),
            )
            return int(existing["id"])

        item_id = db.insert_returning_id(
            """
            INSERT INTO items (name, code, unit_price, stock_qty, min_stock, image_path, active, created_at)
            VALUES (?, ?, ?, ?, ?, ?, 1, ?)
            """,
            (clean_name, clean_code, float(unit_price), int(stock_qty), int(min_stock), image_path, now_str),
        )
        if int(stock_qty) != 0:
            db.execute(
                """
                INSERT INTO stock_movements
                (item_id, item_name, movement_type, qty_change, balance_after, reference_info, created_at)
                VALUES (?, ?, 'CADASTRO', ?, ?, 'Saldo inicial ao cadastrar brinde', ?)
                """,
                (item_id, clean_name, int(stock_qty), int(stock_qty), now_str),
            )
        return item_id


def update_item(
    item_id: int,
    name: str,
    code: str,
    unit_price: float,
    stock_qty: int,
    min_stock: int,
    image_path: Optional[str] = None,
    db_path: str = DEFAULT_DB_PATH,
) -> None:
    clean_name = name.strip()
    clean_code = (code.strip().upper() or generate_code_from_name(clean_name))[:6]
    now_str = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    with get_connection(db_path) as db:
        current = db.fetchone("SELECT * FROM items WHERE id = ?", (item_id,))
        if not current:
            return
        old_stock = int(current["stock_qty"])
        new_stock = int(stock_qty)
        diff = new_stock - old_stock

        if image_path is not None:
            db.execute(
                """
                UPDATE items
                SET name = ?, code = ?, unit_price = ?, stock_qty = ?, min_stock = ?, image_path = ?
                WHERE id = ?
                """,
                (clean_name, clean_code, float(unit_price), new_stock, int(min_stock), image_path, item_id),
            )
        else:
            db.execute(
                """
                UPDATE items
                SET name = ?, code = ?, unit_price = ?, stock_qty = ?, min_stock = ?
                WHERE id = ?
                """,
                (clean_name, clean_code, float(unit_price), new_stock, int(min_stock), item_id),
            )

        if diff != 0:
            db.execute(
                """
                INSERT INTO stock_movements
                (item_id, item_name, movement_type, qty_change, balance_after, reference_info, created_at)
                VALUES (?, ?, 'AJUSTE_MANUAL', ?, ?, 'Alteração direta no cadastro/inventário', ?)
                """,
                (item_id, clean_name, diff, new_stock, now_str),
            )


def adjust_item_stock(
    item_id: int,
    delta: int,
    reason: str = "Reposição/ajuste de estoque",
    db_path: str = DEFAULT_DB_PATH,
) -> int:
    now_str = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    with get_connection(db_path) as db:
        item = db.fetchone("SELECT * FROM items WHERE id = ?", (item_id,))
        if not item:
            raise ValueError("Brinde não encontrado.")
        new_balance = max(0, int(item["stock_qty"]) + int(delta))
        actual_delta = new_balance - int(item["stock_qty"])
        db.execute("UPDATE items SET stock_qty = ? WHERE id = ?", (new_balance, item_id))
        if actual_delta != 0:
            mov_type = "REPOSICAO" if actual_delta > 0 else "SAIDA_MANUAL"
            db.execute(
                """
                INSERT INTO stock_movements
                (item_id, item_name, movement_type, qty_change, balance_after, reference_info, created_at)
                VALUES (?, ?, ?, ?, ?, ?, ?)
                """,
                (item_id, item["name"], mov_type, actual_delta, new_balance, reason, now_str),
            )
        return new_balance


def delete_item(item_id: int, hard_delete: bool = False, db_path: str = DEFAULT_DB_PATH) -> None:
    with get_connection(db_path) as db:
        if hard_delete:
            db.execute("DELETE FROM items WHERE id = ?", (item_id,))
        else:
            db.execute("UPDATE items SET active = 0 WHERE id = ?", (item_id,))


def list_stock_movements(limit: int = 100, db_path: str = DEFAULT_DB_PATH) -> List[Dict[str, Any]]:
    with get_connection(db_path) as db:
        return db.fetchall(
            "SELECT * FROM stock_movements ORDER BY id DESC LIMIT ?",
            (int(limit),),
        )


# ==============================================================================
# CRUD DE RECEITAS (KITS PADRÃO)
# ==============================================================================

def list_recipes(db_path: str = DEFAULT_DB_PATH) -> List[Dict[str, Any]]:
    with get_connection(db_path) as db:
        recipes = db.fetchall("SELECT * FROM recipes ORDER BY name ASC")
        for rec in recipes:
            rec["items"] = db.fetchall(
                """
                SELECT ri.id, ri.item_id, ri.quantity, i.name AS item_name, i.code AS item_code,
                       i.unit_price, i.stock_qty, i.image_path
                FROM recipe_items ri
                JOIN items i ON i.id = ri.item_id
                WHERE ri.recipe_id = ? AND i.active = 1
                ORDER BY i.name ASC
                """,
                (rec["id"],),
            )
        return recipes


def save_recipe(
    name: str,
    description: str,
    items_qty: Dict[int, int],
    recipe_id: Optional[int] = None,
    db_path: str = DEFAULT_DB_PATH,
) -> int:
    clean_name = name.strip()
    now_str = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    with get_connection(db_path) as db:
        if recipe_id is not None:
            db.execute(
                "UPDATE recipes SET name = ?, description = ? WHERE id = ?",
                (clean_name, description.strip(), int(recipe_id)),
            )
            db.execute("DELETE FROM recipe_items WHERE recipe_id = ?", (int(recipe_id),))
            target_id = int(recipe_id)
        else:
            existing = db.fetchone("SELECT id FROM recipes WHERE LOWER(name) = LOWER(?)", (clean_name,))
            if existing:
                target_id = int(existing["id"])
                db.execute(
                    "UPDATE recipes SET description = ? WHERE id = ?",
                    (description.strip(), target_id),
                )
                db.execute("DELETE FROM recipe_items WHERE recipe_id = ?", (target_id,))
            else:
                target_id = db.insert_returning_id(
                    "INSERT INTO recipes (name, description, created_at) VALUES (?, ?, ?)",
                    (clean_name, description.strip(), now_str),
                )

        for item_id, qty in items_qty.items():
            if int(qty) > 0:
                db.execute(
                    "INSERT INTO recipe_items (recipe_id, item_id, quantity) VALUES (?, ?, ?)",
                    (target_id, int(item_id), int(qty)),
                )
        return target_id


def delete_recipe(recipe_id: int, db_path: str = DEFAULT_DB_PATH) -> None:
    with get_connection(db_path) as db:
        db.execute("DELETE FROM recipes WHERE id = ?", (int(recipe_id),))


# ==============================================================================
# CRUD DE KITS FINALIZADOS COM BAIXA E DEVOLUÇÃO AUTOMÁTICA DE ESTOQUE
# ==============================================================================

def _deduct_stock_for_kit(
    db: DBWrapper,
    kit_id: int,
    recipient_name: str,
    selected_items: List[Dict[str, Any]],
    now_str: str,
) -> None:
    for entry in selected_items:
        item_id = entry.get("item_id")
        qty = int(entry.get("quantity", 0))
        if item_id is None or qty <= 0:
            continue
        item_row = db.fetchone("SELECT * FROM items WHERE id = ?", (int(item_id),))
        if not item_row:
            continue
        current_stock = int(item_row["stock_qty"])
        new_stock = max(0, current_stock - qty)
        actual_deducted = new_stock - current_stock  # negativo
        db.execute("UPDATE items SET stock_qty = ? WHERE id = ?", (new_stock, int(item_id)))
        db.execute(
            """
            INSERT INTO stock_movements
            (item_id, item_name, movement_type, qty_change, balance_after, reference_info, created_at)
            VALUES (?, ?, 'BAIXA_KIT', ?, ?, ?, ?)
            """,
            (
                int(item_id),
                item_row["name"],
                actual_deducted,
                new_stock,
                f"Kit #{kit_id} - {recipient_name}",
                now_str,
            ),
        )


def _restore_stock_for_kit(db: DBWrapper, kit_id: int, recipient_name: str, now_str: str) -> None:
    old_items = db.fetchall("SELECT * FROM kit_items WHERE kit_id = ?", (int(kit_id),))
    for old in old_items:
        item_id = old.get("item_id")
        qty = int(old.get("quantity", 0))
        if item_id is None or qty <= 0:
            continue
        item_row = db.fetchone("SELECT * FROM items WHERE id = ?", (int(item_id),))
        if not item_row:
            continue
        new_stock = int(item_row["stock_qty"]) + qty
        db.execute("UPDATE items SET stock_qty = ? WHERE id = ?", (new_stock, int(item_id)))
        db.execute(
            """
            INSERT INTO stock_movements
            (item_id, item_name, movement_type, qty_change, balance_after, reference_info, created_at)
            VALUES (?, ?, 'ESTORNO_KIT', ?, ?, ?, ?)
            """,
            (
                int(item_id),
                item_row["name"],
                qty,
                new_stock,
                f"Estorno Kit #{kit_id} - {recipient_name}",
                now_str,
            ),
        )


def create_kit(
    recipient_name: str,
    recipient_address: str,
    notes: str,
    selected_items: List[Dict[str, Any]],
    deduct_stock: bool = True,
    db_path: str = DEFAULT_DB_PATH,
) -> int:
    now_str = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    total_value = round(
        sum(float(i["unit_price"]) * int(i["quantity"]) for i in selected_items if int(i["quantity"]) > 0),
        2,
    )
    with get_connection(db_path) as db:
        kit_id = db.insert_returning_id(
            """
            INSERT INTO kits (recipient_name, recipient_address, notes, total_value, created_at)
            VALUES (?, ?, ?, ?, ?)
            """,
            (recipient_name.strip(), recipient_address.strip(), notes.strip(), total_value, now_str),
        )
        for entry in selected_items:
            qty = int(entry["quantity"])
            if qty <= 0:
                continue
            db.execute(
                """
                INSERT INTO kit_items (kit_id, item_id, item_name, item_code, unit_price, quantity)
                VALUES (?, ?, ?, ?, ?, ?)
                """,
                (
                    kit_id,
                    entry.get("item_id"),
                    entry["item_name"],
                    entry["item_code"],
                    float(entry["unit_price"]),
                    qty,
                ),
            )
        if deduct_stock:
            _deduct_stock_for_kit(db, kit_id, recipient_name.strip(), selected_items, now_str)
        return kit_id


def update_kit(
    kit_id: int,
    recipient_name: str,
    recipient_address: str,
    notes: str,
    selected_items: List[Dict[str, Any]],
    db_path: str = DEFAULT_DB_PATH,
) -> None:
    now_str = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    total_value = round(
        sum(float(i["unit_price"]) * int(i["quantity"]) for i in selected_items if int(i["quantity"]) > 0),
        2,
    )
    with get_connection(db_path) as db:
        existing = db.fetchone("SELECT * FROM kits WHERE id = ?", (int(kit_id),))
        if not existing:
            return
        # Devolve saldo anterior ao estoque e aplica o novo
        _restore_stock_for_kit(db, int(kit_id), existing["recipient_name"], now_str)
        db.execute("DELETE FROM kit_items WHERE kit_id = ?", (int(kit_id),))
        db.execute(
            """
            UPDATE kits
            SET recipient_name = ?, recipient_address = ?, notes = ?, total_value = ?
            WHERE id = ?
            """,
            (recipient_name.strip(), recipient_address.strip(), notes.strip(), total_value, int(kit_id)),
        )
        for entry in selected_items:
            qty = int(entry["quantity"])
            if qty <= 0:
                continue
            db.execute(
                """
                INSERT INTO kit_items (kit_id, item_id, item_name, item_code, unit_price, quantity)
                VALUES (?, ?, ?, ?, ?, ?)
                """,
                (
                    int(kit_id),
                    entry.get("item_id"),
                    entry["item_name"],
                    entry["item_code"],
                    float(entry["unit_price"]),
                    qty,
                ),
            )
        _deduct_stock_for_kit(db, int(kit_id), recipient_name.strip(), selected_items, now_str)


def delete_kit(kit_id: int, restore_stock: bool = True, db_path: str = DEFAULT_DB_PATH) -> None:
    now_str = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    with get_connection(db_path) as db:
        existing = db.fetchone("SELECT * FROM kits WHERE id = ?", (int(kit_id),))
        if not existing:
            return
        if restore_stock:
            _restore_stock_for_kit(db, int(kit_id), existing["recipient_name"], now_str)
        db.execute("DELETE FROM kits WHERE id = ?", (int(kit_id),))


def clear_all_kits(restore_stock: bool = True, db_path: str = DEFAULT_DB_PATH) -> None:
    now_str = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    with get_connection(db_path) as db:
        kits = db.fetchall("SELECT id, recipient_name FROM kits")
        if restore_stock:
            for k in kits:
                _restore_stock_for_kit(db, int(k["id"]), k["recipient_name"], now_str)
        db.execute("DELETE FROM kit_items")
        db.execute("DELETE FROM kits")


def list_kits(search_query: str = "", db_path: str = DEFAULT_DB_PATH) -> List[Dict[str, Any]]:
    with get_connection(db_path) as db:
        q = search_query.strip()
        if q:
            like = f"%{q}%"
            kits = db.fetchall(
                """
                SELECT DISTINCT k.*
                FROM kits k
                LEFT JOIN kit_items ki ON ki.kit_id = k.id
                WHERE LOWER(k.recipient_name) LIKE LOWER(?)
                   OR LOWER(k.recipient_address) LIKE LOWER(?)
                   OR LOWER(k.notes) LIKE LOWER(?)
                   OR LOWER(ki.item_name) LIKE LOWER(?)
                ORDER BY k.id DESC
                """,
                (like, like, like, like),
            )
        else:
            kits = db.fetchall("SELECT * FROM kits ORDER BY id DESC")

        for kit in kits:
            kit["items"] = db.fetchall(
                "SELECT * FROM kit_items WHERE kit_id = ? ORDER BY id ASC",
                (kit["id"],),
            )
        return kits