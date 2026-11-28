import asyncio
import csv
import io
import logging
import os
import re
import secrets
import sqlite3
from datetime import datetime, timedelta
from typing import Optional

from aiogram import Bot, Dispatcher, F, types
from aiogram.filters import CommandObject, CommandStart
from aiogram.fsm.context import FSMContext
from aiogram.fsm.state import State, StatesGroup
from aiogram.fsm.storage.memory import MemoryStorage
from aiogram.types import (
    BufferedInputFile,
    InlineKeyboardButton,
    InlineKeyboardMarkup,
    KeyboardButton,
    ReplyKeyboardMarkup,
)

try:
    import openpyxl
    from openpyxl.styles import PatternFill

    HAS_OPENPYXL = True
except ImportError:
    HAS_OPENPYXL = False

# --- НАСТРОЙКИ ---
BOT_TOKEN = os.getenv("BOT_TOKEN", "8763607619:AAGcLqQjNSkjCPYmFcjE_HhC458D_HrUyig")
DB_FILE = os.getenv("DB_FILE", "sklad_data.db")
ADMIN_IDS: set[int] = {
    int(x.strip())
    for x in os.getenv("ADMIN_IDS", "363447845").split(",")
    if x.strip().isdigit()
}
LOW_STOCK_THRESHOLD = 0.0
BOT_VERSION = os.getenv("BOT_VERSION", "2026.11.28-full-kb-refresh")

logging.basicConfig(level=logging.INFO)
bot = Bot(token=BOT_TOKEN)
dp = Dispatcher(storage=MemoryStorage())


def get_conn() -> sqlite3.Connection:
    conn = sqlite3.connect(DB_FILE, timeout=30.0)
    conn.row_factory = sqlite3.Row
    conn.execute("PRAGMA journal_mode=WAL;")
    return conn


def init_db():
    conn = get_conn()
    cur = conn.cursor()

    cur.execute(
        """
        CREATE TABLE IF NOT EXISTS workspaces (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            title TEXT NOT NULL,
            owner_tg_id INTEGER NOT NULL,
            created_at TEXT NOT NULL
        )
        """
    )

    cur.execute(
        """
        CREATE TABLE IF NOT EXISTS users (
            tg_id INTEGER PRIMARY KEY,
            username TEXT NOT NULL DEFAULT '',
            full_name TEXT NOT NULL DEFAULT '',
            active_workspace_id INTEGER,
            FOREIGN KEY (active_workspace_id) REFERENCES workspaces(id)
        )
        """
    )

    cur.execute(
        """
        CREATE TABLE IF NOT EXISTS workspace_members (
            workspace_id INTEGER NOT NULL,
            tg_id INTEGER NOT NULL,
            role TEXT NOT NULL DEFAULT 'member',
            PRIMARY KEY (workspace_id, tg_id),
            FOREIGN KEY (workspace_id) REFERENCES workspaces(id),
            FOREIGN KEY (tg_id) REFERENCES users(tg_id)
        )
        """
    )

    cur.execute(
        """
        CREATE TABLE IF NOT EXISTS workspace_invites (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            code TEXT NOT NULL UNIQUE,
            workspace_id INTEGER NOT NULL,
            created_by_tg_id INTEGER NOT NULL,
            is_active INTEGER NOT NULL DEFAULT 1,
            created_at TEXT NOT NULL,
            FOREIGN KEY (workspace_id) REFERENCES workspaces(id)
        )
        """
    )

    cur.execute(
        """
        CREATE TABLE IF NOT EXISTS pending_invites (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            workspace_id INTEGER NOT NULL,
            username_norm TEXT NOT NULL,
            invited_by_tg_id INTEGER NOT NULL,
            created_at TEXT NOT NULL,
            UNIQUE(workspace_id, username_norm),
            FOREIGN KEY (workspace_id) REFERENCES workspaces(id)
        )
        """
    )

    cur.execute(
        """
        CREATE TABLE IF NOT EXISTS catalog_items (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            workspace_id INTEGER NOT NULL,
            name TEXT NOT NULL,
            unit TEXT NOT NULL,
            category TEXT NOT NULL,
            UNIQUE(workspace_id, name),
            FOREIGN KEY (workspace_id) REFERENCES workspaces(id)
        )
        """
    )

    cur.execute(
        """
        CREATE TABLE IF NOT EXISTS movements (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            workspace_id INTEGER NOT NULL,
            op_date TEXT NOT NULL,
            op_type TEXT NOT NULL,
            item_name TEXT NOT NULL,
            quantity REAL NOT NULL,
            unit TEXT NOT NULL,
            category TEXT NOT NULL,
            comment TEXT NOT NULL,
            actor_name TEXT NOT NULL,
            actor_tg_id INTEGER NOT NULL,
            created_at TEXT NOT NULL,
            FOREIGN KEY (workspace_id) REFERENCES workspaces(id)
        )
        """
    )

    cur.execute(
        """
        CREATE TABLE IF NOT EXISTS cancelled_movements (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            orig_movement_id INTEGER,
            workspace_id INTEGER NOT NULL,
            op_date TEXT NOT NULL,
            op_type TEXT NOT NULL,
            item_name TEXT NOT NULL,
            quantity REAL NOT NULL,
            unit TEXT NOT NULL,
            category TEXT NOT NULL,
            comment TEXT NOT NULL,
            actor_name TEXT NOT NULL,
            actor_tg_id INTEGER NOT NULL,
            cancelled_by_tg_id INTEGER NOT NULL,
            cancelled_at TEXT NOT NULL
        )
        """
    )

    user_cols = [r["name"] for r in cur.execute("PRAGMA table_info(users)").fetchall()]
    if "workspace_id" in user_cols and "active_workspace_id" not in user_cols:
        cur.execute("ALTER TABLE users ADD COLUMN active_workspace_id INTEGER")
        cur.execute("UPDATE users SET active_workspace_id = workspace_id WHERE workspace_id IS NOT NULL")

    cur.execute(
        """
        INSERT OR IGNORE INTO workspace_members (workspace_id, tg_id, role)
        SELECT id, owner_tg_id, 'owner' FROM workspaces
        """
    )
    cur.execute(
        """
        INSERT OR IGNORE INTO workspace_members (workspace_id, tg_id, role)
        SELECT active_workspace_id, tg_id, 'member'
        FROM users
        WHERE active_workspace_id IS NOT NULL
        """
    )

    conn.commit()
    conn.close()


async def db_upsert_user(tg_id: int, username: Optional[str], full_name: str):
    uname_clean = (username or "").strip().lstrip("@")
    uname_norm = uname_clean.lower()

    conn = get_conn()
    cur = conn.cursor()
    cur.execute(
        """
        INSERT INTO users (tg_id, username, full_name)
        VALUES (?, ?, ?)
        ON CONFLICT(tg_id) DO UPDATE SET
            username = excluded.username,
            full_name = excluded.full_name
        """,
        (tg_id, uname_clean, full_name),
    )

    applied_workspaces: list[tuple[int, str]] = []
    if uname_norm:
        invites = cur.execute(
            """
            SELECT p.id, p.workspace_id, w.title
            FROM pending_invites p
            JOIN workspaces w ON w.id = p.workspace_id
            WHERE p.username_norm = ?
            """,
            (uname_norm,),
        ).fetchall()

        for inv in invites:
            ws_id = int(inv["workspace_id"])
            cur.execute(
                """
                INSERT OR IGNORE INTO workspace_members (workspace_id, tg_id, role)
                VALUES (?, ?, 'member')
                """,
                (ws_id, tg_id),
            )
            cur.execute(
                """
                UPDATE users
                SET active_workspace_id = COALESCE(active_workspace_id, ?)
                WHERE tg_id = ?
                """,
                (ws_id, tg_id),
            )
            cur.execute("DELETE FROM pending_invites WHERE id = ?", (inv["id"],))
            applied_workspaces.append((ws_id, inv["title"]))

    conn.commit()
    conn.close()

    for ws_id, ws_title in applied_workspaces:
        try:
            await bot.send_message(
                tg_id,
                f"✅ Вас подключили к складу **«{ws_title}»**.",
                parse_mode="Markdown",
            )
        except Exception as e:
            logging.warning(f"Не удалось уведомить пользователя {tg_id} о подключении к складу {ws_id}: {e}")


def db_get_or_create_invite(workspace_id: int, creator_tg_id: int, rotate: bool = False) -> str:
    conn = get_conn()
    cur = conn.cursor()
    if rotate:
        cur.execute(
            "UPDATE workspace_invites SET is_active = 0 WHERE workspace_id = ? AND is_active = 1",
            (workspace_id,),
        )
    else:
        row = cur.execute(
            """
            SELECT code FROM workspace_invites
            WHERE workspace_id = ? AND is_active = 1
            ORDER BY id DESC LIMIT 1
            """,
            (workspace_id,),
        ).fetchone()
        if row:
            conn.close()
            return str(row["code"])

    code = f"inv_{secrets.token_urlsafe(8)}"
    now_str = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    cur.execute(
        """
        INSERT INTO workspace_invites (code, workspace_id, created_by_tg_id, is_active, created_at)
        VALUES (?, ?, ?, 1, ?)
        """,
        (code, workspace_id, creator_tg_id, now_str),
    )
    conn.commit()
    conn.close()
    return code


def db_accept_invite_code(code: str, tg_id: int) -> tuple[str, Optional[dict]]:
    conn = get_conn()
    cur = conn.cursor()
    inv = cur.execute(
        """
        SELECT wi.id, wi.workspace_id, wi.created_by_tg_id, wi.is_active, w.title, w.owner_tg_id
        FROM workspace_invites wi
        JOIN workspaces w ON w.id = wi.workspace_id
        WHERE wi.code = ?
        """,
        (code,),
    ).fetchone()

    if not inv:
        conn.close()
        return "not_found", None
    if int(inv["is_active"]) != 1:
        conn.close()
        return "inactive", None

    ws_id = int(inv["workspace_id"])
    ws_title = str(inv["title"])
    owner_tg_id = int(inv["owner_tg_id"])

    existing = cur.execute(
        "SELECT role FROM workspace_members WHERE workspace_id = ? AND tg_id = ?",
        (ws_id, tg_id),
    ).fetchone()

    cur.execute(
        """
        INSERT OR IGNORE INTO workspace_members (workspace_id, tg_id, role)
        VALUES (?, ?, 'member')
        """,
        (ws_id, tg_id),
    )
    cur.execute(
        """
        INSERT INTO users (tg_id, active_workspace_id)
        VALUES (?, ?)
        ON CONFLICT(tg_id) DO UPDATE SET active_workspace_id = excluded.active_workspace_id
        """,
        (tg_id, ws_id),
    )
    conn.commit()
    conn.close()

    status = "already_member" if existing else "joined"
    return status, {"id": ws_id, "title": ws_title, "owner_tg_id": owner_tg_id}


def db_get_user_workspaces(tg_id: int) -> list[dict]:
    conn = get_conn()
    cur = conn.cursor()
    rows = cur.execute(
        """
        SELECT w.id, w.title, w.owner_tg_id, wm.role
        FROM workspace_members wm
        JOIN workspaces w ON w.id = wm.workspace_id
        WHERE wm.tg_id = ?
        ORDER BY w.id ASC
        """,
        (tg_id,),
    ).fetchall()
    conn.close()
    return [
        {
            "id": int(r["id"]),
            "title": r["title"],
            "owner_tg_id": int(r["owner_tg_id"]),
            "role": r["role"],
        }
        for r in rows
    ]


def db_get_active_workspace(tg_id: int) -> Optional[dict]:
    conn = get_conn()
    cur = conn.cursor()
    row = cur.execute(
        """
        SELECT w.id, w.title, w.owner_tg_id, wm.role
        FROM users u
        JOIN workspaces w ON w.id = u.active_workspace_id
        LEFT JOIN workspace_members wm ON wm.workspace_id = w.id AND wm.tg_id = u.tg_id
        WHERE u.tg_id = ?
        """,
        (tg_id,),
    ).fetchone()

    if row:
        conn.close()
        return {
            "id": int(row["id"]),
            "title": row["title"],
            "owner_tg_id": int(row["owner_tg_id"]),
            "role": row["role"] or "member",
        }

    fallback = cur.execute(
        """
        SELECT w.id, w.title, w.owner_tg_id, wm.role
        FROM workspace_members wm
        JOIN workspaces w ON w.id = wm.workspace_id
        WHERE wm.tg_id = ?
        ORDER BY w.id ASC
        LIMIT 1
        """,
        (tg_id,),
    ).fetchone()
    if fallback:
        ws_id = int(fallback["id"])
        cur.execute(
            """
            INSERT INTO users (tg_id, active_workspace_id)
            VALUES (?, ?)
            ON CONFLICT(tg_id) DO UPDATE SET active_workspace_id = excluded.active_workspace_id
            """,
            (tg_id, ws_id),
        )
        conn.commit()
        conn.close()
        return {
            "id": ws_id,
            "title": fallback["title"],
            "owner_tg_id": int(fallback["owner_tg_id"]),
            "role": fallback["role"] or "member",
        }

    conn.close()
    return None


def db_set_active_workspace(tg_id: int, workspace_id: int, is_admin: bool) -> bool:
    conn = get_conn()
    cur = conn.cursor()
    ws = cur.execute("SELECT id FROM workspaces WHERE id = ?", (workspace_id,)).fetchone()
    if not ws:
        conn.close()
        return False

    if not is_admin:
        member = cur.execute(
            "SELECT 1 FROM workspace_members WHERE workspace_id = ? AND tg_id = ?",
            (workspace_id, tg_id),
        ).fetchone()
        if not member:
            conn.close()
            return False

    cur.execute(
        """
        INSERT INTO users (tg_id, active_workspace_id)
        VALUES (?, ?)
        ON CONFLICT(tg_id) DO UPDATE SET active_workspace_id = excluded.active_workspace_id
        """,
        (tg_id, workspace_id),
    )
    conn.commit()
    conn.close()
    return True


def db_create_and_bind_workspace(tg_id: int, title: str) -> int:
    conn = get_conn()
    cur = conn.cursor()
    now_str = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    cur.execute(
        "INSERT INTO workspaces (title, owner_tg_id, created_at) VALUES (?, ?, ?)",
        (title, tg_id, now_str),
    )
    ws_id = int(cur.lastrowid)
    cur.execute(
        """
        INSERT OR IGNORE INTO workspace_members (workspace_id, tg_id, role)
        VALUES (?, ?, 'owner')
        """,
        (ws_id, tg_id),
    )
    cur.execute(
        """
        INSERT INTO users (tg_id, active_workspace_id)
        VALUES (?, ?)
        ON CONFLICT(tg_id) DO UPDATE SET active_workspace_id = excluded.active_workspace_id
        """,
        (tg_id, ws_id),
    )
    conn.commit()
    conn.close()
    return ws_id


def db_get_workspace_info(workspace_id: int) -> Optional[dict]:
    conn = get_conn()
    cur = conn.cursor()
    row = cur.execute(
        "SELECT id, title, owner_tg_id FROM workspaces WHERE id = ?",
        (workspace_id,),
    ).fetchone()
    conn.close()
    if not row:
        return None
    return {
        "id": int(row["id"]),
        "title": row["title"],
        "owner_tg_id": int(row["owner_tg_id"]),
    }


def db_rename_workspace(workspace_id: int, new_title: str, caller_tg_id: int, is_admin: bool) -> tuple[str, Optional[str]]:
    conn = get_conn()
    cur = conn.cursor()
    ws = cur.execute(
        "SELECT id, title, owner_tg_id FROM workspaces WHERE id = ?",
        (workspace_id,),
    ).fetchone()
    if not ws:
        conn.close()
        return "not_found", None
    if not is_admin and int(ws["owner_tg_id"]) != caller_tg_id:
        conn.close()
        return "forbidden", None

    old_title = ws["title"]
    cur.execute("UPDATE workspaces SET title = ? WHERE id = ?", (new_title, workspace_id))
    conn.commit()
    conn.close()
    return "ok", old_title


def db_delete_workspace(workspace_id: int, caller_tg_id: int, is_admin: bool) -> tuple[str, Optional[str]]:
    conn = get_conn()
    cur = conn.cursor()
    ws = cur.execute(
        "SELECT id, title, owner_tg_id FROM workspaces WHERE id = ?",
        (workspace_id,),
    ).fetchone()
    if not ws:
        conn.close()
        return "not_found", None

    if not is_admin and int(ws["owner_tg_id"]) != caller_tg_id:
        conn.close()
        return "forbidden", None

    title = ws["title"]
    cur.execute("DELETE FROM movements WHERE workspace_id = ?", (workspace_id,))
    cur.execute("DELETE FROM catalog_items WHERE workspace_id = ?", (workspace_id,))
    cur.execute("DELETE FROM workspace_invites WHERE workspace_id = ?", (workspace_id,))
    cur.execute("DELETE FROM pending_invites WHERE workspace_id = ?", (workspace_id,))
    cur.execute("DELETE FROM workspace_members WHERE workspace_id = ?", (workspace_id,))
    cur.execute("DELETE FROM workspaces WHERE id = ?", (workspace_id,))

    affected_users = cur.execute(
        "SELECT tg_id FROM users WHERE active_workspace_id = ?",
        (workspace_id,),
    ).fetchall()
    for u in affected_users:
        uid = int(u["tg_id"])
        next_ws = cur.execute(
            "SELECT workspace_id FROM workspace_members WHERE tg_id = ? ORDER BY workspace_id ASC LIMIT 1",
            (uid,),
        ).fetchone()
        next_id = int(next_ws["workspace_id"]) if next_ws else None
        cur.execute("UPDATE users SET active_workspace_id = ? WHERE tg_id = ?", (next_id, uid))

    conn.commit()
    conn.close()
    return "ok", title


def db_get_workspace_members(workspace_id: int, include_admins: bool) -> list[dict]:
    conn = get_conn()
    cur = conn.cursor()
    rows = cur.execute(
        """
        SELECT wm.tg_id, wm.role, u.username, u.full_name
        FROM workspace_members wm
        LEFT JOIN users u ON u.tg_id = wm.tg_id
        WHERE wm.workspace_id = ?
        ORDER BY CASE wm.role WHEN 'owner' THEN 0 ELSE 1 END, u.full_name ASC
        """,
        (workspace_id,),
    ).fetchall()
    conn.close()

    result = []
    for r in rows:
        uid = int(r["tg_id"])
        if not include_admins and uid in ADMIN_IDS:
            continue
        result.append(
            {
                "tg_id": uid,
                "role": r["role"],
                "username": (r["username"] or "").strip(),
                "full_name": (r["full_name"] or "").strip(),
            }
        )
    return result


def db_get_pending_invites(workspace_id: int) -> list[dict]:
    conn = get_conn()
    cur = conn.cursor()
    rows = cur.execute(
        """
        SELECT id, username_norm
        FROM pending_invites
        WHERE workspace_id = ?
        ORDER BY id ASC
        """,
        (workspace_id,),
    ).fetchall()
    conn.close()
    return [{"id": int(r["id"]), "username_norm": r["username_norm"]} for r in rows]


def db_remove_member(workspace_id: int, target_tg_id: int) -> tuple[str, str]:
    conn = get_conn()
    cur = conn.cursor()
    ws = cur.execute("SELECT owner_tg_id, title FROM workspaces WHERE id = ?", (workspace_id,)).fetchone()
    if not ws:
        conn.close()
        return "not_found", ""
    if int(ws["owner_tg_id"]) == target_tg_id:
        conn.close()
        return "owner", ws["title"]

    u = cur.execute("SELECT username, full_name FROM users WHERE tg_id = ?", (target_tg_id,)).fetchone()
    label = (
        f"@{u['username']}"
        if u and u["username"]
        else ((u["full_name"] if u and u["full_name"] else "Сотрудник"))
    )

    cur.execute("DELETE FROM workspace_members WHERE workspace_id = ? AND tg_id = ?", (workspace_id, target_tg_id))
    cur.execute(
        "UPDATE users SET active_workspace_id = NULL WHERE tg_id = ? AND active_workspace_id = ?",
        (target_tg_id, workspace_id),
    )
    conn.commit()
    conn.close()
    return "ok", label


def db_remove_pending_invite(workspace_id: int, invite_id: int) -> Optional[str]:
    conn = get_conn()
    cur = conn.cursor()
    row = cur.execute(
        "SELECT username_norm FROM pending_invites WHERE id = ? AND workspace_id = ?",
        (invite_id, workspace_id),
    ).fetchone()
    if not row:
        conn.close()
        return None
    uname = row["username_norm"]
    cur.execute("DELETE FROM pending_invites WHERE id = ?", (invite_id,))
    conn.commit()
    conn.close()
    return uname


def db_list_workspaces() -> list[sqlite3.Row]:
    conn = get_conn()
    cur = conn.cursor()
    rows = cur.execute(
        """
        SELECT w.id, w.title, w.owner_tg_id,
               (SELECT COUNT(*) FROM workspace_members wm WHERE wm.workspace_id = w.id) AS users_count,
               (SELECT COUNT(*) FROM movements m WHERE m.workspace_id = w.id) AS ops_count
        FROM workspaces w
        ORDER BY w.id DESC
        """
    ).fetchall()
    conn.close()
    return rows


def db_get_catalog(workspace_id: int) -> list[dict]:
    conn = get_conn()
    cur = conn.cursor()
    rows = cur.execute(
        "SELECT id, name, unit, category FROM catalog_items WHERE workspace_id = ? ORDER BY id ASC",
        (workspace_id,),
    ).fetchall()
    conn.close()
    return [
        {"id": int(r["id"]), "name": r["name"], "unit": r["unit"], "category": r["category"]}
        for r in rows
    ]


def db_get_catalog_item(workspace_id: int, item_id: int) -> Optional[dict]:
    conn = get_conn()
    cur = conn.cursor()
    row = cur.execute(
        "SELECT id, name, unit, category FROM catalog_items WHERE id = ? AND workspace_id = ?",
        (item_id, workspace_id),
    ).fetchone()
    conn.close()
    if not row:
        return None
    return {"id": int(row["id"]), "name": row["name"], "unit": row["unit"], "category": row["category"]}


def db_catalog_name_exists(workspace_id: int, name: str, exclude_id: Optional[int] = None) -> bool:
    conn = get_conn()
    cur = conn.cursor()
    if exclude_id is not None:
        row = cur.execute(
            "SELECT 1 FROM catalog_items WHERE workspace_id = ? AND lower(trim(name)) = lower(trim(?)) AND id != ?",
            (workspace_id, name, exclude_id),
        ).fetchone()
    else:
        row = cur.execute(
            "SELECT 1 FROM catalog_items WHERE workspace_id = ? AND lower(trim(name)) = lower(trim(?))",
            (workspace_id, name),
        ).fetchone()
    conn.close()
    return row is not None


def db_add_catalog_item(workspace_id: int, name: str, unit: str, category: str):
    conn = get_conn()
    cur = conn.cursor()
    cur.execute(
        "INSERT INTO catalog_items (workspace_id, name, unit, category) VALUES (?, ?, ?, ?)",
        (workspace_id, name, unit, category),
    )
    conn.commit()
    conn.close()


def db_update_catalog_item(
    workspace_id: int, item_id: int, new_name: str, new_unit: str, new_category: str
) -> tuple[str, Optional[dict]]:
    conn = get_conn()
    cur = conn.cursor()
    cur_item = cur.execute(
        "SELECT id, name, unit, category FROM catalog_items WHERE id = ? AND workspace_id = ?",
        (item_id, workspace_id),
    ).fetchone()
    if not cur_item:
        conn.close()
        return "not_found", None

    dup = cur.execute(
        "SELECT 1 FROM catalog_items WHERE workspace_id = ? AND lower(trim(name)) = lower(trim(?)) AND id != ?",
        (workspace_id, new_name, item_id),
    ).fetchone()
    if dup:
        conn.close()
        return "duplicate", None

    old_data = {
        "name": cur_item["name"],
        "unit": cur_item["unit"],
        "category": cur_item["category"],
    }
    cur.execute(
        "UPDATE catalog_items SET name = ?, unit = ?, category = ? WHERE id = ? AND workspace_id = ?",
        (new_name, new_unit, new_category, item_id, workspace_id),
    )
    cur.execute(
        """
        UPDATE movements
        SET item_name = ?, unit = ?, category = ?
        WHERE workspace_id = ? AND lower(trim(item_name)) = lower(trim(?))
        """,
        (new_name, new_unit, new_category, workspace_id, old_data["name"]),
    )
    conn.commit()
    conn.close()
    return "ok", old_data


def db_get_stock(workspace_id: int) -> list[dict]:
    conn = get_conn()
    cur = conn.cursor()
    rows = cur.execute(
        """
        SELECT
            c.name AS name,
            c.unit AS unit,
            c.category AS category,
            COALESCE(SUM(
                CASE
                    WHEN m.op_type = 'Приход' THEN m.quantity
                    WHEN m.op_type = 'Расход' THEN -m.quantity
                    ELSE 0
                END
            ), 0) AS balance
        FROM catalog_items c
        LEFT JOIN movements m
            ON m.workspace_id = c.workspace_id
           AND lower(trim(m.item_name)) = lower(trim(c.name))
        WHERE c.workspace_id = ?
        GROUP BY c.id, c.name, c.unit, c.category
        ORDER BY c.category ASC, c.name ASC
        """,
        (workspace_id,),
    ).fetchall()
    conn.close()
    return [
        {"name": r["name"], "unit": r["unit"], "category": r["category"], "balance": float(r["balance"])}
        for r in rows
    ]


def db_add_movement(
    workspace_id: int,
    op_date: str,
    op_type: str,
    item_name: str,
    quantity: float,
    unit: str,
    category: str,
    comment: str,
    actor_name: str,
    actor_tg_id: int,
):
    conn = get_conn()
    cur = conn.cursor()
    now_str = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    cur.execute(
        """
        INSERT INTO movements (
            workspace_id, op_date, op_type, item_name, quantity,
            unit, category, comment, actor_name, actor_tg_id, created_at
        )
        VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
        """,
        (workspace_id, op_date, op_type, item_name, quantity, unit, category, comment, actor_name, actor_tg_id, now_str),
    )
    conn.commit()
    conn.close()


def db_import_stock_rows(
    workspace_id: int,
    rows: list[dict],
    actor_name: str,
    actor_tg_id: int,
) -> dict:
    conn = get_conn()
    cur = conn.cursor()
    now_str = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    today_str = datetime.now().strftime("%d.%m.%Y")

    added_catalog = 0
    adjusted_items = 0
    zero_diff_items = 0

    for r in rows:
        name = r["name"]
        target_qty = float(r["qty"])
        category = r["category"] or "Основное"
        unit = r["unit"] or "шт"

        cat_row = cur.execute(
            "SELECT id, name, unit, category FROM catalog_items WHERE workspace_id = ? AND lower(trim(name)) = lower(trim(?))",
            (workspace_id, name),
        ).fetchone()

        if not cat_row:
            cur.execute(
                "INSERT INTO catalog_items (workspace_id, name, unit, category) VALUES (?, ?, ?, ?)",
                (workspace_id, name, unit, category),
            )
            added_catalog += 1
            if target_qty > 0:
                cur.execute(
                    """
                    INSERT INTO movements (
                        workspace_id, op_date, op_type, item_name, quantity,
                        unit, category, comment, actor_name, actor_tg_id, created_at
                    )
                    VALUES (?, ?, 'Приход', ?, ?, ?, ?, 'Загрузка остатков из файла', ?, ?, ?)
                    """,
                    (workspace_id, today_str, name, target_qty, unit, category, actor_name, actor_tg_id, now_str),
                )
            elif target_qty < 0:
                cur.execute(
                    """
                    INSERT INTO movements (
                        workspace_id, op_date, op_type, item_name, quantity,
                        unit, category, comment, actor_name, actor_tg_id, created_at
                    )
                    VALUES (?, ?, 'Расход', ?, ?, ?, ?, 'Загрузка остатков из файла', ?, ?, ?)
                    """,
                    (workspace_id, today_str, name, abs(target_qty), unit, category, actor_name, actor_tg_id, now_str),
                )
        else:
            canon_name = cat_row["name"]
            canon_unit = cat_row["unit"] or unit
            canon_cat = cat_row["category"] or category
            bal_row = cur.execute(
                """
                SELECT COALESCE(SUM(
                    CASE
                        WHEN op_type = 'Приход' THEN quantity
                        WHEN op_type = 'Расход' THEN -quantity
                        ELSE 0
                    END
                ), 0) AS bal
                FROM movements
                WHERE workspace_id = ? AND lower(trim(item_name)) = lower(trim(?))
                """,
                (workspace_id, canon_name),
            ).fetchone()
            cur_bal = float(bal_row["bal"] or 0.0)
            diff = target_qty - cur_bal
            if abs(diff) > 1e-9:
                op_type = "Приход" if diff > 0 else "Расход"
                cur.execute(
                    """
                    INSERT INTO movements (
                        workspace_id, op_date, op_type, item_name, quantity,
                        unit, category, comment, actor_name, actor_tg_id, created_at
                    )
                    VALUES (?, ?, ?, ?, ?, ?, ?, 'Корректировка по файлу остатков', ?, ?, ?)
                    """,
                    (workspace_id, today_str, op_type, canon_name, abs(diff), canon_unit, canon_cat, actor_name, actor_tg_id, now_str),
                )
                adjusted_items += 1
            else:
                zero_diff_items += 1

    conn.commit()
    conn.close()
    return {
        "total": len(rows),
        "added_catalog": added_catalog,
        "adjusted_items": adjusted_items,
        "zero_diff_items": zero_diff_items,
    }


def db_get_recent_movements(workspace_id: int, limit: int = 10) -> list[sqlite3.Row]:
    conn = get_conn()
    cur = conn.cursor()
    rows = cur.execute(
        """
        SELECT id, op_date, op_type, item_name, quantity, unit, comment, actor_name, actor_tg_id
        FROM movements
        WHERE workspace_id = ?
        ORDER BY id DESC
        LIMIT ?
        """,
        (workspace_id, limit),
    ).fetchall()
    conn.close()
    return rows


def db_get_movements_by_date(workspace_id: int, op_date: str, limit: int = 30) -> list[sqlite3.Row]:
    conn = get_conn()
    cur = conn.cursor()
    rows = cur.execute(
        """
        SELECT id, op_date, op_type, item_name, quantity, unit, comment, actor_name, actor_tg_id
        FROM movements
        WHERE workspace_id = ? AND trim(op_date) = trim(?)
        ORDER BY id DESC
        LIMIT ?
        """,
        (workspace_id, op_date, limit),
    ).fetchall()
    conn.close()
    return rows


def db_get_movements_by_item(workspace_id: int, item_name: str, limit: int = 25) -> list[sqlite3.Row]:
    conn = get_conn()
    cur = conn.cursor()
    rows = cur.execute(
        """
        SELECT id, op_date, op_type, item_name, quantity, unit, comment, actor_name, actor_tg_id
        FROM movements
        WHERE workspace_id = ? AND lower(trim(item_name)) = lower(trim(?))
        ORDER BY id DESC
        LIMIT ?
        """,
        (workspace_id, item_name, limit),
    ).fetchall()
    conn.close()
    return rows


def db_get_all_movements_for_export(workspace_id: int) -> list[sqlite3.Row]:
    conn = get_conn()
    cur = conn.cursor()
    rows = cur.execute(
        """
        SELECT op_date, op_type, category, item_name, quantity, unit, comment, actor_name
        FROM movements
        WHERE workspace_id = ?
        ORDER BY id DESC
        """,
        (workspace_id,),
    ).fetchall()
    conn.close()
    return rows


def db_delete_movement_by_id(
    workspace_id: int, movement_id: int, caller_tg_id: int, is_admin: bool
) -> tuple[str, Optional[str]]:
    conn = get_conn()
    cur = conn.cursor()
    row = cur.execute(
        """
        SELECT id, op_date, op_type, item_name, quantity, unit, category, comment, actor_name, actor_tg_id
        FROM movements
        WHERE id = ? AND workspace_id = ?
        """,
        (movement_id, workspace_id),
    ).fetchone()
    if not row:
        conn.close()
        return "not_found", None

    if not is_admin and int(row["actor_tg_id"]) != caller_tg_id:
        conn.close()
        return "forbidden", None

    info = f"{row['op_date']} · {row['op_type']} · {row['item_name']} · {float(row['quantity']):g} {row['unit']}".strip()
    now_str = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    cur.execute(
        """
        INSERT INTO cancelled_movements (
            orig_movement_id, workspace_id, op_date, op_type, item_name, quantity,
            unit, category, comment, actor_name, actor_tg_id, cancelled_by_tg_id, cancelled_at
        )
        VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
        """,
        (
            int(row["id"]),
            workspace_id,
            row["op_date"],
            row["op_type"],
            row["item_name"],
            float(row["quantity"]),
            row["unit"],
            row["category"],
            row["comment"],
            row["actor_name"],
            int(row["actor_tg_id"]),
            caller_tg_id,
            now_str,
        ),
    )
    cur.execute("DELETE FROM movements WHERE id = ?", (movement_id,))
    conn.commit()
    conn.close()
    return "ok", info


async def notify_admins(text: str):
    for admin_id in ADMIN_IDS:
        try:
            await bot.send_message(admin_id, text, parse_mode="Markdown")
        except Exception as e:
            logging.warning(f"Не удалось отправить уведомление админу {admin_id}: {e}")


def format_actor_public(actor_name: str, chat: types.Chat) -> str:
    clean_name = actor_name or "Пользователь"
    chat_info = f" в группе «{chat.title}»" if chat.type in ("group", "supergroup") else ""
    return f"{clean_name}{chat_info}"


def format_movements_lines(rows: list[sqlite3.Row], chat: types.Chat, is_admin: bool) -> list[str]:
    lines: list[str] = []
    for r in rows:
        sign = "➕" if r["op_type"] == "Приход" else "➖"
        extra = []
        if r["actor_name"]:
            actor_text = (
                f"{r['actor_name']} [ID: `{r['actor_tg_id']}`]"
                if is_admin and r["actor_tg_id"]
                else format_actor_public(r["actor_name"], chat)
            )
            extra.append(actor_text)
        if r["comment"]:
            extra.append(f"«{r['comment']}»")
        extra_str = f" — {', '.join(extra)}" if extra else ""
        unit_str = f" {r['unit']}" if r["unit"] else ""
        lines.append(f"{sign} `{r['op_date']}` **{r['item_name']}**: {float(r['quantity']):g}{unit_str}{extra_str}")
    return lines


def parse_number(raw: str) -> Optional[float]:
    try:
        return float(raw.replace(",", ".").replace(" ", "").strip())
    except ValueError:
        return None


def parse_uploaded_stock_file(file_bytes: bytes, filename: str) -> tuple[list[dict], list[str]]:
    errors: list[str] = []
    parsed_rows: list[dict] = []
    lower_name = (filename or "").lower()

    raw_table: list[list[str]] = []
    if lower_name.endswith(".xlsx"):
        if not HAS_OPENPYXL:
            return [], ["На сервере не установлен openpyxl для чтения .xlsx"]
        wb = openpyxl.load_workbook(io.BytesIO(file_bytes), data_only=True)
        sheet = wb.active
        for row in sheet.iter_rows(values_only=True):
            if not row:
                continue
            vals = ["" if v is None else str(v).strip() for v in row]
            if any(vals):
                raw_table.append(vals)
    else:
        text = None
        for enc in ("utf-8-sig", "utf-8", "cp1251"):
            try:
                text = file_bytes.decode(enc)
                break
            except UnicodeDecodeError:
                continue
        if text is None:
            return [], ["Не удалось прочитать кодировку файла"]
        delim = ";" if text.count(";") >= text.count(",") else ","
        reader = csv.reader(io.StringIO(text), delimiter=delim)
        for row in reader:
            if row and any(c.strip() for c in row):
                raw_table.append([c.strip() for c in row])

    if not raw_table:
        return [], ["Файл пуст"]

    col_map = {"category": 0, "name": 1, "qty": 2, "unit": 3}
    start_idx = 0
    first_lower = [c.lower() for c in raw_table[0]]
    header_hits = 0
    for idx, h in enumerate(first_lower):
        if any(k in h for k in ("категор", "групп")):
            col_map["category"] = idx
            header_hits += 1
        elif any(k in h for k in ("наимен", "назван", "позиц", "товар", "номенкл")):
            col_map["name"] = idx
            header_hits += 1
        elif any(k in h for k in ("остат", "количест", "кол-во", "колво", "qty", "баланс")):
            col_map["qty"] = idx
            header_hits += 1
        elif any(k in h for k in ("ед.", "един", "unit", "измер")):
            col_map["unit"] = idx
            header_hits += 1

    if header_hits >= 2:
        start_idx = 1

    for line_num, row in enumerate(raw_table[start_idx:], start=start_idx + 1):
        first_val_low = row[0].lower() if row else ""
        if first_val_low in ("склад", "дата выгрузки", "итого", "всего"):
            continue

        def get_col(key: str, default: str = "") -> str:
            i = col_map[key]
            return row[i].strip() if i < len(row) else default

        name = " ".join(get_col("name").split())
        if not name:
            continue

        qty_raw = get_col("qty")
        qty = parse_number(qty_raw)
        if qty is None:
            errors.append(f"Строка {line_num} («{name}»): некорректное количество «{qty_raw}»")
            continue

        category = " ".join(get_col("category", "Основное").split()) or "Основное"
        unit = get_col("unit", "шт").strip() or "шт"
        parsed_rows.append({"category": category, "name": name, "qty": qty, "unit": unit})

    return parsed_rows, errors


def build_stock_xlsx_bytes(ws_title: str, stock_list: list[dict], movements: list[sqlite3.Row]) -> bytes:
    wb = openpyxl.Workbook()
    sheet = wb.active
    sheet.title = "Остатки"

    sheet.append(["Склад", ws_title])
    sheet.append(["Дата выгрузки", datetime.now().strftime("%d.%m.%Y %H:%M")])
    sheet.append([])
    sheet.append(["Категория", "Наименование", "Остаток", "Ед. измерения"])

    zero_fill = PatternFill(start_color="FFF2CC", end_color="FFF2CC", fill_type="solid")
    neg_fill = PatternFill(start_color="FCE4D6", end_color="FCE4D6", fill_type="solid")

    for row_idx, item in enumerate(stock_list, start=5):
        sheet.append([item["category"], item["name"], item["balance"], item["unit"]])
        if item["balance"] < 0:
            for col_i in range(1, 5):
                sheet.cell(row=row_idx, column=col_i).fill = neg_fill
        elif item["balance"] == 0:
            for col_i in range(1, 5):
                sheet.cell(row=row_idx, column=col_i).fill = zero_fill

    sheet.column_dimensions["A"].width = 22
    sheet.column_dimensions["B"].width = 34
    sheet.column_dimensions["C"].width = 14
    sheet.column_dimensions["D"].width = 16

    ops_sheet = wb.create_sheet(title="Операции")
    ops_sheet.append(["Дата", "Тип", "Категория", "Наименование", "Количество", "Ед. измерения", "Комментарий", "Кто"])
    for m in movements:
        ops_sheet.append([
            m["op_date"],
            m["op_type"],
            m["category"],
            m["item_name"],
            float(m["quantity"]),
            m["unit"],
            m["comment"],
            m["actor_name"],
        ])
    ops_sheet.column_dimensions["A"].width = 14
    ops_sheet.column_dimensions["B"].width = 12
    ops_sheet.column_dimensions["C"].width = 20
    ops_sheet.column_dimensions["D"].width = 32
    ops_sheet.column_dimensions["E"].width = 14
    ops_sheet.column_dimensions["F"].width = 14
    ops_sheet.column_dimensions["G"].width = 28
    ops_sheet.column_dimensions["H"].width = 22

    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()


def build_stock_csv_bytes(ws_title: str, stock_list: list[dict]) -> bytes:
    sio = io.StringIO()
    writer = csv.writer(sio, delimiter=";")
    writer.writerow(["Склад", ws_title])
    writer.writerow(["Дата выгрузки", datetime.now().strftime("%d.%m.%Y %H:%M")])
    writer.writerow([])
    writer.writerow(["Категория", "Наименование", "Остаток", "Ед. измерения"])
    for item in stock_list:
        writer.writerow([item["category"], item["name"], f"{item['balance']:g}", item["unit"]])
    return sio.getvalue().encode("utf-8-sig")


class HistoryFilterStates(StatesGroup):
    entering_date = State()
    searching_item = State()


class OperationStates(StatesGroup):
    choosing_date = State()
    entering_custom_date = State()
    choosing_category = State()
    searching_item = State()
    choosing_item = State()
    entering_quantity = State()
    entering_comment = State()
    confirming = State()


class RepeatOperationStates(StatesGroup):
    choosing_date = State()
    entering_custom_date = State()
    confirming = State()


class ImportStockStates(StatesGroup):
    waiting_file = State()


class AddCatalogStates(StatesGroup):
    entering_name = State()
    entering_unit = State()
    entering_category = State()


class EditCatalogStates(StatesGroup):
    entering_name = State()
    entering_unit = State()
    entering_category = State()


class CreateWorkspaceStates(StatesGroup):
    entering_title = State()


class EditWorkspaceStates(StatesGroup):
    entering_title = State()


class FeedbackStates(StatesGroup):
    entering_message = State()


def get_main_keyboard(has_workspace: bool) -> ReplyKeyboardMarkup:
    if not has_workspace:
        rows = [
            [KeyboardButton(text="🆕 Создать склад")],
            [KeyboardButton(text="✍️ Написать админу"), KeyboardButton(text="📋 Меню")],
        ]
    else:
        rows = [
            [KeyboardButton(text="➕ Приход"), KeyboardButton(text="➖ Расход")],
            [KeyboardButton(text="📊 Текущий остаток"), KeyboardButton(text="📝 Добавить позицию")],
            [KeyboardButton(text="✍️ Написать админу"), KeyboardButton(text="📋 Меню")],
        ]
    return ReplyKeyboardMarkup(keyboard=rows, resize_keyboard=True)


def get_menu_list_keyboard(has_workspace: bool, is_admin: bool) -> ReplyKeyboardMarkup:
    rows: list[list[KeyboardButton]] = []
    if has_workspace:
        rows.append([KeyboardButton(text="📜 Последние операции"), KeyboardButton(text="🔁 Повторить последнюю")])
        rows.append([KeyboardButton(text="🗑 Отменить операцию")])
        rows.append([KeyboardButton(text="📤 Загрузить остатки из Excel")])
        rows.append([KeyboardButton(text="📥 Остатки в Excel")])
        rows.append([KeyboardButton(text="✏️ Редактировать позицию")])
        rows.append([KeyboardButton(text="✏️ Редактировать склад"), KeyboardButton(text="🔄 Сменить склад")])
        rows.append([KeyboardButton(text="👤 Участники склада"), KeyboardButton(text="👥 Пригласить сотрудника")])
        rows.append([KeyboardButton(text="❌ Отвязать сотрудника"), KeyboardButton(text="ℹ️ Мой склад")])
        rows.append([KeyboardButton(text="🆕 Создать склад"), KeyboardButton(text="🗑 Удалить склад")])
    else:
        rows.append([KeyboardButton(text="🆕 Создать склад")])

    rows.append([KeyboardButton(text="✍️ Написать админу")])

    if is_admin:
        rows.append([KeyboardButton(text="👑 Все склады (Админ)"), KeyboardButton(text="💾 Бэкап базы (Админ)")])

    rows.append([KeyboardButton(text="⬅️ Назад")])
    return ReplyKeyboardMarkup(keyboard=rows, resize_keyboard=True)


def get_cancel_keyboard() -> ReplyKeyboardMarkup:
    return ReplyKeyboardMarkup(
        keyboard=[[KeyboardButton(text="❌ Отмена")]],
        resize_keyboard=True,
    )


async def send_with_fresh_keyboard(
    message: types.Message,
    text: str,
    has_workspace: bool,
    is_admin: bool = False,
    in_menu: bool = False,
    parse_mode: Optional[str] = "Markdown",
):
    kb = (
        get_menu_list_keyboard(has_workspace=has_workspace, is_admin=is_admin)
        if in_menu
        else get_main_keyboard(has_workspace=has_workspace)
    )
    await message.answer(text, reply_markup=kb, parse_mode=parse_mode)


async def ensure_workspace(message: types.Message) -> Optional[dict]:
    user = message.from_user
    await db_upsert_user(user.id, user.username, user.full_name or "")
    ws = db_get_active_workspace(user.id)
    if not ws:
        await send_with_fresh_keyboard(
            message,
            "У вас пока нет подключённого склада.\nНажмите **«🆕 Создать склад»**, чтобы создать свой склад.",
            has_workspace=False,
        )
        return None
    return ws


async def build_invite_message(
    workspace_id: int, ws_title: str, creator_tg_id: int, rotate: bool = False
) -> tuple[str, InlineKeyboardMarkup]:
    bot_info = await bot.get_me()
    code = db_get_or_create_invite(workspace_id, creator_tg_id, rotate=rotate)
    link = f"https://t.me/{bot_info.username}?start={code}"
    text = (
        f"🔗 **Ссылка для приглашения на склад «{ws_title}»**\n\n"
        f"`{link}`\n\n"
        f"Отправьте эту ссылку сотруднику в WhatsApp или Telegram. "
        f"Когда он откроет ссылку и нажмёт **Запустить**, бот автоматически подключит его к складу."
    )
    markup = InlineKeyboardMarkup(
        inline_keyboard=[
            [InlineKeyboardButton(text="🔄 Перевыпустить ссылку", callback_data=f"rotinv:{workspace_id}")],
            [InlineKeyboardButton(text="❌ Закрыть", callback_data="cancel")],
        ]
    )
    return text, markup


@dp.message(CommandStart())
async def cmd_start(message: types.Message, state: FSMContext, command: CommandObject):
    await state.clear()
    user = message.from_user
    await db_upsert_user(user.id, user.username, user.full_name or "")

    args = (command.args or "").strip() if command else ""
    if args.startswith("inv_"):
        status, ws_info = db_accept_invite_code(args, user.id)
        if status == "joined" and ws_info:
            actor_label = user.full_name or (f"@{user.username}" if user.username else f"ID {user.id}")
            uname_part = f" (@{user.username})" if user.username else ""
            await send_with_fresh_keyboard(
                message,
                f"✅ Вы подключены к складу **«{ws_info['title']}»**!",
                has_workspace=True,
            )
            if ws_info["owner_tg_id"] != user.id:
                try:
                    await bot.send_message(
                        ws_info["owner_tg_id"],
                        f"👤 По ссылке к складу **«{ws_info['title']}»** подключился сотрудник: **{actor_label}**{uname_part}.",
                        parse_mode="Markdown",
                    )
                except Exception as e:
                    logging.warning(f"Не удалось уведомить владельца {ws_info['owner_tg_id']}: {e}")
            return

        if status == "already_member" and ws_info:
            await send_with_fresh_keyboard(
                message,
                f"Вы уже состоите на складе **«{ws_info['title']}»** — он выбран активным.",
                has_workspace=True,
            )
            return

        active_ws = db_get_active_workspace(user.id)
        await send_with_fresh_keyboard(
            message,
            "⚠️ Ссылка приглашения недействительна или устарела. Попросите владельца склада отправить новую ссылку.",
            has_workspace=bool(active_ws),
            parse_mode=None,
        )
        return

    ws = db_get_active_workspace(user.id)
    if ws:
        text = (
            f"Складской учёт готов к работе.\n"
            f"Активный склад: **«{ws['title']}»**\n"
            f"Выберите действие:"
        )
    else:
        text = (
            "Добро пожаловать в систему складского учёта!\n"
            "Нажмите **«🆕 Создать склад»**, чтобы создать свой склад."
        )

    await send_with_fresh_keyboard(
        message,
        text,
        has_workspace=bool(ws),
    )


@dp.message(F.text == "📋 Меню")
async def open_menu_buttons(message: types.Message, state: FSMContext):
    await state.clear()
    user = message.from_user
    await db_upsert_user(user.id, user.username, user.full_name or "")
    ws = db_get_active_workspace(user.id)
    is_admin = user.id in ADMIN_IDS

    await send_with_fresh_keyboard(
        message,
        "Выберите действие:",
        has_workspace=bool(ws),
        is_admin=is_admin,
        in_menu=True,
        parse_mode=None,
    )


@dp.message(F.text.in_(["⬅️ Назад", "❌ Отмена"]))
async def back_to_main_keyboard(message: types.Message, state: FSMContext):
    await state.clear()
    user = message.from_user
    await db_upsert_user(user.id, user.username, user.full_name or "")
    ws = db_get_active_workspace(user.id)
    await send_with_fresh_keyboard(
        message,
        "Главное меню:",
        has_workspace=bool(ws),
        parse_mode=None,
    )


@dp.callback_query(F.data == "cancel")
async def cancel_operation_cb(callback: types.CallbackQuery, state: FSMContext):
    await state.clear()
    user = callback.from_user
    ws = db_get_active_workspace(user.id)
    await callback.message.edit_text("Действие отменено.")
    await send_with_fresh_keyboard(
        callback.message,
        "Главное меню:",
        has_workspace=bool(ws),
        parse_mode=None,
    )
    await callback.answer()


# --- НАПИСАТЬ АДМИНУ И ОТВЕТЫ ЧЕРЕЗ REPLY ---
@dp.message(F.text == "✍️ Написать админу")
async def start_feedback(message: types.Message, state: FSMContext):
    user = message.from_user
    await db_upsert_user(user.id, user.username, user.full_name or "")
    await state.clear()
    await state.set_state(FeedbackStates.entering_message)
    await message.answer(
        "✍️ **Написать администратору**\n\n"
        "Напишите ваше сообщение, вопрос или пожелание одним сообщением. Я перешлю его администратору.",
        reply_markup=get_cancel_keyboard(),
        parse_mode="Markdown",
    )


@dp.message(FeedbackStates.entering_message)
async def process_feedback_message(message: types.Message, state: FSMContext):
    text = (message.text or "").strip()
    if not text:
        await message.answer(
            "Сообщение не может быть пустым. Напишите текст или нажмите «❌ Отмена»:",
            reply_markup=get_cancel_keyboard(),
        )
        return

    user = message.from_user
    ws = db_get_active_workspace(user.id)
    await state.clear()

    uname_str = f"@{user.username}" if user.username else "нет"
    name_str = user.full_name or "Пользователь"
    admin_text = (
        f"📩 **Новое сообщение от пользователя!**\n\n"
        f"• От: {name_str} (ID: `{user.id}`, {uname_str})\n"
        f"• Текст:\n{text}\n\n"
        f"_Ответьте Reply на это сообщение, чтобы написать пользователю._"
    )
    await notify_admins(admin_text)

    await send_with_fresh_keyboard(
        message,
        "✅ **Спасибо!** Ваше сообщение отправлено администратору.",
        has_workspace=bool(ws),
    )


@dp.message(F.reply_to_message, F.text)
async def handle_reply_messages(message: types.Message):
    replied = message.reply_to_message
    if not replied or not replied.text:
        return

    user = message.from_user
    text = (message.text or "").strip()
    if not text:
        return

    if user.id in ADMIN_IDS:
        match = re.search(r"ID:\s*`?(\d+)`?", replied.text)
        if not match:
            return
        target_user_id = int(match.group(1))
        try:
            await bot.send_message(
                chat_id=target_user_id,
                text=f"💬 **Ответ от администратора:**\n\n{text}\n\n_Ответьте Reply на это сообщение, чтобы продолжить диалог._",
                parse_mode="Markdown",
            )
            await message.reply("✅ Ответ успешно отправлен пользователю.")
        except Exception as e:
            await message.reply(f"❌ Не удалось отправить ответ: {e}")
        return

    if replied.from_user and replied.from_user.id == bot.id:
        uname_str = f"@{user.username}" if user.username else "нет"
        name_str = user.full_name or "Пользователь"
        admin_notification = (
            f"📩 **Ответ от пользователя в диалоге!**\n\n"
            f"• От: {name_str} (ID: `{user.id}`, {uname_str})\n"
            f"• Текст:\n{text}\n\n"
            f"_Ответьте Reply на это сообщение, чтобы ответить._"
        )
        await notify_admins(admin_notification)
        await message.reply("✅ Ваше сообщение отправлено администратору.")


# --- БЭКАП БАЗЫ ДЛЯ АДМИНА ---
@dp.message(F.text == "💾 Бэкап базы (Админ)")
async def admin_backup_db(message: types.Message):
    if message.from_user.id not in ADMIN_IDS:
        return
    if not os.path.exists(DB_FILE):
        await message.answer("Файл базы данных не найден.")
        return
    try:
        conn = get_conn()
        conn.execute("PRAGMA wal_checkpoint(FULL);")
        conn.close()
        with open(DB_FILE, "rb") as f:
            data = f.read()
        stamp = datetime.now().strftime("%Y%m%d_%H%M")
        doc = BufferedInputFile(data, filename=f"backup_sklad_{stamp}.db")
        await message.answer_document(doc, caption=f"💾 Резервная копия базы данных (сборка `{BOT_VERSION}`)", parse_mode="Markdown")
    except Exception as e:
        await message.answer(f"❌ Ошибка при создании бэкапа: {e}")


# --- СОЗДАНИЕ, РЕДАКТИРОВАНИЕ И УДАЛЕНИЕ СКЛАДА ---
@dp.message(F.text == "🆕 Создать склад")
async def msg_create_workspace(message: types.Message, state: FSMContext):
    user = message.from_user
    await db_upsert_user(user.id, user.username, user.full_name or "")
    await state.clear()
    await state.set_state(CreateWorkspaceStates.entering_title)
    await message.answer(
        "Введите **название нового склада** (например: `Склад КХ` или `Магазин Астана`):",
        reply_markup=get_cancel_keyboard(),
        parse_mode="Markdown",
    )


@dp.message(CreateWorkspaceStates.entering_title)
async def process_create_workspace_title(message: types.Message, state: FSMContext):
    title = " ".join((message.text or "").strip().split())
    if not title:
        await message.answer("Название не может быть пустым. Введите название склада:", reply_markup=get_cancel_keyboard())
        return

    user = message.from_user
    await db_upsert_user(user.id, user.username, user.full_name or "")
    db_create_and_bind_workspace(user.id, title)
    await state.clear()

    await send_with_fresh_keyboard(
        message,
        f"✅ Склад **«{title}»** создан и выбран активным!",
        has_workspace=True,
    )


@dp.message(F.text == "✏️ Редактировать склад")
async def msg_edit_workspace(message: types.Message, state: FSMContext):
    ws = await ensure_workspace(message)
    if not ws:
        return
    user = message.from_user
    is_admin = user.id in ADMIN_IDS
    if not is_admin and user.id != ws["owner_tg_id"]:
        await send_with_fresh_keyboard(
            message,
            "Редактировать название склада может только его владелец или администратор.",
            has_workspace=True,
            parse_mode=None,
        )
        return

    await state.clear()
    await state.update_data(target_ws_id=ws["id"], old_title=ws["title"])
    await state.set_state(EditWorkspaceStates.entering_title)
    await message.answer(
        f"Текущее название склада: **«{ws['title']}»**\nВведите **новое название склада**:",
        reply_markup=get_cancel_keyboard(),
        parse_mode="Markdown",
    )


@dp.message(EditWorkspaceStates.entering_title)
async def process_edit_workspace_title(message: types.Message, state: FSMContext):
    new_title = " ".join((message.text or "").strip().split())
    if not new_title:
        await message.answer("Название не может быть пустым. Введите новое название склада:", reply_markup=get_cancel_keyboard())
        return

    data = await state.get_data()
    ws_id = int(data["target_ws_id"])
    user = message.from_user
    is_admin = user.id in ADMIN_IDS

    status, old_title = db_rename_workspace(ws_id, new_title, user.id, is_admin)
    await state.clear()

    if status == "not_found":
        await send_with_fresh_keyboard(message, "Склад не найден.", has_workspace=True, parse_mode=None)
        return
    if status == "forbidden":
        await send_with_fresh_keyboard(message, "Недостаточно прав для редактирования склада.", has_workspace=True, parse_mode=None)
        return

    await send_with_fresh_keyboard(
        message,
        f"✅ Название склада изменено: **«{old_title}»** → **«{new_title}»**.",
        has_workspace=True,
    )


@dp.message(F.text == "🗑 Удалить склад")
async def request_delete_workspace(message: types.Message):
    ws = await ensure_workspace(message)
    if not ws:
        return
    user = message.from_user
    is_admin = user.id in ADMIN_IDS
    if not is_admin and user.id != ws["owner_tg_id"]:
        await send_with_fresh_keyboard(
            message,
            "Удалить склад может только его владелец или администратор.",
            has_workspace=True,
            parse_mode=None,
        )
        return

    markup = InlineKeyboardMarkup(
        inline_keyboard=[
            [InlineKeyboardButton(text="🗑 Да, удалить этот склад", callback_data=f"delws_ask:{ws['id']}")],
            [InlineKeyboardButton(text="❌ Отмена", callback_data="cancel")],
        ]
    )
    await message.answer(
        f"Вы уверены, что хотите удалить склад **«{ws['title']}»** вместе со всеми позициями и операциями?",
        reply_markup=markup,
        parse_mode="Markdown",
    )


@dp.callback_query(F.data.startswith("delws_ask:"))
async def confirm_delete_workspace_step(callback: types.CallbackQuery):
    ws_id = int(callback.data.split(":", 1)[1])
    user = callback.from_user
    is_admin = user.id in ADMIN_IDS
    ws = db_get_workspace_info(ws_id)
    if not ws:
        await callback.message.edit_text("Склад уже удалён.")
        await callback.answer()
        return
    if not is_admin and user.id != ws["owner_tg_id"]:
        await callback.answer("Недостаточно прав", show_alert=True)
        return

    markup = InlineKeyboardMarkup(
        inline_keyboard=[
            [InlineKeyboardButton(text="⚠️ Подтверждаю, удалить навсегда", callback_data=f"delws_go:{ws_id}")],
            [InlineKeyboardButton(text="❌ Отмена", callback_data="cancel")],
        ]
    )
    await callback.message.edit_text(
        f"⚠️ **Последнее подтверждение!**\nСклад **«{ws['title']}»** и вся его история будут удалены без возможности восстановления.",
        reply_markup=markup,
        parse_mode="Markdown",
    )
    await callback.answer()


@dp.callback_query(F.data.startswith("delws_go:"))
async def execute_delete_workspace(callback: types.CallbackQuery):
    ws_id = int(callback.data.split(":", 1)[1])
    user = callback.from_user
    is_admin = user.id in ADMIN_IDS

    status, deleted_title = db_delete_workspace(ws_id, user.id, is_admin)
    if status == "not_found":
        await callback.message.edit_text("Склад не найден.")
        await callback.answer()
        return
    if status == "forbidden":
        await callback.answer("Недостаточно прав", show_alert=True)
        return

    await callback.message.edit_text(f"🗑 Склад **«{deleted_title}»** удалён.", parse_mode="Markdown")
    await callback.answer()

    new_active = db_get_active_workspace(user.id)
    await bot.send_message(
        user.id,
        f"Активный склад: **«{new_active['title']}»**" if new_active else "У вас больше нет активных складов.",
        reply_markup=get_main_keyboard(has_workspace=bool(new_active)),
        parse_mode="Markdown",
    )


# --- ПЕРЕКЛЮЧЕНИЕ СКЛАДА, ИНФОРМАЦИЯ И СОТРУДНИКИ ---
@dp.message(F.text == "🔄 Сменить склад")
async def switch_user_workspace_menu(message: types.Message):
    user = message.from_user
    await db_upsert_user(user.id, user.username, user.full_name or "")
    workspaces = db_get_user_workspaces(user.id)
    active = db_get_active_workspace(user.id)
    active_id = active["id"] if active else None

    if not workspaces:
        await send_with_fresh_keyboard(
            message,
            "У вас пока нет доступных складов. Нажмите **«🆕 Создать склад»**.",
            has_workspace=False,
        )
        return

    buttons = []
    for ws in workspaces:
        marker = "✅ " if ws["id"] == active_id else ""
        role_label = "Создал сам" if ws["role"] == "owner" else "Подключили к складу"
        buttons.append(
            [
                InlineKeyboardButton(
                    text=f"{marker}{ws['title']} ({role_label})",
                    callback_data=f"myws:{ws['id']}",
                )
            ]
        )
    buttons.append([InlineKeyboardButton(text="❌ Отмена", callback_data="cancel")])

    await message.answer(
        "Выберите склад для работы:",
        reply_markup=InlineKeyboardMarkup(inline_keyboard=buttons),
    )


@dp.callback_query(F.data.startswith("myws:"))
async def process_switch_my_workspace(callback: types.CallbackQuery):
    ws_id = int(callback.data.split(":", 1)[1])
    user = callback.from_user
    await db_upsert_user(user.id, user.username, user.full_name or "")

    if db_set_active_workspace(user.id, ws_id, is_admin=user.id in ADMIN_IDS):
        ws = db_get_active_workspace(user.id)
        title = ws["title"] if ws else "Склад"
        await callback.message.edit_text(
            f"✅ Активный склад изменён на **«{title}»**.",
            parse_mode="Markdown",
        )
        await send_with_fresh_keyboard(
            callback.message,
            "Главное меню:",
            has_workspace=True,
            parse_mode=None,
        )
    else:
        await callback.message.edit_text("Не удалось переключиться на этот склад.")
    await callback.answer()


@dp.message(F.text == "ℹ️ Мой склад")
async def my_workspace_info(message: types.Message):
    ws = await ensure_workspace(message)
    if not ws:
        return

    user = message.from_user
    all_ws = db_get_user_workspaces(user.id)
    own_count = sum(1 for x in all_ws if x["role"] == "owner")
    linked_count = sum(1 for x in all_ws if x["role"] != "owner")
    lines = [
        f"📦 **Активный склад:** «{ws['title']}»",
        f"• Доступно складов: {len(all_ws)} (создано вами: {own_count}, подключили к складу: {linked_count})",
        "• Подключить сотрудника: нажмите **«📋 Меню»** → **«👥 Пригласить сотрудника»** и отправьте ему готовую ссылку.",
    ]
    if user.id in ADMIN_IDS:
        lines.append(f"• *(Админ)* ID владельца: `{ws['owner_tg_id']}`")
        lines.append(f"• *(Админ)* Версия сборки: `{BOT_VERSION}`")

    await send_with_fresh_keyboard(
        message,
        "\n".join(lines),
        has_workspace=True,
    )


@dp.message(F.text == "👤 Участники склада")
async def show_workspace_members(message: types.Message):
    ws = await ensure_workspace(message)
    if not ws:
        return

    is_admin = message.from_user.id in ADMIN_IDS
    members = db_get_workspace_members(ws["id"], include_admins=is_admin)
    pending = db_get_pending_invites(ws["id"])

    lines = [f"👤 **Участники склада «{ws['title']}»:**\n"]
    if not members:
        lines.append("• Подключённых пользователей пока нет.")
    else:
        for m in members:
            role_txt = "владелец" if m["role"] == "owner" else "сотрудник"
            uname_txt = f" (@{m['username']})" if m["username"] else ""
            name_txt = m["full_name"] or (f"@{m['username']}" if m["username"] else "Участник")
            id_txt = f" [ID: `{m['tg_id']}`]" if is_admin else ""
            lines.append(f"• **{name_txt}**{uname_txt} — {role_txt}{id_txt}")

    if pending:
        lines.append("\n⏳ **Ожидают первого входа в бот (по @username):**")
        for p in pending:
            lines.append(f"• @{p['username_norm']}")

    await send_with_fresh_keyboard(
        message,
        "\n".join(lines),
        has_workspace=True,
    )


@dp.message(F.text == "👥 Пригласить сотрудника")
async def send_workspace_invite_link(message: types.Message, state: FSMContext):
    ws = await ensure_workspace(message)
    if not ws:
        return
    user = message.from_user
    if user.id not in ADMIN_IDS and user.id != ws["owner_tg_id"]:
        await send_with_fresh_keyboard(
            message,
            "Создавать ссылки для приглашения может только владелец склада или администратор.",
            has_workspace=True,
            parse_mode=None,
        )
        return

    await state.clear()
    text, markup = await build_invite_message(ws["id"], ws["title"], user.id, rotate=False)
    await message.answer(text, reply_markup=markup, parse_mode="Markdown")


@dp.callback_query(F.data.startswith("rotinv:"))
async def rotate_workspace_invite_link(callback: types.CallbackQuery):
    ws_id = int(callback.data.split(":", 1)[1])
    user = callback.from_user
    ws = db_get_workspace_info(ws_id)
    if not ws:
        await callback.message.edit_text("Склад не найден.")
        await callback.answer()
        return
    if user.id not in ADMIN_IDS and user.id != ws["owner_tg_id"]:
        await callback.answer("Недостаточно прав", show_alert=True)
        return

    text, markup = await build_invite_message(ws["id"], ws["title"], user.id, rotate=True)
    await callback.message.edit_text(text, reply_markup=markup, parse_mode="Markdown")
    await callback.answer("Ссылка обновлена")


@dp.message(F.text == "❌ Отвязать сотрудника")
async def start_unlink_user_ui(message: types.Message):
    ws = await ensure_workspace(message)
    if not ws:
        return
    user = message.from_user
    is_admin = user.id in ADMIN_IDS
    if not is_admin and user.id != ws["owner_tg_id"]:
        await send_with_fresh_keyboard(
            message,
            "Отвязывать сотрудников может только владелец склада или администратор.",
            has_workspace=True,
            parse_mode=None,
        )
        return

    members = db_get_workspace_members(ws["id"], include_admins=is_admin)
    pending = db_get_pending_invites(ws["id"])

    buttons = []
    for m in members:
        if m["role"] == "owner":
            continue
        label_name = f"@{m['username']}" if m["username"] else (m["full_name"] or "Сотрудник")
        buttons.append(
            [InlineKeyboardButton(text=f"❌ Отвязать {label_name}", callback_data=f"unlink_u:{ws['id']}:{m['tg_id']}")]
        )

    for p in pending:
        buttons.append(
            [InlineKeyboardButton(text=f"⏳ Отменить приглашение @{p['username_norm']}", callback_data=f"unlink_p:{ws['id']}:{p['id']}")]
        )

    if not buttons:
        await send_with_fresh_keyboard(
            message,
            f"На складе **«{ws['title']}»** нет сотрудников для отвязки.",
            has_workspace=True,
        )
        return

    buttons.append([InlineKeyboardButton(text="❌ Отмена", callback_data="cancel")])
    await message.answer(
        f"Выберите сотрудника или приглашение для отвязки от склада **«{ws['title']}»**:",
        reply_markup=InlineKeyboardMarkup(inline_keyboard=buttons),
        parse_mode="Markdown",
    )


@dp.callback_query(F.data.startswith("unlink_u:"))
async def process_unlink_member(callback: types.CallbackQuery):
    _, ws_id_str, target_id_str = callback.data.split(":", 2)
    ws_id = int(ws_id_str)
    target_tg_id = int(target_id_str)
    user = callback.from_user
    is_admin = user.id in ADMIN_IDS

    ws = db_get_workspace_info(ws_id)
    if not ws:
        await callback.message.edit_text("Склад не найден.")
        await callback.answer()
        return
    if not is_admin and user.id != ws["owner_tg_id"]:
        await callback.answer("Недостаточно прав", show_alert=True)
        return

    status, label = db_remove_member(ws_id, target_tg_id)
    if status == "owner":
        await callback.answer("Нельзя отвязать владельца склада", show_alert=True)
        return
    if status != "ok":
        await callback.message.edit_text("Пользователь уже не привязан к этому складу.")
        await callback.answer()
        return

    await callback.message.edit_text(
        f"✅ Сотрудник **{label}** отвязан от склада **«{ws['title']}»**.",
        parse_mode="Markdown",
    )
    await send_with_fresh_keyboard(
        callback.message,
        "Главное меню:",
        has_workspace=True,
        parse_mode=None,
    )
    await callback.answer()

    try:
        await bot.send_message(
            target_tg_id,
            f"ℹ️ Вы отключены от склада **«{ws['title']}»**.",
            parse_mode="Markdown",
        )
    except Exception:
        pass


@dp.callback_query(F.data.startswith("unlink_p:"))
async def process_unlink_pending(callback: types.CallbackQuery):
    _, ws_id_str, inv_id_str = callback.data.split(":", 2)
    ws_id = int(ws_id_str)
    inv_id = int(inv_id_str)
    user = callback.from_user
    is_admin = user.id in ADMIN_IDS

    ws = db_get_workspace_info(ws_id)
    if not ws:
        await callback.message.edit_text("Склад не найден.")
        await callback.answer()
        return
    if not is_admin and user.id != ws["owner_tg_id"]:
        await callback.answer("Недостаточно прав", show_alert=True)
        return

    uname = db_remove_pending_invite(ws_id, inv_id)
    if not uname:
        await callback.message.edit_text("Приглашение уже удалено.")
    else:
        await callback.message.edit_text(
            f"✅ Приглашение для **@{uname}** на склад **«{ws['title']}»** отменено.",
            parse_mode="Markdown",
        )
    await send_with_fresh_keyboard(
        callback.message,
        "Главное меню:",
        has_workspace=True,
        parse_mode=None,
    )
    await callback.answer()


@dp.message(F.text == "👑 Все склады (Админ)")
async def admin_list_workspaces(message: types.Message):
    if message.from_user.id not in ADMIN_IDS:
        return
    rows = db_list_workspaces()
    if not rows:
        await message.answer("Пока не создано ни одного рабочего пространства.")
        return

    lines = [f"👑 **Все склады (Админ)** · сборка `{BOT_VERSION}`:\n"]
    buttons = []
    for r in rows[:20]:
        lines.append(
            f"• **{r['title']}** · владелец `{r['owner_tg_id']}` · "
            f"участников: `{r['users_count']}` · операций: `{r['ops_count']}`"
        )
        buttons.append(
            [InlineKeyboardButton(text=f"Переключиться на {r['title']}", callback_data=f"switch_ws:{r['id']}")]
        )
    buttons.append([InlineKeyboardButton(text="❌ Закрыть", callback_data="cancel")])

    await message.answer(
        "\n".join(lines),
        reply_markup=InlineKeyboardMarkup(inline_keyboard=buttons),
        parse_mode="Markdown",
    )


@dp.callback_query(F.data.startswith("switch_ws:"))
async def admin_switch_workspace(callback: types.CallbackQuery):
    if callback.from_user.id not in ADMIN_IDS:
        await callback.answer("Нет доступа", show_alert=True)
        return
    ws_id = int(callback.data.split(":", 1)[1])
    await db_upsert_user(callback.from_user.id, callback.from_user.username, callback.from_user.full_name or "")
    if db_set_active_workspace(callback.from_user.id, ws_id, is_admin=True):
        ws = db_get_active_workspace(callback.from_user.id)
        title = ws["title"] if ws else "Склад"
        await callback.message.edit_text(
            f"✅ Вы переключились на склад **«{title}»**.",
            parse_mode="Markdown",
        )
        await send_with_fresh_keyboard(
            callback.message,
            "Главное меню:",
            has_workspace=True,
            parse_mode=None,
        )
    else:
        await callback.message.edit_text("Склад не найден.")
    await callback.answer()


# --- ЗАГРУЗКА ОСТАТКОВ ИЗ EXCEL ---
@dp.message(F.text == "📤 Загрузить остатки из Excel")
async def start_import_stock_file(message: types.Message, state: FSMContext):
    ws = await ensure_workspace(message)
    if not ws:
        return
    await state.clear()
    await state.update_data(workspace_id=ws["id"], workspace_title=ws["title"])
    await state.set_state(ImportStockStates.waiting_file)
    await message.answer(
        f"📤 **Загрузка остатков на склад «{ws['title']}»**\n\n"
        f"Отправьте файл `.xlsx` (или `.csv`) со списком позиций.\n"
        f"Колонки могут называться: **Категория | Наименование | Остаток | Ед. измерения** "
        f"(либо идти в этом порядке без шапки).\n"
        f"• Новые позиции добавятся в справочник и получат стартовый остаток.\n"
        f"• Существующие позиции выровняются до указанного в файле остатка.",
        reply_markup=get_cancel_keyboard(),
        parse_mode="Markdown",
    )


@dp.message(ImportStockStates.waiting_file, F.document)
async def process_import_stock_file(message: types.Message, state: FSMContext):
    doc = message.document
    fname = doc.file_name or "import.xlsx"
    if not (fname.lower().endswith(".xlsx") or fname.lower().endswith(".csv")):
        await message.answer(
            "Пожалуйста, отправьте файл в формате `.xlsx` или `.csv` либо нажмите «❌ Отмена».",
            reply_markup=get_cancel_keyboard(),
            parse_mode="Markdown",
        )
        return

    data = await state.get_data()
    ws_id = int(data["workspace_id"])
    ws_title = data["workspace_title"]
    user = message.from_user
    actor_name = user.full_name or (f"@{user.username}" if user.username else "Пользователь")

    buf = io.BytesIO()
    await bot.download(doc, destination=buf)
    parsed_rows, parse_errors = parse_uploaded_stock_file(buf.getvalue(), fname)

    if not parsed_rows:
        err_txt = "\n".join(f"• {e}" for e in parse_errors[:10])
        await message.answer(
            f"❌ Не удалось загрузить позиции из файла.\n{err_txt if err_txt else 'Проверьте колонки и значения.'}",
            reply_markup=get_cancel_keyboard(),
        )
        return

    stats = db_import_stock_rows(ws_id, parsed_rows, actor_name, user.id)
    await state.clear()

    lines = [
        f"✅ **Загрузка остатков на склад «{ws_title}» завершена!**",
        f"• Обработано строк: **{stats['total']}**",
        f"• Новых позиций добавлено: **{stats['added_catalog']}**",
        f"• Обновлено / скорректировано позиций: **{stats['adjusted_items']}**",
    ]
    if stats["zero_diff_items"]:
        lines.append(f"• Уже совпадали с остатком: **{stats['zero_diff_items']}**")
    if parse_errors:
        lines.append(f"\n⚠️ Пропущено строк с ошибками: {len(parse_errors)}")
        for e in parse_errors[:5]:
            lines.append(f"• {e}")

    await send_with_fresh_keyboard(
        message,
        "\n".join(lines),
        has_workspace=True,
    )


@dp.message(ImportStockStates.waiting_file)
async def waiting_import_file_fallback(message: types.Message):
    await message.answer(
        "Прикрепите файл `.xlsx` или `.csv` как документ либо нажмите «❌ Отмена».",
        reply_markup=get_cancel_keyboard(),
        parse_mode="Markdown",
    )


# --- ДОБАВЛЕНИЕ И РЕДАКТИРОВАНИЕ ПОЗИЦИЙ ---
@dp.message(F.text == "📝 Добавить позицию")
async def start_add_catalog_item(message: types.Message, state: FSMContext):
    ws = await ensure_workspace(message)
    if not ws:
        return
    await state.clear()
    await state.update_data(workspace_id=ws["id"], workspace_title=ws["title"])
    await state.set_state(AddCatalogStates.entering_name)
    await message.answer(
        "Введите **наименование** новой позиции:",
        reply_markup=get_cancel_keyboard(),
        parse_mode="Markdown",
    )


@dp.message(AddCatalogStates.entering_name)
async def process_catalog_name(message: types.Message, state: FSMContext):
    item_name = " ".join((message.text or "").strip().split())
    if not item_name:
        await message.answer("Наименование не может быть пустым. Введите название позиции:", reply_markup=get_cancel_keyboard())
        return
    data = await state.get_data()
    ws_id = int(data["workspace_id"])
    if db_catalog_name_exists(ws_id, item_name):
        await message.answer(
            f"Позиция **«{item_name}»** уже есть в справочнике. Введите другое наименование:",
            reply_markup=get_cancel_keyboard(),
            parse_mode="Markdown",
        )
        return
    catalog = db_get_catalog(ws_id)
    existing_categories = []
    for item in catalog:
        c = item["category"].strip()
        if c and c not in existing_categories:
            existing_categories.append(c)
    await state.update_data(new_name=item_name, categories=existing_categories)
    await state.set_state(AddCatalogStates.entering_unit)
    unit_markup = InlineKeyboardMarkup(
        inline_keyboard=[
            [
                InlineKeyboardButton(text="литр", callback_data="unit:литр"),
                InlineKeyboardButton(text="шт", callback_data="unit:шт"),
                InlineKeyboardButton(text="кг", callback_data="unit:кг"),
                InlineKeyboardButton(text="м", callback_data="unit:м"),
            ],
            [InlineKeyboardButton(text="❌ Отмена", callback_data="cancel")],
        ]
    )
    await message.answer(
        f"Позиция: **{item_name}**\nВыберите **единицу измерения** кнопкой или введите текстом:",
        reply_markup=unit_markup,
        parse_mode="Markdown",
    )


async def ask_for_category(target_message: types.Message, state: FSMContext, unit: str, edit: bool = False):
    await state.update_data(new_unit=unit)
    await state.set_state(AddCatalogStates.entering_category)
    data = await state.get_data()
    categories = data.get("categories", [])
    rows = [[InlineKeyboardButton(text=cat, callback_data=f"cat:{cat}")] for cat in categories]
    rows.append([InlineKeyboardButton(text="❌ Отмена", callback_data="cancel")])
    cat_markup = InlineKeyboardMarkup(inline_keyboard=rows)
    text = f"Ед. измерения: **{unit}**\nВыберите **категорию** из списка или введите новую текстом:"
    if edit:
        await target_message.edit_text(text, reply_markup=cat_markup, parse_mode="Markdown")
    else:
        await target_message.answer(text, reply_markup=cat_markup, parse_mode="Markdown")


@dp.callback_query(AddCatalogStates.entering_unit, F.data.startswith("unit:"))
async def process_unit_callback(callback: types.CallbackQuery, state: FSMContext):
    unit = callback.data.split(":", 1)[1]
    await ask_for_category(callback.message, state, unit, edit=True)
    await callback.answer()


@dp.message(AddCatalogStates.entering_unit)
async def process_unit_text(message: types.Message, state: FSMContext):
    unit = (message.text or "").strip()
    if not unit:
        await message.answer("Укажите единицу измерения текстом или выберите кнопку выше:", reply_markup=get_cancel_keyboard())
        return
    await ask_for_category(message, state, unit, edit=False)


async def save_catalog_item(message_or_call, state: FSMContext, category: str):
    data = await state.get_data()
    ws_id = int(data["workspace_id"])
    name = data["new_name"]
    unit = data["new_unit"]
    cat_clean = " ".join((category or "").strip().split())
    if not cat_clean:
        msg = "Категория не может быть пустой. Введите категорию текстом или выберите из списка:"
        if isinstance(message_or_call, types.CallbackQuery):
            await message_or_call.message.answer(msg, reply_markup=get_cancel_keyboard())
            await message_or_call.answer()
        else:
            await message_or_call.answer(msg, reply_markup=get_cancel_keyboard())
        return
    try:
        db_add_catalog_item(ws_id, name, unit, cat_clean)
        result_text = (
            f"✅ **Новая позиция добавлена в справочник!**\n"
            f"• Наименование: {name}\n• Ед. измерения: {unit}\n• Категория: {cat_clean}"
        )
    except Exception as e:
        result_text = f"❌ Ошибка при сохранении в справочник: {e}"
    await state.clear()
    if isinstance(message_or_call, types.CallbackQuery):
        await message_or_call.message.edit_text(result_text, parse_mode="Markdown")
        await send_with_fresh_keyboard(message_or_call.message, "Главное меню:", has_workspace=True, parse_mode=None)
        await message_or_call.answer()
    else:
        await send_with_fresh_keyboard(
            message_or_call,
            result_text,
            has_workspace=True,
        )


@dp.callback_query(AddCatalogStates.entering_category, F.data.startswith("cat:"))
async def process_category_callback(callback: types.CallbackQuery, state: FSMContext):
    await save_catalog_item(callback, state, callback.data.split(":", 1)[1])


@dp.message(AddCatalogStates.entering_category)
async def process_category_text(message: types.Message, state: FSMContext):
    await save_catalog_item(message, state, message.text or "")


@dp.message(F.text == "✏️ Редактировать позицию")
async def start_edit_catalog_item(message: types.Message, state: FSMContext):
    ws = await ensure_workspace(message)
    if not ws:
        return

    catalog = db_get_catalog(ws["id"])
    if not catalog:
        await send_with_fresh_keyboard(
            message,
            "В справочнике пока нет позиций для редактирования.",
            has_workspace=True,
            parse_mode=None,
        )
        return

    await state.clear()
    buttons = [
        [InlineKeyboardButton(text=f"{it['name']} ({it['unit']}, {it['category']})", callback_data=f"edititem:{ws['id']}:{it['id']}")]
        for it in catalog
    ]
    buttons.append([InlineKeyboardButton(text="❌ Отмена", callback_data="cancel")])
    await message.answer(
        f"Выберите позицию на складе **«{ws['title']}»**, которую хотите отредактировать:",
        reply_markup=InlineKeyboardMarkup(inline_keyboard=buttons),
        parse_mode="Markdown",
    )


@dp.callback_query(F.data.startswith("edititem:"))
async def process_choose_item_to_edit(callback: types.CallbackQuery, state: FSMContext):
    _, ws_id_str, item_id_str = callback.data.split(":", 2)
    ws_id = int(ws_id_str)
    item_id = int(item_id_str)

    item = db_get_catalog_item(ws_id, item_id)
    if not item:
        await callback.message.edit_text("Позиция не найдена.")
        await callback.answer()
        return

    catalog = db_get_catalog(ws_id)
    existing_categories = []
    for c_item in catalog:
        c = c_item["category"].strip()
        if c and c not in existing_categories:
            existing_categories.append(c)

    await state.update_data(
        workspace_id=ws_id,
        item_id=item_id,
        old_name=item["name"],
        old_unit=item["unit"],
        old_category=item["category"],
        categories=existing_categories,
    )
    await state.set_state(EditCatalogStates.entering_name)

    markup = InlineKeyboardMarkup(
        inline_keyboard=[
            [InlineKeyboardButton(text=f"Оставить «{item['name']}»", callback_data="editname:keep")],
            [InlineKeyboardButton(text="❌ Отмена", callback_data="cancel")],
        ]
    )
    await callback.message.edit_text(
        f"Редактирование позиции **«{item['name']}»**\n"
        f"Введите **новое наименование** или нажмите кнопку, чтобы оставить текущее:",
        reply_markup=markup,
        parse_mode="Markdown",
    )
    await callback.answer()


async def ask_edit_unit(target_message: types.Message, state: FSMContext, name: str, edit: bool = False):
    await state.update_data(edit_name=name)
    await state.set_state(EditCatalogStates.entering_unit)
    data = await state.get_data()
    old_unit = data.get("old_unit", "шт")

    markup = InlineKeyboardMarkup(
        inline_keyboard=[
            [
                InlineKeyboardButton(text="литр", callback_data="eunit:литр"),
                InlineKeyboardButton(text="шт", callback_data="eunit:шт"),
                InlineKeyboardButton(text="кг", callback_data="eunit:кг"),
                InlineKeyboardButton(text="м", callback_data="eunit:м"),
            ],
            [InlineKeyboardButton(text=f"Оставить «{old_unit}»", callback_data="eunit:__KEEP__")],
            [InlineKeyboardButton(text="❌ Отмена", callback_data="cancel")],
        ]
    )
    text = f"Наименование: **{name}**\nВыберите **единицу измерения** или введите свою текстом:"
    if edit:
        await target_message.edit_text(text, reply_markup=markup, parse_mode="Markdown")
    else:
        await target_message.answer(text, reply_markup=markup, parse_mode="Markdown")


@dp.callback_query(EditCatalogStates.entering_name, F.data == "editname:keep")
async def keep_edit_name_cb(callback: types.CallbackQuery, state: FSMContext):
    data = await state.get_data()
    await ask_edit_unit(callback.message, state, data["old_name"], edit=True)
    await callback.answer()


@dp.message(EditCatalogStates.entering_name)
async def process_edit_item_name(message: types.Message, state: FSMContext):
    name = " ".join((message.text or "").strip().split())
    if not name:
        await message.answer("Наименование не может быть пустым. Введите название или нажмите кнопку:", reply_markup=get_cancel_keyboard())
        return
    data = await state.get_data()
    ws_id = int(data["workspace_id"])
    item_id = int(data["item_id"])
    if db_catalog_name_exists(ws_id, name, exclude_id=item_id):
        await message.answer(
            f"Позиция **«{name}»** уже есть в справочнике. Введите другое наименование:",
            reply_markup=get_cancel_keyboard(),
            parse_mode="Markdown",
        )
        return
    await ask_edit_unit(message, state, name, edit=False)


async def ask_edit_category(target_message: types.Message, state: FSMContext, unit: str, edit: bool = False):
    await state.update_data(edit_unit=unit)
    await state.set_state(EditCatalogStates.entering_category)
    data = await state.get_data()
    categories = data.get("categories", [])
    old_cat = data.get("old_category", "")

    rows = [[InlineKeyboardButton(text=cat, callback_data=f"ecat:{cat}")] for cat in categories]
    if old_cat:
        rows.append([InlineKeyboardButton(text=f"Оставить «{old_cat}»", callback_data="ecat:__KEEP__")])
    rows.append([InlineKeyboardButton(text="❌ Отмена", callback_data="cancel")])
    markup = InlineKeyboardMarkup(inline_keyboard=rows)
    text = f"Ед. измерения: **{unit}**\nВыберите **категорию** из списка или введите новую текстом:"
    if edit:
        await target_message.edit_text(text, reply_markup=markup, parse_mode="Markdown")
    else:
        await target_message.answer(text, reply_markup=markup, parse_mode="Markdown")


@dp.callback_query(EditCatalogStates.entering_unit, F.data.startswith("eunit:"))
async def process_edit_unit_cb(callback: types.CallbackQuery, state: FSMContext):
    val = callback.data.split(":", 1)[1]
    data = await state.get_data()
    unit = data["old_unit"] if val == "__KEEP__" else val
    await ask_edit_category(callback.message, state, unit, edit=True)
    await callback.answer()


@dp.message(EditCatalogStates.entering_unit)
async def process_edit_unit_text(message: types.Message, state: FSMContext):
    unit = (message.text or "").strip()
    if not unit:
        await message.answer("Укажите единицу измерения текстом или выберите кнопку выше:", reply_markup=get_cancel_keyboard())
        return
    await ask_edit_category(message, state, unit, edit=False)


async def finish_edit_catalog_item(message_or_call, state: FSMContext, raw_category: str):
    data = await state.get_data()
    ws_id = int(data["workspace_id"])
    item_id = int(data["item_id"])
    new_name = data["edit_name"]
    new_unit = data["edit_unit"]
    cat_clean = data["old_category"] if raw_category == "__KEEP__" else " ".join((raw_category or "").strip().split())

    if not cat_clean:
        msg = "Категория не может быть пустой. Введите категорию текстом или выберите из списка:"
        if isinstance(message_or_call, types.CallbackQuery):
            await message_or_call.message.answer(msg, reply_markup=get_cancel_keyboard())
            await message_or_call.answer()
        else:
            await message_or_call.answer(msg, reply_markup=get_cancel_keyboard())
        return

    status, _ = db_update_catalog_item(ws_id, item_id, new_name, new_unit, cat_clean)
    await state.clear()

    if status == "not_found":
        res = "❌ Позиция не найдена."
    elif status == "duplicate":
        res = f"❌ Позиция с названием **«{new_name}»** уже существует."
    else:
        res = (
            f"✅ **Позиция обновлена!**\n"
            f"• Наименование: {new_name}\n"
            f"• Ед. измерения: {new_unit}\n"
            f"• Категория: {cat_clean}"
        )

    if isinstance(message_or_call, types.CallbackQuery):
        await message_or_call.message.edit_text(res, parse_mode="Markdown")
        await send_with_fresh_keyboard(message_or_call.message, "Главное меню:", has_workspace=True, parse_mode=None)
        await message_or_call.answer()
    else:
        await send_with_fresh_keyboard(message_or_call, res, has_workspace=True)


@dp.callback_query(EditCatalogStates.entering_category, F.data.startswith("ecat:"))
async def process_edit_category_cb(callback: types.CallbackQuery, state: FSMContext):
    await finish_edit_catalog_item(callback, state, callback.data.split(":", 1)[1])


@dp.message(EditCatalogStates.entering_category)
async def process_edit_category_text(message: types.Message, state: FSMContext):
    await finish_edit_catalog_item(message, state, message.text or "")


# --- ПРИХОД, РАСХОД И БЫСТРЫЙ ПОВТОР ---
@dp.message(F.text.in_(["➕ Приход", "➖ Расход"]))
async def start_operation(message: types.Message, state: FSMContext):
    ws = await ensure_workspace(message)
    if not ws:
        return
    op_type = "Приход" if "Приход" in message.text else "Расход"
    await state.clear()
    await state.update_data(op_type=op_type, workspace_id=ws["id"], workspace_title=ws["title"])
    today_str = datetime.now().strftime("%d.%m.%Y")
    date_markup = InlineKeyboardMarkup(
        inline_keyboard=[
            [InlineKeyboardButton(text=f"Сегодня ({today_str})", callback_data=f"date:{today_str}")],
            [InlineKeyboardButton(text="📅 Другая дата", callback_data="date:custom")],
            [InlineKeyboardButton(text="❌ Отмена", callback_data="cancel")],
        ]
    )
    await state.set_state(OperationStates.choosing_date)
    await message.answer(f"Операция: **{op_type}**\nВыберите дату операции:", reply_markup=date_markup, parse_mode="Markdown")


async def show_items_or_categories(target_message: types.Message, state: FSMContext, op_date: str, edit: bool = False):
    await state.update_data(op_date=op_date)
    data = await state.get_data()
    ws_id = int(data["workspace_id"])
    op_type = data["op_type"]
    catalog = db_get_catalog(ws_id)
    if not catalog:
        await target_message.answer(
            "В справочнике пока нет товаров. Нажмите «📝 Добавить позицию» или «📤 Загрузить остатки из Excel»."
        )
        return
    await state.update_data(catalog=catalog)
    categories = []
    for c in catalog:
        if c["category"] not in categories:
            categories.append(c["category"])
    if len(categories) > 1:
        buttons = [[InlineKeyboardButton(text=cat, callback_data=f"opcat:{cat}")] for cat in categories]
        buttons.append([InlineKeyboardButton(text="📋 Все позиции списком", callback_data="opcat:__ALL__")])
        buttons.append([InlineKeyboardButton(text="🔍 Поиск позиции", callback_data="op_search")])
        buttons.append([InlineKeyboardButton(text="❌ Отмена", callback_data="cancel")])
        await state.set_state(OperationStates.choosing_category)
        text = f"Операция: **{op_type}** · Дата: **{op_date}**\nВыберите категорию:"
    else:
        buttons = [[InlineKeyboardButton(text=item["name"], callback_data=f"item:{item['name']}")] for item in catalog[:30]]
        if len(catalog) > 8:
            buttons.append([InlineKeyboardButton(text="🔍 Поиск позиции", callback_data="op_search")])
        buttons.append([InlineKeyboardButton(text="❌ Отмена", callback_data="cancel")])
        await state.set_state(OperationStates.choosing_item)
        text = f"Операция: **{op_type}** · Дата: **{op_date}**\nВыберите позицию:"
    markup = InlineKeyboardMarkup(inline_keyboard=buttons)
    if edit:
        await target_message.edit_text(text, reply_markup=markup, parse_mode="Markdown")
    else:
        await target_message.answer(text, reply_markup=markup, parse_mode="Markdown")


@dp.callback_query(F.data == "op_search")
async def start_item_search_cb(callback: types.CallbackQuery, state: FSMContext):
    await state.set_state(OperationStates.searching_item)
    await callback.message.edit_text("🔍 Введите часть названия позиции для поиска:")
    await callback.answer()


@dp.message(OperationStates.searching_item)
async def process_item_search_text(message: types.Message, state: FSMContext):
    query = (message.text or "").strip().lower()
    if not query:
        await message.answer("Введите часть названия позиции:", reply_markup=get_cancel_keyboard())
        return
    data = await state.get_data()
    catalog = data.get("catalog", [])
    matches = [x for x in catalog if query in x["name"].lower()]
    if not matches:
        await message.answer(
            f"По запросу **«{message.text.strip()}»** ничего не найдено. Введите другой запрос или нажмите «❌ Отмена»:",
            reply_markup=get_cancel_keyboard(),
            parse_mode="Markdown",
        )
        return
    buttons = [[InlineKeyboardButton(text=x["name"], callback_data=f"item:{x['name']}")] for x in matches[:25]]
    buttons.append([InlineKeyboardButton(text="🔍 Искать снова", callback_data="op_search")])
    buttons.append([InlineKeyboardButton(text="❌ Отмена", callback_data="cancel")])
    await state.set_state(OperationStates.choosing_item)
    await message.answer(
        f"Найдено позиций: **{len(matches)}**. Выберите позицию:",
        reply_markup=InlineKeyboardMarkup(inline_keyboard=buttons),
        parse_mode="Markdown",
    )


@dp.callback_query(OperationStates.choosing_date, F.data.startswith("date:"))
async def process_date_choice(callback: types.CallbackQuery, state: FSMContext):
    val = callback.data.split(":", 1)[1]
    if val == "custom":
        await state.set_state(OperationStates.entering_custom_date)
        await callback.message.edit_text("Введите дату в формате `ДД.ММ.ГГГГ` (например: `28.11.2026`):", parse_mode="Markdown")
    else:
        await show_items_or_categories(callback.message, state, val, edit=True)
    await callback.answer()


@dp.message(OperationStates.entering_custom_date)
async def process_custom_date(message: types.Message, state: FSMContext):
    raw = (message.text or "").strip()
    try:
        parsed = datetime.strptime(raw, "%d.%m.%Y")
        date_str = parsed.strftime("%d.%m.%Y")
    except ValueError:
        await message.answer("Неверный формат даты. Введите как `ДД.ММ.ГГГГ` (например, `28.11.2026`):", parse_mode="Markdown")
        return
    await show_items_or_categories(message, state, date_str, edit=False)


@dp.callback_query(OperationStates.choosing_category, F.data.startswith("opcat:"))
async def process_op_category(callback: types.CallbackQuery, state: FSMContext):
    cat = callback.data.split(":", 1)[1]
    data = await state.get_data()
    catalog = data.get("catalog", [])
    items = catalog if cat == "__ALL__" else [x for x in catalog if x["category"] == cat]
    buttons = [[InlineKeyboardButton(text=x["name"], callback_data=f"item:{x['name']}")] for x in items[:30]]
    buttons.append([InlineKeyboardButton(text="🔍 Поиск позиции", callback_data="op_search")])
    buttons.append([InlineKeyboardButton(text="⬅️ Назад к категориям", callback_data="back_to_cats")])
    buttons.append([InlineKeyboardButton(text="❌ Отмена", callback_data="cancel")])
    await state.set_state(OperationStates.choosing_item)
    await callback.message.edit_text(
        f"Категория: **{cat if cat != '__ALL__' else 'Все позиции'}**\nВыберите номенклатуру:",
        reply_markup=InlineKeyboardMarkup(inline_keyboard=buttons),
        parse_mode="Markdown",
    )
    await callback.answer()


@dp.callback_query(OperationStates.choosing_item, F.data == "back_to_cats")
async def back_to_categories(callback: types.CallbackQuery, state: FSMContext):
    data = await state.get_data()
    await show_items_or_categories(callback.message, state, data["op_date"], edit=True)
    await callback.answer()


@dp.callback_query(OperationStates.choosing_item, F.data.startswith("item:"))
async def process_item_choice(callback: types.CallbackQuery, state: FSMContext):
    item_name = callback.data.split(":", 1)[1]
    data = await state.get_data()
    catalog = data.get("catalog", [])
    item_info = next((x for x in catalog if x["name"] == item_name), {"name": item_name, "unit": "", "category": ""})
    await state.update_data(item_name=item_name, item_unit=item_info.get("unit", ""), item_category=item_info.get("category", ""))
    await state.set_state(OperationStates.entering_quantity)
    unit_hint = f" ({item_info['unit']})" if item_info.get("unit") else ""
    await callback.message.edit_text(f"Выбрано: **{item_name}**{unit_hint}\nВведите количество (числом):", parse_mode="Markdown")
    await callback.answer()


@dp.message(OperationStates.entering_quantity)
async def process_quantity(message: types.Message, state: FSMContext):
    qty = parse_number(message.text or "")
    if qty is None or qty <= 0:
        await message.answer("Введите корректное положительное число (например: 50 или 12.5):")
        return
    data = await state.get_data()
    op_type = data["op_type"]
    item_name = data["item_name"]
    ws_id = int(data["workspace_id"])
    if op_type == "Расход":
        stock_list = db_get_stock(ws_id)
        item_stock = next((x for x in stock_list if x["name"].lower() == item_name.lower()), None)
        cur_bal = item_stock["balance"] if item_stock else 0.0
        unit = item_stock["unit"] if item_stock else data.get("item_unit", "")
        if qty > cur_bal:
            warn_markup = InlineKeyboardMarkup(
                inline_keyboard=[
                    [InlineKeyboardButton(text="⚠️ Всё равно списать", callback_data="force_qty_ok")],
                    [InlineKeyboardButton(text="✏️ Ввести другое число", callback_data="reenter_qty")],
                    [InlineKeyboardButton(text="❌ Отмена", callback_data="cancel")],
                ]
            )
            await state.update_data(pending_qty=qty)
            await message.answer(
                f"⚠️ **Недостаточно остатка на складе!**\n• Текущий остаток по «{item_name}»: **{cur_bal:g}** {unit}\n• Вы пытаетесь списать: **{qty:g}** {unit}",
                reply_markup=warn_markup,
                parse_mode="Markdown",
            )
            return
    await ask_comment_step(message, state, qty)


@dp.callback_query(OperationStates.entering_quantity, F.data == "reenter_qty")
async def reenter_quantity_cb(callback: types.CallbackQuery, state: FSMContext):
    data = await state.get_data()
    await callback.message.edit_text(f"Позиция: **{data['item_name']}**\nВведите новое количество:", parse_mode="Markdown")
    await callback.answer()


@dp.callback_query(OperationStates.entering_quantity, F.data == "force_qty_ok")
async def force_qty_ok_cb(callback: types.CallbackQuery, state: FSMContext):
    data = await state.get_data()
    qty = float(data.get("pending_qty", 0.0))
    await callback.answer()
    await ask_comment_step(callback.message, state, qty, edit=True)


async def ask_comment_step(target_message: types.Message, state: FSMContext, qty: float, edit: bool = False):
    await state.update_data(quantity=qty)
    await state.set_state(OperationStates.entering_comment)
    skip_markup = InlineKeyboardMarkup(
        inline_keyboard=[
            [InlineKeyboardButton(text="Без комментария ➡️", callback_data="skip_comment")],
            [InlineKeyboardButton(text="❌ Отмена", callback_data="cancel")],
        ]
    )
    text = f"Количество: **{qty:g}**\nВведите комментарий (например, номер авто или накладной) или нажмите кнопку:"
    if edit:
        await target_message.edit_text(text, reply_markup=skip_markup, parse_mode="Markdown")
    else:
        await target_message.answer(text, reply_markup=skip_markup, parse_mode="Markdown")


async def show_confirmation(target_message: types.Message, state: FSMContext, comment: str, edit: bool = False):
    await state.update_data(comment=comment)
    await state.set_state(OperationStates.confirming)
    data = await state.get_data()
    unit = f" {data.get('item_unit')}" if data.get("item_unit") else ""
    text = (
        f"🔎 **Проверьте данные перед записью:**\n"
        f"• Дата: **{data['op_date']}**\n• Операция: **{data['op_type']}**\n"
        f"• Позиция: **{data['item_name']}**\n• Количество: **{data['quantity']:g}**{unit}\n"
        f"• Комментарий: {comment if comment else '—'}"
    )
    confirm_markup = InlineKeyboardMarkup(
        inline_keyboard=[
            [
                InlineKeyboardButton(text="✅ Записать", callback_data="confirm_save"),
                InlineKeyboardButton(text="❌ Отмена", callback_data="cancel"),
            ]
        ]
    )
    if edit:
        await target_message.edit_text(text, reply_markup=confirm_markup, parse_mode="Markdown")
    else:
        await target_message.answer(text, reply_markup=confirm_markup, parse_mode="Markdown")


@dp.callback_query(OperationStates.entering_comment, F.data == "skip_comment")
async def skip_comment_handler(callback: types.CallbackQuery, state: FSMContext):
    await show_confirmation(callback.message, state, "", edit=True)
    await callback.answer()


@dp.message(OperationStates.entering_comment)
async def comment_text_handler(message: types.Message, state: FSMContext):
    await show_confirmation(message, state, (message.text or "").strip(), edit=False)


@dp.callback_query(OperationStates.confirming, F.data == "confirm_save")
async def confirm_save_handler(callback: types.CallbackQuery, state: FSMContext):
    data = await state.get_data()
    ws_id = int(data["workspace_id"])
    op_type = data["op_type"]
    item_name = data["item_name"]
    qty = float(data["quantity"])
    op_date = data["op_date"]
    unit = data.get("item_unit", "")
    category = data.get("item_category", "")
    comment = data.get("comment", "")

    user = callback.from_user
    actor_name = user.full_name or (f"@{user.username}" if user.username else "Пользователь")
    low_stock_alert = ""

    try:
        db_add_movement(ws_id, op_date, op_type, item_name, qty, unit, category, comment, actor_name, user.id)
        if op_type == "Расход":
            stock_list = db_get_stock(ws_id)
            item_stock = next((x for x in stock_list if x["name"].lower() == item_name.lower()), None)
            rem_bal = item_stock["balance"] if item_stock else 0.0
            if rem_bal <= LOW_STOCK_THRESHOLD:
                low_stock_alert = f"\n⚠️ **Позиция закончилась (остаток: {rem_bal:g} {unit})**".strip()
        summary = (
            f"✅ **Запись добавлена!**\n"
            f"• Дата: {op_date}\n• Операция: {op_type}\n• Позиция: {item_name}\n"
            f"• Количество: {qty:g} {unit}".strip() + "\n"
            f"• Комментарий: {comment if comment else '—'}"
            f"{low_stock_alert}"
        )
    except Exception as e:
        summary = f"❌ Ошибка при записи: {e}"

    await state.clear()
    await callback.message.edit_text(summary, parse_mode="Markdown")
    await send_with_fresh_keyboard(callback.message, "Главное меню:", has_workspace=True, parse_mode=None)
    await callback.answer()


# --- БЫСТРЫЙ ПОВТОР ПОСЛЕДНЕЙ ОПЕРАЦИИ ---
@dp.message(F.text == "🔁 Повторить последнюю")
async def start_repeat_last_operation(message: types.Message, state: FSMContext):
    ws = await ensure_workspace(message)
    if not ws:
        return
    rows = db_get_recent_movements(ws["id"], limit=1)
    if not rows:
        await send_with_fresh_keyboard(
            message,
            "На этом складе пока нет операций для повтора.",
            has_workspace=True,
            parse_mode=None,
        )
        return

    last = rows[0]
    await state.clear()
    await state.update_data(
        workspace_id=ws["id"],
        op_type=last["op_type"],
        item_name=last["item_name"],
        quantity=float(last["quantity"]),
        item_unit=last["unit"],
        item_category=last["category"],
        comment=last["comment"],
    )
    today_str = datetime.now().strftime("%d.%m.%Y")
    markup = InlineKeyboardMarkup(
        inline_keyboard=[
            [InlineKeyboardButton(text=f"Сегодня ({today_str})", callback_data=f"repdate:{today_str}")],
            [InlineKeyboardButton(text="📅 Другая дата", callback_data="repdate:custom")],
            [InlineKeyboardButton(text="❌ Отмена", callback_data="cancel")],
        ]
    )
    await state.set_state(RepeatOperationStates.choosing_date)
    unit_str = f" {last['unit']}" if last["unit"] else ""
    await message.answer(
        f"🔁 **Повтор последней операции:**\n"
        f"• {last['op_type']}: **{last['item_name']}** — **{float(last['quantity']):g}**{unit_str}\n"
        f"Выберите дату для новой записи:",
        reply_markup=markup,
        parse_mode="Markdown",
    )


@dp.callback_query(RepeatOperationStates.choosing_date, F.data.startswith("repdate:"))
async def process_repeat_date_cb(callback: types.CallbackQuery, state: FSMContext):
    val = callback.data.split(":", 1)[1]
    if val == "custom":
        await state.set_state(RepeatOperationStates.entering_custom_date)
        await callback.message.edit_text("Введите дату в формате `ДД.ММ.ГГГГ` (например: `28.11.2026`):", parse_mode="Markdown")
        await callback.answer()
        return

    await state.update_data(op_date=val)
    await state.set_state(RepeatOperationStates.confirming)
    data = await state.get_data()
    unit_str = f" {data['item_unit']}" if data.get("item_unit") else ""
    confirm_markup = InlineKeyboardMarkup(
        inline_keyboard=[
            [InlineKeyboardButton(text="✅ Записать", callback_data="rep_confirm")],
            [InlineKeyboardButton(text="❌ Отмена", callback_data="cancel")],
        ]
    )
    await callback.message.edit_text(
        f"🔎 **Подтвердите повтор операции:**\n"
        f"• Дата: **{val}**\n"
        f"• Операция: **{data['op_type']}**\n"
        f"• Позиция: **{data['item_name']}**\n"
        f"• Количество: **{float(data['quantity']):g}**{unit_str}\n"
        f"• Комментарий: {data['comment'] if data.get('comment') else '—'}",
        reply_markup=confirm_markup,
        parse_mode="Markdown",
    )
    await callback.answer()


@dp.message(RepeatOperationStates.entering_custom_date)
async def process_repeat_custom_date(message: types.Message, state: FSMContext):
    raw = (message.text or "").strip()
    try:
        date_str = datetime.strptime(raw, "%d.%m.%Y").strftime("%d.%m.%Y")
    except ValueError:
        await message.answer("Неверный формат даты. Введите как `ДД.ММ.ГГГГ`:", parse_mode="Markdown")
        return

    await state.update_data(op_date=date_str)
    await state.set_state(RepeatOperationStates.confirming)
    data = await state.get_data()
    unit_str = f" {data['item_unit']}" if data.get("item_unit") else ""
    confirm_markup = InlineKeyboardMarkup(
        inline_keyboard=[
            [InlineKeyboardButton(text="✅ Записать", callback_data="rep_confirm")],
            [InlineKeyboardButton(text="❌ Отмена", callback_data="cancel")],
        ]
    )
    await message.answer(
        f"🔎 **Подтвердите повтор операции:**\n"
        f"• Дата: **{date_str}**\n"
        f"• Операция: **{data['op_type']}**\n"
        f"• Позиция: **{data['item_name']}**\n"
        f"• Количество: **{float(data['quantity']):g}**{unit_str}\n"
        f"• Комментарий: {data['comment'] if data.get('comment') else '—'}",
        reply_markup=confirm_markup,
        parse_mode="Markdown",
    )


@dp.callback_query(RepeatOperationStates.confirming, F.data == "rep_confirm")
async def confirm_repeat_operation(callback: types.CallbackQuery, state: FSMContext):
    data = await state.get_data()
    user = callback.from_user
    actor_name = user.full_name or (f"@{user.username}" if user.username else "Пользователь")
    db_add_movement(
        int(data["workspace_id"]),
        data["op_date"],
        data["op_type"],
        data["item_name"],
        float(data["quantity"]),
        data.get("item_unit", ""),
        data.get("item_category", ""),
        data.get("comment", ""),
        actor_name,
        user.id,
    )
    await state.clear()
    await callback.message.edit_text("✅ Операция повторена и записана!")
    await send_with_fresh_keyboard(callback.message, "Главное меню:", has_workspace=True, parse_mode=None)
    await callback.answer()


# --- ОТМЕНА ОПЕРАЦИИ, ИСТОРИЯ С ФИЛЬТРАМИ, ОСТАТКИ И EXCEL ---
@dp.message(F.text.in_(["🗑 Отменить операцию", "↩️ Отменить последнюю"]))
async def choose_operation_to_cancel(message: types.Message):
    ws = await ensure_workspace(message)
    if not ws:
        return

    rows = db_get_recent_movements(ws["id"], limit=10)
    if not rows:
        await send_with_fresh_keyboard(
            message,
            "На этом складе пока нет операций для отмены.",
            has_workspace=True,
            parse_mode=None,
        )
        return

    buttons = []
    for r in rows:
        sign = "➕" if r["op_type"] == "Приход" else "➖"
        unit_str = f" {r['unit']}" if r["unit"] else ""
        label = f"{sign} {r['op_date']} · {r['item_name']} · {float(r['quantity']):g}{unit_str}"
        buttons.append([InlineKeyboardButton(text=label[:60], callback_data=f"delop:{ws['id']}:{r['id']}")])
    buttons.append([InlineKeyboardButton(text="❌ Отмена", callback_data="cancel")])

    await message.answer(
        f"Выберите операцию на складе **«{ws['title']}»**, которую нужно отменить:",
        reply_markup=InlineKeyboardMarkup(inline_keyboard=buttons),
        parse_mode="Markdown",
    )


@dp.callback_query(F.data.startswith("delop:"))
async def process_cancel_selected_operation(callback: types.CallbackQuery):
    _, ws_id_str, mov_id_str = callback.data.split(":", 2)
    ws_id = int(ws_id_str)
    mov_id = int(mov_id_str)
    user = callback.from_user
    is_admin = user.id in ADMIN_IDS

    status, info = db_delete_movement_by_id(ws_id, mov_id, user.id, is_admin)
    if status == "not_found":
        await callback.message.edit_text("Эта операция уже была отменена или не найдена.")
        await callback.answer()
        return
    if status == "forbidden":
        await callback.message.edit_text("Вы можете отменить только свою операцию. Обратитесь к администратору.")
        await callback.answer()
        return

    await callback.message.edit_text(f"🗑 Операция отменена (сохранена в журнале отмен): **{info}**", parse_mode="Markdown")
    await send_with_fresh_keyboard(callback.message, "Главное меню:", has_workspace=True, parse_mode=None)
    await callback.answer()


@dp.message(F.text == "📜 Последние операции")
async def recent_movements(message: types.Message, state: FSMContext):
    ws = await ensure_workspace(message)
    if not ws:
        return
    await state.clear()

    markup = InlineKeyboardMarkup(
        inline_keyboard=[
            [
                InlineKeyboardButton(text="🕒 Последние 10", callback_data=f"hist_last:{ws['id']}"),
                InlineKeyboardButton(text="📅 По дате", callback_data=f"hist_date_menu:{ws['id']}"),
            ],
            [InlineKeyboardButton(text="📦 По позиции", callback_data=f"hist_item_menu:{ws['id']}")],
            [InlineKeyboardButton(text="❌ Закрыть", callback_data="cancel")],
        ]
    )
    await message.answer(
        f"📜 **История операций («{ws['title']}»)**\nВыберите способ просмотра:",
        reply_markup=markup,
        parse_mode="Markdown",
    )


@dp.callback_query(F.data.startswith("hist_last:"))
async def hist_show_last_cb(callback: types.CallbackQuery):
    ws_id = int(callback.data.split(":", 1)[1])
    ws = db_get_workspace_info(ws_id)
    if not ws:
        await callback.message.edit_text("Склад не найден.")
        await callback.answer()
        return

    rows = db_get_recent_movements(ws_id, limit=10)
    if not rows:
        await callback.message.edit_text(f"На складе **«{ws['title']}»** операций пока не было.", parse_mode="Markdown")
        await callback.answer()
        return

    is_admin = callback.from_user.id in ADMIN_IDS
    lines = [f"📜 **Последние 10 операций («{ws['title']}»):**\n"]
    lines.extend(format_movements_lines(rows, callback.message.chat, is_admin))
    await callback.message.edit_text("\n".join(lines), parse_mode="Markdown")
    await send_with_fresh_keyboard(
        callback.message,
        "Выберите следующее действие:",
        has_workspace=True,
        is_admin=is_admin,
        in_menu=True,
        parse_mode=None,
    )
    await callback.answer()


@dp.callback_query(F.data.startswith("hist_date_menu:"))
async def hist_date_menu_cb(callback: types.CallbackQuery):
    ws_id = int(callback.data.split(":", 1)[1])
    ws = db_get_workspace_info(ws_id)
    if not ws:
        await callback.message.edit_text("Склад не найден.")
        await callback.answer()
        return

    now = datetime.now()
    today_str = now.strftime("%d.%m.%Y")
    yday_str = (now - timedelta(days=1)).strftime("%d.%m.%Y")
    markup = InlineKeyboardMarkup(
        inline_keyboard=[
            [
                InlineKeyboardButton(text=f"Сегодня ({today_str})", callback_data=f"hdate:{ws_id}:{today_str}"),
                InlineKeyboardButton(text=f"Вчера ({yday_str})", callback_data=f"hdate:{ws_id}:{yday_str}"),
            ],
            [InlineKeyboardButton(text="⌨️ Ввести другую дату", callback_data=f"hdate_custom:{ws_id}")],
            [InlineKeyboardButton(text="⬅️ Назад", callback_data=f"hist_back:{ws_id}")],
        ]
    )
    await callback.message.edit_text(
        f"📅 **История по дате («{ws['title']}»)**\nВыберите дату или введите вручную:",
        reply_markup=markup,
        parse_mode="Markdown",
    )
    await callback.answer()


@dp.callback_query(F.data.startswith("hdate:"))
async def hist_by_date_cb(callback: types.CallbackQuery):
    _, ws_id_str, date_str = callback.data.split(":", 2)
    ws_id = int(ws_id_str)
    ws = db_get_workspace_info(ws_id)
    if not ws:
        await callback.message.edit_text("Склад не найден.")
        await callback.answer()
        return

    is_admin = callback.from_user.id in ADMIN_IDS
    rows = db_get_movements_by_date(ws_id, date_str, limit=30)
    if not rows:
        await callback.message.edit_text(
            f"📅 За **{date_str}** на складе **«{ws['title']}»** операций не найдено.",
            parse_mode="Markdown",
        )
        await send_with_fresh_keyboard(
            callback.message,
            "Выберите следующее действие:",
            has_workspace=True,
            is_admin=is_admin,
            in_menu=True,
            parse_mode=None,
        )
        await callback.answer()
        return

    lines = [f"📅 **Операции за {date_str} («{ws['title']}»):**\n"]
    lines.extend(format_movements_lines(rows, callback.message.chat, is_admin))
    await callback.message.edit_text("\n".join(lines), parse_mode="Markdown")
    await send_with_fresh_keyboard(
        callback.message,
        "Выберите следующее действие:",
        has_workspace=True,
        is_admin=is_admin,
        in_menu=True,
        parse_mode=None,
    )
    await callback.answer()


@dp.callback_query(F.data.startswith("hdate_custom:"))
async def hist_custom_date_start_cb(callback: types.CallbackQuery, state: FSMContext):
    ws_id = int(callback.data.split(":", 1)[1])
    await state.set_state(HistoryFilterStates.entering_date)
    await state.update_data(hist_ws_id=ws_id)
    await callback.message.edit_text(
        "Введите дату в формате `ДД.ММ.ГГГГ` (например, `28.11.2026`):",
        parse_mode="Markdown",
    )
    await callback.answer()


@dp.message(HistoryFilterStates.entering_date)
async def hist_custom_date_process(message: types.Message, state: FSMContext):
    raw = (message.text or "").strip()
    try:
        date_str = datetime.strptime(raw, "%d.%m.%Y").strftime("%d.%m.%Y")
    except ValueError:
        await message.answer(
            "Неверный формат даты. Введите как `ДД.ММ.ГГГГ` (например, `28.11.2026`) или нажмите «❌ Отмена»:",
            reply_markup=get_cancel_keyboard(),
            parse_mode="Markdown",
        )
        return

    data = await state.get_data()
    ws_id = int(data["hist_ws_id"])
    ws = db_get_workspace_info(ws_id)
    await state.clear()
    is_admin = message.from_user.id in ADMIN_IDS

    if not ws:
        await send_with_fresh_keyboard(message, "Склад не найден.", has_workspace=True, parse_mode=None)
        return

    rows = db_get_movements_by_date(ws_id, date_str, limit=30)
    if not rows:
        await send_with_fresh_keyboard(
            message,
            f"📅 За **{date_str}** на складе **«{ws['title']}»** операций не найдено.",
            has_workspace=True,
            is_admin=is_admin,
            in_menu=True,
        )
        return

    lines = [f"📅 **Операции за {date_str} («{ws['title']}»):**\n"]
    lines.extend(format_movements_lines(rows, message.chat, is_admin))
    await send_with_fresh_keyboard(
        message,
        "\n".join(lines),
        has_workspace=True,
        is_admin=is_admin,
        in_menu=True,
    )


@dp.callback_query(F.data.startswith("hist_item_menu:"))
async def hist_item_menu_cb(callback: types.CallbackQuery):
    ws_id = int(callback.data.split(":", 1)[1])
    ws = db_get_workspace_info(ws_id)
    if not ws:
        await callback.message.edit_text("Склад не найден.")
        await callback.answer()
        return

    catalog = db_get_catalog(ws_id)
    if not catalog:
        await callback.message.edit_text("В справочнике пока нет позиций.")
        await callback.answer()
        return

    categories = []
    for c in catalog:
        if c["category"] not in categories:
            categories.append(c["category"])

    buttons: list[list[InlineKeyboardButton]] = []
    if len(categories) > 1:
        for cat in categories[:12]:
            buttons.append([InlineKeyboardButton(text=f"📁 {cat}", callback_data=f"hcat:{ws_id}:{cat}")])
        buttons.append([InlineKeyboardButton(text="📋 Все позиции", callback_data=f"hcat:{ws_id}:__ALL__")])
    else:
        for item in catalog[:25]:
            buttons.append([InlineKeyboardButton(text=item["name"], callback_data=f"hitem:{ws_id}:{item['name']}")])

    buttons.append([InlineKeyboardButton(text="🔍 Поиск позиции", callback_data=f"hsearch:{ws_id}")])
    buttons.append([InlineKeyboardButton(text="⬅️ Назад", callback_data=f"hist_back:{ws_id}")])

    await callback.message.edit_text(
        f"📦 **История по позиции («{ws['title']}»)**\nВыберите категорию или позицию:",
        reply_markup=InlineKeyboardMarkup(inline_keyboard=buttons),
        parse_mode="Markdown",
    )
    await callback.answer()


@dp.callback_query(F.data.startswith("hcat:"))
async def hist_by_cat_cb(callback: types.CallbackQuery):
    _, ws_id_str, cat = callback.data.split(":", 2)
    ws_id = int(ws_id_str)
    catalog = db_get_catalog(ws_id)
    items = catalog if cat == "__ALL__" else [x for x in catalog if x["category"] == cat]
    if not items:
        await callback.answer("В этой категории нет позиций", show_alert=True)
        return

    buttons = [[InlineKeyboardButton(text=x["name"], callback_data=f"hitem:{ws_id}:{x['name']}")] for x in items[:25]]
    buttons.append([InlineKeyboardButton(text="🔍 Поиск позиции", callback_data=f"hsearch:{ws_id}")])
    buttons.append([InlineKeyboardButton(text="⬅️ Назад", callback_data=f"hist_item_menu:{ws_id}")])
    await callback.message.edit_text(
        f"Категория: **{cat if cat != '__ALL__' else 'Все позиции'}**\nВыберите позицию:",
        reply_markup=InlineKeyboardMarkup(inline_keyboard=buttons),
        parse_mode="Markdown",
    )
    await callback.answer()


@dp.callback_query(F.data.startswith("hsearch:"))
async def hist_search_start_cb(callback: types.CallbackQuery, state: FSMContext):
    ws_id = int(callback.data.split(":", 1)[1])
    await state.set_state(HistoryFilterStates.searching_item)
    await state.update_data(hist_ws_id=ws_id)
    await callback.message.edit_text("🔍 Введите часть названия позиции для поиска в истории:")
    await callback.answer()


@dp.message(HistoryFilterStates.searching_item)
async def hist_search_process(message: types.Message, state: FSMContext):
    query = (message.text or "").strip().lower()
    if not query:
        await message.answer("Введите часть названия позиции:", reply_markup=get_cancel_keyboard())
        return

    data = await state.get_data()
    ws_id = int(data["hist_ws_id"])
    catalog = db_get_catalog(ws_id)
    matches = [x for x in catalog if query in x["name"].lower()]
    if not matches:
        await message.answer(
            f"По запросу **«{message.text.strip()}»** позиций не найдено. Введите другой текст или нажмите «❌ Отмена»:",
            reply_markup=get_cancel_keyboard(),
            parse_mode="Markdown",
        )
        return

    await state.clear()
    buttons = [[InlineKeyboardButton(text=x["name"], callback_data=f"hitem:{ws_id}:{x['name']}")] for x in matches[:25]]
    buttons.append([InlineKeyboardButton(text="❌ Закрыть", callback_data="cancel")])
    await message.answer(
        f"Найдено позиций: **{len(matches)}**. Выберите позицию:",
        reply_markup=InlineKeyboardMarkup(inline_keyboard=buttons),
        parse_mode="Markdown",
    )


@dp.callback_query(F.data.startswith("hitem:"))
async def hist_by_item_cb(callback: types.CallbackQuery):
    _, ws_id_str, item_name = callback.data.split(":", 2)
    ws_id = int(ws_id_str)
    ws = db_get_workspace_info(ws_id)
    if not ws:
        await callback.message.edit_text("Склад не найден.")
        await callback.answer()
        return

    is_admin = callback.from_user.id in ADMIN_IDS
    rows = db_get_movements_by_item(ws_id, item_name, limit=25)
    stock_list = db_get_stock(ws_id)
    item_stock = next((x for x in stock_list if x["name"].lower() == item_name.lower()), None)
    cur_bal = item_stock["balance"] if item_stock else 0.0
    unit = item_stock["unit"] if item_stock else (rows[0]["unit"] if rows else "")
    unit_str = f" {unit}" if unit else ""

    if not rows:
        await callback.message.edit_text(
            f"📦 **{item_name}** («{ws['title']}»)\n"
            f"Текущий остаток: **{cur_bal:g}**{unit_str}\n\n"
            f"Операций по этой позиции пока не было.",
            parse_mode="Markdown",
        )
        await send_with_fresh_keyboard(
            callback.message,
            "Выберите следующее действие:",
            has_workspace=True,
            is_admin=is_admin,
            in_menu=True,
            parse_mode=None,
        )
        await callback.answer()
        return

    lines = [
        f"📦 **История по позиции «{item_name}» («{ws['title']}»)**",
        f"Текущий остаток: **{cur_bal:g}**{unit_str}\n",
    ]
    lines.extend(format_movements_lines(rows, callback.message.chat, is_admin))
    await callback.message.edit_text("\n".join(lines), parse_mode="Markdown")
    await send_with_fresh_keyboard(
        callback.message,
        "Выберите следующее действие:",
        has_workspace=True,
        is_admin=is_admin,
        in_menu=True,
        parse_mode=None,
    )
    await callback.answer()


@dp.callback_query(F.data.startswith("hist_back:"))
async def hist_back_cb(callback: types.CallbackQuery):
    ws_id = int(callback.data.split(":", 1)[1])
    ws = db_get_workspace_info(ws_id)
    if not ws:
        await callback.message.edit_text("Склад не найден.")
        await callback.answer()
        return

    markup = InlineKeyboardMarkup(
        inline_keyboard=[
            [
                InlineKeyboardButton(text="🕒 Последние 10", callback_data=f"hist_last:{ws_id}"),
                InlineKeyboardButton(text="📅 По дате", callback_data=f"hist_date_menu:{ws_id}"),
            ],
            [InlineKeyboardButton(text="📦 По позиции", callback_data=f"hist_item_menu:{ws_id}")],
            [InlineKeyboardButton(text="❌ Закрыть", callback_data="cancel")],
        ]
    )
    await callback.message.edit_text(
        f"📜 **История операций («{ws['title']}»)**\nВыберите способ просмотра:",
        reply_markup=markup,
        parse_mode="Markdown",
    )
    await callback.answer()


@dp.message(F.text == "📊 Текущий остаток")
async def show_stock(message: types.Message):
    ws = await ensure_workspace(message)
    if not ws:
        return
    stock_list = db_get_stock(ws["id"])
    if not stock_list:
        await send_with_fresh_keyboard(
            message,
            "В справочнике пока нет позиций. Нажмите «📝 Добавить позицию» или «📤 Загрузить остатки из Excel».",
            has_workspace=True,
            parse_mode=None,
        )
        return

    lines = [f"📦 **Текущие остатки («{ws['title']}»):**\n"]
    for item in stock_list:
        unit_str = f" {item['unit']}" if item["unit"] else ""
        if item["balance"] < 0:
            mark = "🔴 "
        elif item["balance"] == 0:
            mark = "🟡 "
        else:
            mark = "• "
        lines.append(f"{mark}**{item['name']}**: {item['balance']:g}{unit_str}")

    await send_with_fresh_keyboard(
        message,
        "\n".join(lines),
        has_workspace=True,
    )


@dp.message(F.text == "📥 Остатки в Excel")
async def export_stock_excel(message: types.Message):
    ws = await ensure_workspace(message)
    if not ws:
        return

    stock_list = db_get_stock(ws["id"])
    if not stock_list:
        await send_with_fresh_keyboard(
            message,
            "В справочнике пока нет позиций для выгрузки.",
            has_workspace=True,
            parse_mode=None,
        )
        return

    safe_title = "".join(c if c.isalnum() or c in (" ", "_", "-") else "_" for c in ws["title"]).strip() or "sklad"
    stamp = datetime.now().strftime("%Y%m%d_%H%M")

    if HAS_OPENPYXL:
        movements = db_get_all_movements_for_export(ws["id"])
        xlsx_bytes = build_stock_xlsx_bytes(ws["title"], stock_list, movements)
        doc = BufferedInputFile(xlsx_bytes, filename=f"ostatki_{safe_title}_{stamp}.xlsx")
        await message.answer_document(
            doc,
            caption=f"📥 Остатки и операции по складу «{ws['title']}»",
            reply_markup=get_main_keyboard(has_workspace=True),
        )
    else:
        csv_bytes = build_stock_csv_bytes(ws["title"], stock_list)
        doc = BufferedInputFile(csv_bytes, filename=f"ostatki_{safe_title}_{stamp}.csv")
        await message.answer_document(
            doc,
            caption=(
                f"📥 Остатки по складу «{ws['title']}» (CSV).\n"
                "Для формата `.xlsx` установите библиотеку: `pip install openpyxl`"
            ),
            reply_markup=get_main_keyboard(has_workspace=True),
        )


@dp.startup()
async def on_startup():
    logging.info(f"Бот запущен. Версия сборки: {BOT_VERSION}")
    await notify_admins(f"🚀 **Бот запущен**\nВерсия сборки: `{BOT_VERSION}`")


async def main():
    init_db()
    await dp.start_polling(bot)


if __name__ == "__main__":
    asyncio.run(main())
