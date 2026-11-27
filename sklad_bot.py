import asyncio
import csv
import io
import logging
import sqlite3
from datetime import datetime
from typing import Optional

from aiogram import Bot, Dispatcher, F, types
from aiogram.filters import Command
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

    HAS_OPENPYXL = True
except ImportError:
    HAS_OPENPYXL = False

# --- НАСТРОЙКИ ---
BOT_TOKEN = "8763607619:AAGcLqQjNSkjCPYmFcjE_HhC458D_HrUyig"
DB_FILE = "sklad_data.db"
ADMIN_IDS: set[int] = {363447845}
LOW_STOCK_THRESHOLD = 0.0

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


def norm_username(raw: str) -> str:
    return raw.strip().lstrip("@").strip().lower()


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


def db_invite_by_username(workspace_id: int, raw_username: str, inviter_tg_id: int) -> tuple[str, str, Optional[int]]:
    uname_norm = norm_username(raw_username)
    if not uname_norm:
        return "invalid", "", None

    conn = get_conn()
    cur = conn.cursor()
    ws = cur.execute("SELECT id, title FROM workspaces WHERE id = ?", (workspace_id,)).fetchone()
    if not ws:
        conn.close()
        return "no_ws", uname_norm, None

    user_row = cur.execute(
        "SELECT tg_id FROM users WHERE lower(trim(username)) = ?",
        (uname_norm,),
    ).fetchone()

    if user_row:
        target_tg_id = int(user_row["tg_id"])
        cur.execute(
            """
            INSERT OR IGNORE INTO workspace_members (workspace_id, tg_id, role)
            VALUES (?, ?, 'member')
            """,
            (workspace_id, target_tg_id),
        )
        cur.execute(
            """
            UPDATE users
            SET active_workspace_id = COALESCE(active_workspace_id, ?)
            WHERE tg_id = ?
            """,
            (workspace_id, target_tg_id),
        )
        conn.commit()
        conn.close()
        return "linked", uname_norm, target_tg_id

    cur.execute(
        """
        INSERT INTO pending_invites (workspace_id, username_norm, invited_by_tg_id, created_at)
        VALUES (?, ?, ?, ?)
        ON CONFLICT(workspace_id, username_norm) DO UPDATE SET
            invited_by_tg_id = excluded.invited_by_tg_id,
            created_at = excluded.created_at
        """,
        (workspace_id, uname_norm, inviter_tg_id, datetime.now().strftime("%Y-%m-%d %H:%M:%S")),
    )
    conn.commit()
    conn.close()
    return "pending", uname_norm, None


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


def db_delete_movement_by_id(
    workspace_id: int, movement_id: int, caller_tg_id: int, is_admin: bool
) -> tuple[str, Optional[str]]:
    conn = get_conn()
    cur = conn.cursor()
    row = cur.execute(
        """
        SELECT id, op_date, op_type, item_name, quantity, unit, actor_tg_id
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


def format_actor_for_admin(user: types.User, chat: types.Chat) -> str:
    name = user.full_name or (f"@{user.username}" if user.username else "Пользователь")
    uname = f" (@{user.username})" if user.username else ""
    chat_info = f" в группе «{chat.title}»" if chat.type in ("group", "supergroup") else ""
    return f"{name}{uname} [ID: `{user.id}`]{chat_info}"


def format_actor_public(actor_name: str, chat: types.Chat) -> str:
    clean_name = actor_name or "Пользователь"
    chat_info = f" в группе «{chat.title}»" if chat.type in ("group", "supergroup") else ""
    return f"{clean_name}{chat_info}"


def parse_number(raw: str) -> Optional[float]:
    try:
        return float(raw.replace(",", ".").replace(" ", "").strip())
    except ValueError:
        return None


def build_stock_xlsx_bytes(ws_title: str, stock_list: list[dict]) -> bytes:
    wb = openpyxl.Workbook()
    sheet = wb.active
    sheet.title = "Остатки"

    sheet.append(["Склад", ws_title])
    sheet.append(["Дата выгрузки", datetime.now().strftime("%d.%m.%Y %H:%M")])
    sheet.append([])
    sheet.append(["Категория", "Наименование", "Остаток", "Ед. измерения"])

    for item in stock_list:
        sheet.append([item["category"], item["name"], item["balance"], item["unit"]])

    sheet.column_dimensions["A"].width = 22
    sheet.column_dimensions["B"].width = 34
    sheet.column_dimensions["C"].width = 14
    sheet.column_dimensions["D"].width = 16

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


class OperationStates(StatesGroup):
    choosing_date = State()
    entering_custom_date = State()
    choosing_category = State()
    choosing_item = State()
    entering_quantity = State()
    entering_comment = State()
    confirming = State()


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


class LinkUserStates(StatesGroup):
    entering_username = State()


def get_main_keyboard(has_workspace: bool) -> ReplyKeyboardMarkup:
    if not has_workspace:
        rows = [
            [KeyboardButton(text="🆕 Создать склад")],
            [KeyboardButton(text="📋 Меню")],
        ]
    else:
        rows = [
            [KeyboardButton(text="➕ Приход"), KeyboardButton(text="➖ Расход")],
            [KeyboardButton(text="📊 Текущий остаток"), KeyboardButton(text="📝 Добавить позицию")],
            [KeyboardButton(text="📋 Меню")],
        ]
    return ReplyKeyboardMarkup(keyboard=rows, resize_keyboard=True)


def get_menu_list_keyboard(has_workspace: bool, is_admin: bool) -> ReplyKeyboardMarkup:
    rows: list[list[KeyboardButton]] = []
    if has_workspace:
        rows.append([KeyboardButton(text="📜 Последние операции")])
        rows.append([KeyboardButton(text="🗑 Отменить операцию")])
        rows.append([KeyboardButton(text="✏️ Редактировать позицию")])
        rows.append([KeyboardButton(text="✏️ Редактировать склад")])
        rows.append([KeyboardButton(text="🔄 Сменить склад")])
        rows.append([KeyboardButton(text="📥 Остатки в Excel")])
        rows.append([KeyboardButton(text="👤 Участники склада")])
        rows.append([KeyboardButton(text="👥 Привязать сотрудника")])
        rows.append([KeyboardButton(text="❌ Отвязать сотрудника")])
        rows.append([KeyboardButton(text="ℹ️ Мой склад")])
        rows.append([KeyboardButton(text="🆕 Создать склад")])
        rows.append([KeyboardButton(text="🗑 Удалить склад")])
    else:
        rows.append([KeyboardButton(text="🆕 Создать склад")])

    if is_admin:
        rows.append([KeyboardButton(text="👑 Все склады (Админ)")])

    rows.append([KeyboardButton(text="⬅️ Назад")])
    return ReplyKeyboardMarkup(keyboard=rows, resize_keyboard=True)


def get_cancel_keyboard() -> ReplyKeyboardMarkup:
    return ReplyKeyboardMarkup(
        keyboard=[[KeyboardButton(text="❌ Отмена")]],
        resize_keyboard=True,
    )


async def ensure_workspace(message: types.Message) -> Optional[dict]:
    user = message.from_user
    await db_upsert_user(user.id, user.username, user.full_name or "")
    ws = db_get_active_workspace(user.id)
    if not ws:
        await message.answer(
            "У вас пока нет подключённого склада.\nНажмите **«🆕 Создать склад»**, чтобы создать свой склад.",
            reply_markup=get_main_keyboard(has_workspace=False),
            parse_mode="Markdown",
        )
        return None
    return ws


@dp.message(Command("start"))
async def cmd_start(message: types.Message, state: FSMContext):
    await state.clear()
    user = message.from_user
    await db_upsert_user(user.id, user.username, user.full_name or "")
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

    await message.answer(
        text,
        reply_markup=get_main_keyboard(has_workspace=bool(ws)),
        parse_mode="Markdown",
    )


@dp.message(F.text == "📋 Меню")
async def open_menu_buttons(message: types.Message, state: FSMContext):
    await state.clear()
    user = message.from_user
    await db_upsert_user(user.id, user.username, user.full_name or "")
    ws = db_get_active_workspace(user.id)
    is_admin = user.id in ADMIN_IDS

    await message.answer(
        "Выберите действие:",
        reply_markup=get_menu_list_keyboard(has_workspace=bool(ws), is_admin=is_admin),
    )


@dp.message(F.text.in_(["⬅️ Назад", "❌ Отмена"]))
async def back_to_main_keyboard(message: types.Message, state: FSMContext):
    await state.clear()
    user = message.from_user
    await db_upsert_user(user.id, user.username, user.full_name or "")
    ws = db_get_active_workspace(user.id)
    await message.answer(
        "Главное меню:",
        reply_markup=get_main_keyboard(has_workspace=bool(ws)),
    )


@dp.callback_query(F.data == "cancel")
async def cancel_operation_cb(callback: types.CallbackQuery, state: FSMContext):
    await state.clear()
    user = callback.from_user
    ws = db_get_active_workspace(user.id)
    await callback.message.edit_text("Действие отменено.")
    await callback.message.answer(
        "Главное меню:",
        reply_markup=get_main_keyboard(has_workspace=bool(ws)),
    )
    await callback.answer()


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
    ws_id = db_create_and_bind_workspace(user.id, title)
    await state.clear()

    await message.answer(
        f"✅ Склад **«{title}»** создан и выбран активным!",
        reply_markup=get_main_keyboard(has_workspace=True),
        parse_mode="Markdown",
    )

    await notify_admins(
        f"🆕 **Создан новый склад**\n"
        f"• Пользователь: {format_actor_for_admin(user, message.chat)}\n"
        f"• Склад (#{ws_id}): *{title}*"
    )


@dp.message(F.text == "✏️ Редактировать склад")
async def msg_edit_workspace(message: types.Message, state: FSMContext):
    ws = await ensure_workspace(message)
    if not ws:
        return
    user = message.from_user
    is_admin = user.id in ADMIN_IDS
    if not is_admin and user.id != ws["owner_tg_id"]:
        await message.answer(
            "Редактировать название склада может только его владелец или администратор.",
            reply_markup=get_main_keyboard(has_workspace=True),
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
        await message.answer("Склад не найден.", reply_markup=get_main_keyboard(has_workspace=True))
        return
    if status == "forbidden":
        await message.answer("Недостаточно прав для редактирования склада.", reply_markup=get_main_keyboard(has_workspace=True))
        return

    await message.answer(
        f"✅ Название склада изменено: **«{old_title}»** → **«{new_title}»**.",
        reply_markup=get_main_keyboard(has_workspace=True),
        parse_mode="Markdown",
    )
    await notify_admins(
        f"✏️ **Переименован склад (#{ws_id})**\n"
        f"• Кто: {format_actor_for_admin(user, message.chat)}\n"
        f"• Было: *{old_title}* → Стало: *{new_title}*"
    )


@dp.message(F.text == "🗑 Удалить склад")
async def request_delete_workspace(message: types.Message):
    ws = await ensure_workspace(message)
    if not ws:
        return
    user = message.from_user
    is_admin = user.id in ADMIN_IDS
    if not is_admin and user.id != ws["owner_tg_id"]:
        await message.answer(
            "Удалить склад может только его владелец или администратор.",
            reply_markup=get_main_keyboard(has_workspace=True),
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

    await notify_admins(
        f"🗑 **Удалён склад**\n"
        f"• Кто удалил: {format_actor_for_admin(user, callback.message.chat)}\n"
        f"• Склад (#{ws_id}): *{deleted_title}*"
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
        await message.answer(
            "У вас пока нет доступных складов. Нажмите **«🆕 Создать склад»**.",
            reply_markup=get_main_keyboard(has_workspace=False),
            parse_mode="Markdown",
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
        await callback.message.answer(
            "Главное меню:",
            reply_markup=get_main_keyboard(has_workspace=True),
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
        "• Подключить сотрудника: нажмите **«📋 Меню»** → **«👥 Привязать сотрудника»** и укажите его `@никнейм`.",
    ]
    if user.id in ADMIN_IDS:
        lines.append(f"• *(Админ)* ID склада: `#{ws['id']}`, ID владельца: `{ws['owner_tg_id']}`")

    await message.answer(
        "\n".join(lines),
        reply_markup=get_main_keyboard(has_workspace=True),
        parse_mode="Markdown",
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
        lines.append("\n⏳ **Ожидают первого входа в бот:**")
        for p in pending:
            lines.append(f"• @{p['username_norm']}")

    await message.answer(
        "\n".join(lines),
        reply_markup=get_main_keyboard(has_workspace=True),
        parse_mode="Markdown",
    )


@dp.message(F.text == "👥 Привязать сотрудника")
async def start_link_user_ui(message: types.Message, state: FSMContext):
    ws = await ensure_workspace(message)
    if not ws:
        return
    user = message.from_user
    if user.id not in ADMIN_IDS and user.id != ws["owner_tg_id"]:
        await message.answer(
            "Привязывать сотрудников может только владелец склада или администратор.",
            reply_markup=get_main_keyboard(has_workspace=True),
        )
        return

    await state.clear()
    await state.update_data(target_ws_id=ws["id"], target_ws_title=ws["title"])
    await state.set_state(LinkUserStates.entering_username)

    await message.answer(
        f"Привязка сотрудника к складу **«{ws['title']}»**.\n"
        f"Введите **никнейм** сотрудника в формате `@username`:",
        reply_markup=get_cancel_keyboard(),
        parse_mode="Markdown",
    )


@dp.message(LinkUserStates.entering_username)
async def process_link_username(message: types.Message, state: FSMContext):
    raw = (message.text or "").strip()
    uname_norm = norm_username(raw)
    if not uname_norm or len(uname_norm) < 2:
        await message.answer("Укажите корректный никнейм, например `@ivanov`:", reply_markup=get_cancel_keyboard())
        return

    data = await state.get_data()
    ws_id = int(data["target_ws_id"])
    ws_title = data["target_ws_title"]
    user = message.from_user

    status, _, target_tg_id = db_invite_by_username(ws_id, uname_norm, user.id)
    await state.clear()

    if status == "linked":
        await message.answer(
            f"✅ Пользователь **@{uname_norm}** подключён к складу **«{ws_title}»**.",
            reply_markup=get_main_keyboard(has_workspace=True),
            parse_mode="Markdown",
        )
        if target_tg_id:
            try:
                await bot.send_message(
                    target_tg_id,
                    f"✅ Вас подключили к складу **«{ws_title}»**.",
                    parse_mode="Markdown",
                )
            except Exception as e:
                logging.warning(f"Не удалось отправить уведомление пользователю {target_tg_id}: {e}")
    else:
        await message.answer(
            f"⏳ Приглашение для **@{uname_norm}** сохранено.\n"
            f"Он автоматически подключится к складу **«{ws_title}»**, когда нажмёт `/start` в боте.",
            reply_markup=get_main_keyboard(has_workspace=True),
            parse_mode="Markdown",
        )


@dp.message(F.text == "❌ Отвязать сотрудника")
async def start_unlink_user_ui(message: types.Message):
    ws = await ensure_workspace(message)
    if not ws:
        return
    user = message.from_user
    is_admin = user.id in ADMIN_IDS
    if not is_admin and user.id != ws["owner_tg_id"]:
        await message.answer(
            "Отвязывать сотрудников может только владелец склада или администратор.",
            reply_markup=get_main_keyboard(has_workspace=True),
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
        await message.answer(
            f"На складе **«{ws['title']}»** нет сотрудников для отвязки.",
            reply_markup=get_main_keyboard(has_workspace=True),
            parse_mode="Markdown",
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
    await callback.message.answer(
        "Главное меню:",
        reply_markup=get_main_keyboard(has_workspace=True),
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
    await callback.message.answer(
        "Главное меню:",
        reply_markup=get_main_keyboard(has_workspace=True),
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

    lines = ["👑 **Все склады (Админ):**\n"]
    buttons = []
    for r in rows[:20]:
        lines.append(
            f"• `#{r['id']}` **{r['title']}** · владелец `{r['owner_tg_id']}` · "
            f"участников: `{r['users_count']}` · операций: `{r['ops_count']}`"
        )
        buttons.append(
            [InlineKeyboardButton(text=f"Переключиться на #{r['id']} ({r['title']})", callback_data=f"switch_ws:{r['id']}")]
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
        title = ws["title"] if ws else str(ws_id)
        await callback.message.edit_text(
            f"✅ Вы переключились на склад **«{title}»** (`#{ws_id}`).",
            parse_mode="Markdown",
        )
        await callback.message.answer(
            "Главное меню:",
            reply_markup=get_main_keyboard(has_workspace=True),
        )
    else:
        await callback.message.edit_text("Склад не найден.")
    await callback.answer()
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
    ws_title = data.get("workspace_title", "")
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
    user = message_or_call.from_user
    chat = message_or_call.message.chat if isinstance(message_or_call, types.CallbackQuery) else message_or_call.chat
    try:
        db_add_catalog_item(ws_id, name, unit, cat_clean)
        result_text = (
            f"✅ **Новая позиция добавлена в справочник!**\n"
            f"• Наименование: {name}\n• Ед. измерения: {unit}\n• Категория: {cat_clean}"
        )
        await notify_admins(
            f"📝 **Новая позиция в справочнике**\n"
            f"• Кто: {format_actor_for_admin(user, chat)}\n"
            f"• Склад (#{ws_id}): *{ws_title}*\n"
            f"• Позиция: *{name}* ({unit}, {cat_clean})"
        )
    except Exception as e:
        result_text = f"❌ Ошибка при сохранении в справочник: {e}"
    await state.clear()
    if isinstance(message_or_call, types.CallbackQuery):
        await message_or_call.message.edit_text(result_text, parse_mode="Markdown")
        await message_or_call.message.answer("Главное меню:", reply_markup=get_main_keyboard(has_workspace=True))
        await message_or_call.answer()
    else:
        await message_or_call.answer(
            result_text,
            reply_markup=get_main_keyboard(has_workspace=True),
            parse_mode="Markdown",
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
        await message.answer(
            "В справочнике пока нет позиций для редактирования.",
            reply_markup=get_main_keyboard(has_workspace=True),
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

    user = message_or_call.from_user
    chat = message_or_call.message.chat if isinstance(message_or_call, types.CallbackQuery) else message_or_call.chat

    status, old_data = db_update_catalog_item(ws_id, item_id, new_name, new_unit, cat_clean)
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
        await notify_admins(
            f"✏️ **Отредактирована позиция (склад #{ws_id})**\n"
            f"• Кто: {format_actor_for_admin(user, chat)}\n"
            f"• Было: *{old_data['name']}* ({old_data['unit']}, {old_data['category']})\n"
            f"• Стало: *{new_name}* ({new_unit}, {cat_clean})"
        )

    if isinstance(message_or_call, types.CallbackQuery):
        await message_or_call.message.edit_text(res, parse_mode="Markdown")
        await message_or_call.message.answer("Главное меню:", reply_markup=get_main_keyboard(has_workspace=True))
        await message_or_call.answer()
    else:
        await message_or_call.answer(res, reply_markup=get_main_keyboard(has_workspace=True), parse_mode="Markdown")


@dp.callback_query(EditCatalogStates.entering_category, F.data.startswith("ecat:"))
async def process_edit_category_cb(callback: types.CallbackQuery, state: FSMContext):
    await finish_edit_catalog_item(callback, state, callback.data.split(":", 1)[1])


@dp.message(EditCatalogStates.entering_category)
async def process_edit_category_text(message: types.Message, state: FSMContext):
    await finish_edit_catalog_item(message, state, message.text or "")


# --- ПРИХОД И РАСХОД ---
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
        await target_message.answer("В справочнике пока нет товаров. Нажмите «📝 Добавить позицию».")
        return
    await state.update_data(catalog=catalog)
    categories = []
    for c in catalog:
        if c["category"] not in categories:
            categories.append(c["category"])
    if len(categories) > 1:
        buttons = [[InlineKeyboardButton(text=cat, callback_data=f"opcat:{cat}")] for cat in categories]
        buttons.append([InlineKeyboardButton(text="📋 Все позиции списком", callback_data="opcat:__ALL__")])
        buttons.append([InlineKeyboardButton(text="❌ Отмена", callback_data="cancel")])
        await state.set_state(OperationStates.choosing_category)
        text = f"Операция: **{op_type}** · Дата: **{op_date}**\nВыберите категорию:"
    else:
        buttons = [[InlineKeyboardButton(text=item["name"], callback_data=f"item:{item['name']}")] for item in catalog]
        buttons.append([InlineKeyboardButton(text="❌ Отмена", callback_data="cancel")])
        await state.set_state(OperationStates.choosing_item)
        text = f"Операция: **{op_type}** · Дата: **{op_date}**\nВыберите позицию:"
    markup = InlineKeyboardMarkup(inline_keyboard=buttons)
    if edit:
        await target_message.edit_text(text, reply_markup=markup, parse_mode="Markdown")
    else:
        await target_message.answer(text, reply_markup=markup, parse_mode="Markdown")


@dp.callback_query(OperationStates.choosing_date, F.data.startswith("date:"))
async def process_date_choice(callback: types.CallbackQuery, state: FSMContext):
    val = callback.data.split(":", 1)[1]
    if val == "custom":
        await state.set_state(OperationStates.entering_custom_date)
        await callback.message.edit_text("Введите дату в формате `ДД.ММ.ГГГГ` (например: `27.11.2026`):", parse_mode="Markdown")
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
        await message.answer("Неверный формат даты. Введите как `ДД.ММ.ГГГГ` (например, `27.11.2026`):", parse_mode="Markdown")
        return
    await show_items_or_categories(message, state, date_str, edit=False)


@dp.callback_query(OperationStates.choosing_category, F.data.startswith("opcat:"))
async def process_op_category(callback: types.CallbackQuery, state: FSMContext):
    cat = callback.data.split(":", 1)[1]
    data = await state.get_data()
    catalog = data.get("catalog", [])
    items = catalog if cat == "__ALL__" else [x for x in catalog if x["category"] == cat]
    buttons = [[InlineKeyboardButton(text=x["name"], callback_data=f"item:{x['name']}")] for x in items]
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
    ws_title = data.get("workspace_title", "")
    op_type = data["op_type"]
    item_name = data["item_name"]
    qty = float(data["quantity"])
    op_date = data["op_date"]
    unit = data.get("item_unit", "")
    category = data.get("item_category", "")
    comment = data.get("comment", "")

    user = callback.from_user
    chat = callback.message.chat
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
        sign = "➕" if op_type == "Приход" else "➖"
        await notify_admins(
            f"{sign} **{op_type}** ({op_date})\n"
            f"• Кто: {format_actor_for_admin(user, chat)}\n"
            f"• Склад (#{ws_id}): *{ws_title}*\n"
            f"• Позиция: *{item_name}* — *{qty:g} {unit}*\n"
            f"• Комментарий: {comment if comment else '—'}{low_stock_alert}"
        )
    except Exception as e:
        summary = f"❌ Ошибка при записи: {e}"

    await state.clear()
    await callback.message.edit_text(summary, parse_mode="Markdown")
    await callback.message.answer("Главное меню:", reply_markup=get_main_keyboard(has_workspace=True))
    await callback.answer()


# --- ОТМЕНА ОПЕРАЦИИ, ИСТОРИЯ, ОСТАТКИ И EXCEL ---
@dp.message(F.text.in_(["🗑 Отменить операцию", "↩️ Отменить последнюю"]))
async def choose_operation_to_cancel(message: types.Message):
    ws = await ensure_workspace(message)
    if not ws:
        return

    rows = db_get_recent_movements(ws["id"], limit=10)
    if not rows:
        await message.answer(
            "На этом складе пока нет операций для отмены.",
            reply_markup=get_main_keyboard(has_workspace=True),
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

    await callback.message.edit_text(f"🗑 Операция отменена: **{info}**", parse_mode="Markdown")
    await callback.message.answer("Главное меню:", reply_markup=get_main_keyboard(has_workspace=True))
    await callback.answer()

    await notify_admins(
        f"🗑 **Отменена операция**\n"
        f"• Кто отменил: {format_actor_for_admin(user, callback.message.chat)}\n"
        f"• Склад (#{ws_id})\n"
        f"• Запись: `{info}`"
    )


@dp.message(F.text == "📜 Последние операции")
async def recent_movements(message: types.Message):
    ws = await ensure_workspace(message)
    if not ws:
        return
    rows = db_get_recent_movements(ws["id"], limit=10)
    if not rows:
        await message.answer("Операций пока не было.", reply_markup=get_main_keyboard(has_workspace=True))
        return

    is_admin = message.from_user.id in ADMIN_IDS
    lines = [f"📜 **Последние операции («{ws['title']}»):**\n"]
    for r in rows:
        sign = "➕" if r["op_type"] == "Приход" else "➖"
        extra = []
        if r["actor_name"]:
            actor_text = (
                f"{r['actor_name']} [ID: `{r['actor_tg_id']}`]"
                if is_admin and r["actor_tg_id"]
                else format_actor_public(r["actor_name"], message.chat)
            )
            extra.append(actor_text)
        if r["comment"]:
            extra.append(f"«{r['comment']}»")
        extra_str = f" — {', '.join(extra)}" if extra else ""
        unit_str = f" {r['unit']}" if r["unit"] else ""
        lines.append(f"{sign} `{r['op_date']}` **{r['item_name']}**: {float(r['quantity']):g}{unit_str}{extra_str}")

    await message.answer(
        "\n".join(lines),
        reply_markup=get_main_keyboard(has_workspace=True),
        parse_mode="Markdown",
    )


@dp.message(F.text == "📊 Текущий остаток")
async def show_stock(message: types.Message):
    ws = await ensure_workspace(message)
    if not ws:
        return
    stock_list = db_get_stock(ws["id"])
    if not stock_list:
        await message.answer(
            "В справочнике пока нет позиций. Нажмите «📝 Добавить позицию».",
            reply_markup=get_main_keyboard(has_workspace=True),
        )
        return

    lines = [f"📦 **Текущие остатки («{ws['title']}»):**\n"]
    for item in stock_list:
        unit_str = f" {item['unit']}" if item["unit"] else ""
        lines.append(f"• **{item['name']}**: {item['balance']:g}{unit_str}")

    await message.answer(
        "\n".join(lines),
        reply_markup=get_main_keyboard(has_workspace=True),
        parse_mode="Markdown",
    )


@dp.message(F.text == "📥 Остатки в Excel")
async def export_stock_excel(message: types.Message):
    ws = await ensure_workspace(message)
    if not ws:
        return

    stock_list = db_get_stock(ws["id"])
    if not stock_list:
        await message.answer(
            "В справочнике пока нет позиций для выгрузки.",
            reply_markup=get_main_keyboard(has_workspace=True),
        )
        return

    safe_title = "".join(c if c.isalnum() or c in (" ", "_", "-") else "_" for c in ws["title"]).strip() or "sklad"
    stamp = datetime.now().strftime("%Y%m%d_%H%M")

    if HAS_OPENPYXL:
        xlsx_bytes = build_stock_xlsx_bytes(ws["title"], stock_list)
        doc = BufferedInputFile(xlsx_bytes, filename=f"ostatki_{safe_title}_{stamp}.xlsx")
        await message.answer_document(
            doc,
            caption=f"📥 Остатки по складу «{ws['title']}»",
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


async def main():
    init_db()
    await dp.start_polling(bot)


if __name__ == "__main__":
    asyncio.run(main())
