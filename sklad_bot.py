import asyncio
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
    InlineKeyboardButton,
    InlineKeyboardMarkup,
    KeyboardButton,
    ReplyKeyboardMarkup,
)

# --- НАСТРОЙКИ ---
BOT_TOKEN = "8763607619:AAGcLqQjNSkjCPYmFcjE_HhC458D_HrUyig"
DB_FILE = "sklad_data.db"
ADMIN_IDS: set[int] = {363447845}
LOW_STOCK_THRESHOLD = 10.0

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
            username TEXT,
            full_name TEXT,
            workspace_id INTEGER,
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
    conn.commit()
    conn.close()


def db_upsert_user(tg_id: int, username: Optional[str], full_name: str):
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
        (tg_id, username or "", full_name),
    )
    conn.commit()
    conn.close()


def db_get_user_workspace(tg_id: int) -> Optional[dict]:
    conn = get_conn()
    cur = conn.cursor()
    cur.execute(
        """
        SELECT w.id, w.title, w.owner_tg_id
        FROM users u
        JOIN workspaces w ON u.workspace_id = w.id
        WHERE u.tg_id = ?
        """,
        (tg_id,),
    )
    row = cur.fetchone()
    conn.close()
    if not row:
        return None
    return {"id": row["id"], "title": row["title"], "owner_tg_id": row["owner_tg_id"]}


def db_create_and_bind_workspace(tg_id: int, title: str) -> int:
    conn = get_conn()
    cur = conn.cursor()
    now_str = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    cur.execute(
        "INSERT INTO workspaces (title, owner_tg_id, created_at) VALUES (?, ?, ?)",
        (title, tg_id, now_str),
    )
    ws_id = cur.lastrowid
    cur.execute("UPDATE users SET workspace_id = ? WHERE tg_id = ?", (ws_id, tg_id))
    conn.commit()
    conn.close()
    return ws_id


def db_bind_user_to_workspace(tg_id: int, workspace_id: int) -> bool:
    conn = get_conn()
    cur = conn.cursor()
    cur.execute("SELECT id FROM workspaces WHERE id = ?", (workspace_id,))
    if not cur.fetchone():
        conn.close()
        return False
    cur.execute(
        """
        INSERT INTO users (tg_id, username, full_name, workspace_id)
        VALUES (?, '', ?, ?)
        ON CONFLICT(tg_id) DO UPDATE SET workspace_id = excluded.workspace_id
        """,
        (tg_id, str(tg_id), workspace_id),
    )
    conn.commit()
    conn.close()
    return True


def db_list_workspaces() -> list[sqlite3.Row]:
    conn = get_conn()
    cur = conn.cursor()
    cur.execute(
        """
        SELECT w.id, w.title, w.owner_tg_id,
               (SELECT COUNT(*) FROM users u WHERE u.workspace_id = w.id) AS users_count,
               (SELECT COUNT(*) FROM movements m WHERE m.workspace_id = w.id) AS ops_count
        FROM workspaces w
        ORDER BY w.id DESC
        """
    )
    rows = cur.fetchall()
    conn.close()
    return rows


def db_get_catalog(workspace_id: int) -> list[dict]:
    conn = get_conn()
    cur = conn.cursor()
    cur.execute(
        "SELECT name, unit, category FROM catalog_items WHERE workspace_id = ? ORDER BY id ASC",
        (workspace_id,),
    )
    rows = cur.fetchall()
    conn.close()
    return [{"name": r["name"], "unit": r["unit"], "category": r["category"]} for r in rows]


def db_catalog_name_exists(workspace_id: int, name: str) -> bool:
    conn = get_conn()
    cur = conn.cursor()
    cur.execute(
        "SELECT 1 FROM catalog_items WHERE workspace_id = ? AND lower(trim(name)) = lower(trim(?))",
        (workspace_id, name),
    )
    exists = cur.fetchone() is not None
    conn.close()
    return exists


def db_add_catalog_item(workspace_id: int, name: str, unit: str, category: str):
    conn = get_conn()
    cur = conn.cursor()
    cur.execute(
        "INSERT INTO catalog_items (workspace_id, name, unit, category) VALUES (?, ?, ?, ?)",
        (workspace_id, name, unit, category),
    )
    conn.commit()
    conn.close()


def db_get_stock(workspace_id: int) -> list[dict]:
    conn = get_conn()
    cur = conn.cursor()
    cur.execute(
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
    )
    rows = cur.fetchall()
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
    cur.execute(
        """
        SELECT op_date, op_type, item_name, quantity, unit, comment, actor_name
        FROM movements
        WHERE workspace_id = ?
        ORDER BY id DESC
        LIMIT ?
        """,
        (workspace_id, limit),
    )
    rows = cur.fetchall()
    conn.close()
    return rows


def db_undo_last_movement(workspace_id: int, caller_tg_id: int, is_admin: bool) -> tuple[str, Optional[str]]:
    conn = get_conn()
    cur = conn.cursor()
    cur.execute(
        """
        SELECT id, op_date, op_type, item_name, quantity, unit, actor_tg_id
        FROM movements
        WHERE workspace_id = ?
        ORDER BY id DESC
        LIMIT 1
        """,
        (workspace_id,),
    )
    last_row = cur.fetchone()
    if not last_row:
        conn.close()
        return "empty", None
    if not is_admin and int(last_row["actor_tg_id"]) != caller_tg_id:
        conn.close()
        return "forbidden", None
    info = f"{last_row['op_date']} · {last_row['op_type']} · {last_row['item_name']} · {last_row['quantity']:g} {last_row['unit']}".strip()
    cur.execute("DELETE FROM movements WHERE id = ?", (last_row["id"],))
    conn.commit()
    conn.close()
    return "ok", info
async def notify_admins(text: str):
    for admin_id in ADMIN_IDS:
        try:
            await bot.send_message(admin_id, text, parse_mode="Markdown")
        except Exception as e:
            logging.warning(f"Не удалось отправить уведомление админу {admin_id}: {e}")


def format_actor(user: types.User, chat: types.Chat) -> str:
    name = user.full_name or str(user.id)
    uname = f" (@{user.username})" if user.username else ""
    chat_info = f" в группе «{chat.title}»" if chat.type in ("group", "supergroup") else ""
    return f"{name}{uname} [ID: `{user.id}`]{chat_info}"


def parse_number(raw: str) -> Optional[float]:
    try:
        return float(raw.replace(",", ".").replace(" ", "").strip())
    except ValueError:
        return None


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


class LinkUserStates(StatesGroup):
    entering_tg_id = State()


def get_main_keyboard(has_workspace: bool, is_admin: bool) -> ReplyKeyboardMarkup:
    if not has_workspace:
        rows = [[KeyboardButton(text="🆕 Создать склад")]]
    else:
        rows = [
            [KeyboardButton(text="➕ Приход"), KeyboardButton(text="➖ Расход")],
            [KeyboardButton(text="📊 Текущий остаток"), KeyboardButton(text="📜 Последние операции")],
            [KeyboardButton(text="📝 Добавить позицию"), KeyboardButton(text="↩️ Отменить последнюю")],
            [KeyboardButton(text="👥 Привязать сотрудника"), KeyboardButton(text="ℹ️ Мой склад")],
        ]
    if is_admin:
        rows.append([KeyboardButton(text="👑 Все склады (Админ)")])
    return ReplyKeyboardMarkup(keyboard=rows, resize_keyboard=True)


async def ensure_workspace(message: types.Message) -> Optional[dict]:
    user = message.from_user
    db_upsert_user(user.id, user.username, user.full_name or str(user.id))
    ws = db_get_user_workspace(user.id)
    if not ws:
        await message.answer(
            "У вас ещё не создано рабочее пространство склада.\nНажмите **«🆕 Создать склад»**, чтобы начать работу.",
            reply_markup=get_main_keyboard(has_workspace=False, is_admin=user.id in ADMIN_IDS),
            parse_mode="Markdown",
        )
        return None
    return ws


@dp.message(Command("start"))
async def cmd_start(message: types.Message, state: FSMContext):
    await state.clear()
    user = message.from_user
    db_upsert_user(user.id, user.username, user.full_name or str(user.id))
    ws = db_get_user_workspace(user.id)
    is_admin = user.id in ADMIN_IDS
    if ws:
        text = f"Складской учёт готов к работе (SQLite).\nАктивное пространство: **#{ws['id']} — {ws['title']}**\nВыберите действие:"
    else:
        text = "Добро пожаловать в систему складского учёта!\nНажмите **«🆕 Создать склад»**, чтобы создать своё рабочее пространство."
    await message.answer(text, reply_markup=get_main_keyboard(has_workspace=bool(ws), is_admin=is_admin), parse_mode="Markdown")


@dp.message(F.text == "🆕 Создать склад")
async def create_user_workspace(message: types.Message, state: FSMContext):
    await state.clear()
    user = message.from_user
    db_upsert_user(user.id, user.username, user.full_name or str(user.id))
    existing = db_get_user_workspace(user.id)
    if existing:
        await message.answer(
            f"У вас уже есть активное пространство: **#{existing['id']} — {existing['title']}**",
            reply_markup=get_main_keyboard(has_workspace=True, is_admin=user.id in ADMIN_IDS),
            parse_mode="Markdown",
        )
        return
    title = f"Склад — {user.full_name or user.id}"
    ws_id = db_create_and_bind_workspace(user.id, title)
    await message.answer(
        f"✅ Рабочее пространство создано!\n• ID пространства: `#{ws_id}`\n• Название: **{title}**",
        reply_markup=get_main_keyboard(has_workspace=True, is_admin=user.id in ADMIN_IDS),
        parse_mode="Markdown",
    )
    await notify_admins(
        f"🆕 **Создано новое пространство склада**\n• Пользователь: {format_actor(user, message.chat)}\n• Пространство: `#{ws_id}` — *{title}*"
    )


@dp.message(F.text == "ℹ️ Мой склад")
async def my_workspace_info(message: types.Message):
    ws = await ensure_workspace(message)
    if not ws:
        return
    await message.answer(
        f"📦 **Текущее пространство:** `#{ws['id']}` — **{ws['title']}**\n"
        f"• Владелец ID: `{ws['owner_tg_id']}`\n"
        f"• Подключить сотрудника: `/link {ws['id']} <Telegram_ID>` или кнопка **«👥 Привязать сотрудника»**",
        parse_mode="Markdown",
    )


@dp.message(F.text == "👥 Привязать сотрудника")
async def start_link_user_ui(message: types.Message, state: FSMContext):
    ws = await ensure_workspace(message)
    if not ws:
        return
    user = message.from_user
    if user.id not in ADMIN_IDS and user.id != ws["owner_tg_id"]:
        await message.answer("Привязывать сотрудников может только владелец склада или администратор.")
        return
    await state.clear()
    await state.update_data(target_ws_id=ws["id"], target_ws_title=ws["title"])
    await state.set_state(LinkUserStates.entering_tg_id)
    cancel_markup = InlineKeyboardMarkup(inline_keyboard=[[InlineKeyboardButton(text="❌ Отмена", callback_data="cancel")]])
    await message.answer(
        f"Привязка к пространству **#{ws['id']} — {ws['title']}**.\nВведите **Telegram ID** сотрудника (числом):",
        reply_markup=cancel_markup,
        parse_mode="Markdown",
    )


@dp.message(LinkUserStates.entering_tg_id)
async def process_link_user_id(message: types.Message, state: FSMContext):
    raw = (message.text or "").strip()
    if not raw.isdigit():
        await message.answer("Telegram ID должен состоять только из цифр. Введите ID сотрудника:")
        return
    target_tg_id = int(raw)
    data = await state.get_data()
    ws_id = data["target_ws_id"]
    ws_title = data["target_ws_title"]
    db_bind_user_to_workspace(target_tg_id, ws_id)
    await state.clear()
    await message.answer(
        f"✅ Пользователь `{target_tg_id}` привязан к пространству **#{ws_id} — {ws_title}**.",
        reply_markup=get_main_keyboard(has_workspace=True, is_admin=message.from_user.id in ADMIN_IDS),
        parse_mode="Markdown",
    )


@dp.message(Command("link"))
async def cmd_link_user(message: types.Message):
    parts = (message.text or "").split()
    if len(parts) != 3 or not parts[1].isdigit() or not parts[2].isdigit():
        await message.answer("Формат команды: `/link <ID_склада> <Telegram_ID>`", parse_mode="Markdown")
        return
    ws_id = int(parts[1])
    target_tg_id = int(parts[2])
    caller_id = message.from_user.id
    conn = get_conn()
    cur = conn.cursor()
    cur.execute("SELECT owner_tg_id, title FROM workspaces WHERE id = ?", (ws_id,))
    row = cur.fetchone()
    conn.close()
    if not row:
        await message.answer("Пространство с таким ID не найдено.")
        return
    if caller_id not in ADMIN_IDS and caller_id != row["owner_tg_id"]:
        await message.answer("Привязывать пользователей может только владелец склада или администратор.")
        return
    db_bind_user_to_workspace(target_tg_id, ws_id)
    await message.answer(f"✅ Пользователь `{target_tg_id}` привязан к пространству **#{ws_id} — {row['title']}**.", parse_mode="Markdown")


@dp.message(F.text == "👑 Все склады (Админ)")
async def admin_list_workspaces(message: types.Message):
    if message.from_user.id not in ADMIN_IDS:
        return
    rows = db_list_workspaces()
    if not rows:
        await message.answer("Пока не создано ни одного рабочего пространства.")
        return
    lines = ["👑 **Все пространства в базе SQLite:**\n"]
    buttons = []
    for r in rows[:20]:
        lines.append(
            f"• `#{r['id']}` **{r['title']}** · владелец `{r['owner_tg_id']}` · людей: `{r['users_count']}` · операций: `{r['ops_count']}`"
        )
        buttons.append([InlineKeyboardButton(text=f"Переключиться на #{r['id']}", callback_data=f"switch_ws:{r['id']}")])
    await message.answer("\n".join(lines), reply_markup=InlineKeyboardMarkup(inline_keyboard=buttons), parse_mode="Markdown")


@dp.callback_query(F.data.startswith("switch_ws:"))
async def admin_switch_workspace(callback: types.CallbackQuery):
    if callback.from_user.id not in ADMIN_IDS:
        await callback.answer("Нет доступа", show_alert=True)
        return
    ws_id = int(callback.data.split(":", 1)[1])
    db_upsert_user(callback.from_user.id, callback.from_user.username, callback.from_user.full_name or "")
    if db_bind_user_to_workspace(callback.from_user.id, ws_id):
        ws = db_get_user_workspace(callback.from_user.id)
        await callback.message.edit_text(f"✅ Вы переключились на пространство **#{ws_id} — {ws['title']}**.", parse_mode="Markdown")
    else:
        await callback.message.edit_text("Пространство не найдено.")
    await callback.answer()


@dp.callback_query(F.data == "cancel")
async def cancel_operation(callback: types.CallbackQuery, state: FSMContext):
    await state.clear()
    await callback.message.edit_text("Действие отменено.")
    await callback.answer()


@dp.message(F.text == "📝 Добавить позицию")
async def start_add_catalog_item(message: types.Message, state: FSMContext):
    ws = await ensure_workspace(message)
    if not ws:
        return
    await state.clear()
    await state.update_data(workspace_id=ws["id"], workspace_title=ws["title"])
    await state.set_state(AddCatalogStates.entering_name)
    cancel_markup = InlineKeyboardMarkup(inline_keyboard=[[InlineKeyboardButton(text="❌ Отмена", callback_data="cancel")]])
    await message.answer("Введите **наименование** новой позиции:", reply_markup=cancel_markup, parse_mode="Markdown")


@dp.message(AddCatalogStates.entering_name)
async def process_catalog_name(message: types.Message, state: FSMContext):
    item_name = " ".join((message.text or "").strip().split())
    if not item_name:
        await message.answer("Наименование не может быть пустым. Введите название позиции:")
        return
    data = await state.get_data()
    ws_id = data["workspace_id"]
    if db_catalog_name_exists(ws_id, item_name):
        await message.answer(f"Позиция **«{item_name}»** уже есть в справочнике. Введите другое наименование:", parse_mode="Markdown")
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
        await message.answer("Укажите единицу измерения текстом или выберите кнопку выше:")
        return
    await ask_for_category(message, state, unit, edit=False)


async def save_catalog_item(message_or_call, state: FSMContext, category: str):
    data = await state.get_data()
    ws_id = data["workspace_id"]
    ws_title = data.get("workspace_title", "")
    name = data["new_name"]
    unit = data["new_unit"]
    cat_clean = " ".join((category or "").strip().split())
    if not cat_clean:
        msg = "Категория не может быть пустой. Введите категорию текстом или выберите из списка:"
        if isinstance(message_or_call, types.CallbackQuery):
            await message_or_call.message.answer(msg)
            await message_or_call.answer()
        else:
            await message_or_call.answer(msg)
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
            f"📝 **Новая позиция в справочнике**\n• Кто: {format_actor(user, chat)}\n• Склад: `#{ws_id}` — *{ws_title}*\n• Позиция: *{name}* ({unit}, {cat_clean})"
        )
    except Exception as e:
        result_text = f"❌ Ошибка при сохранении в справочник: {e}"
    await state.clear()
    if isinstance(message_or_call, types.CallbackQuery):
        await message_or_call.message.edit_text(result_text, parse_mode="Markdown")
        await message_or_call.answer()
    else:
        await message_or_call.answer(
            result_text,
            reply_markup=get_main_keyboard(has_workspace=True, is_admin=user.id in ADMIN_IDS),
            parse_mode="Markdown",
        )


@dp.callback_query(AddCatalogStates.entering_category, F.data.startswith("cat:"))
async def process_category_callback(callback: types.CallbackQuery, state: FSMContext):
    await save_catalog_item(callback, state, callback.data.split(":", 1)[1])


@dp.message(AddCatalogStates.entering_category)
async def process_category_text(message: types.Message, state: FSMContext):
    await save_catalog_item(message, state, message.text or "")


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
    ws_id = data["workspace_id"]
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
        await callback.message.edit_text("Введите дату в формате `ДД.ММ.ГГГГ` (например: `26.11.2026`):", parse_mode="Markdown")
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
        await message.answer("Неверный формат даты. Введите как `ДД.ММ.ГГГГ` (например, `26.11.2026`):", parse_mode="Markdown")
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
    ws_id = data["workspace_id"]
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
    qty = data.get("pending_qty", 0.0)
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
    ws_id = data["workspace_id"]
    ws_title = data.get("workspace_title", "")
    op_type = data["op_type"]
    item_name = data["item_name"]
    qty = data["quantity"]
    op_date = data["op_date"]
    unit = data.get("item_unit", "")
    category = data.get("item_category", "")
    comment = data.get("comment", "")
    user = callback.from_user
    chat = callback.message.chat
    actor_name = user.full_name or (f"@{user.username}" if user.username else str(user.id))
    low_stock_alert = ""
    try:
        db_add_movement(ws_id, op_date, op_type, item_name, qty, unit, category, comment, actor_name, user.id)
        if op_type == "Расход":
            stock_list = db_get_stock(ws_id)
            item_stock = next((x for x in stock_list if x["name"].lower() == item_name.lower()), None)
            rem_bal = item_stock["balance"] if item_stock else 0.0
            if rem_bal <= LOW_STOCK_THRESHOLD:
                low_stock_alert = f"\n⚠️ **Низкий остаток:** по «{item_name}» осталось **{rem_bal:g}** {unit}".strip()
        summary = (
            f"✅ **Запись добавлена!**\n"
            f"• Дата: {op_date}\n• Операция: {op_type}\n• Позиция: {item_name}\n"
            f"• Количество: {qty:g} {unit}".strip() + "\n"
            f"• Комментарий: {comment if comment else '—'}"
            f"{low_stock_alert}"
        )
        sign = "➕" if op_type == "Приход" else "➖"
        await notify_admins(
            f"{sign} **{op_type}** ({op_date})\n• Кто: {format_actor(user, chat)}\n• Склад: `#{ws_id}` — *{ws_title}*\n"
            f"• Позиция: *{item_name}* — *{qty:g} {unit}*\n• Комментарий: {comment if comment else '—'}{low_stock_alert}"
        )
    except Exception as e:
        summary = f"❌ Ошибка при записи: {e}"
    await state.clear()
    await callback.message.edit_text(summary, parse_mode="Markdown")
    await callback.answer()


@dp.message(F.text == "↩️ Отменить последнюю")
async def undo_last_movement(message: types.Message):
    ws = await ensure_workspace(message)
    if not ws:
        return
    user = message.from_user
    status, info = db_undo_last_movement(ws["id"], user.id, user.id in ADMIN_IDS)
    if status == "empty":
        await message.answer("В этом пространстве пока нет записей для отмены.")
        return
    if status == "forbidden":
        await message.answer("Вы можете отменить только свою последнюю запись. Обратитесь к администратору.")
        return
    await message.answer(f"↩️ Последняя запись удалена: **{info}**", parse_mode="Markdown")
    await notify_admins(
        f"↩️ **Отменена последняя запись**\n• Кто отменил: {format_actor(user, message.chat)}\n• Склад: `#{ws['id']}` — *{ws['title']}*\n• Запись: `{info}`"
    )


@dp.message(F.text == "📜 Последние операции")
async def recent_movements(message: types.Message):
    ws = await ensure_workspace(message)
    if not ws:
        return
    rows = db_get_recent_movements(ws["id"], limit=10)
    if not rows:
        await message.answer("Операций пока не было.")
        return
    lines = [f"📜 **Последние операции (#{ws['id']} — {ws['title']}):**\n"]
    for r in rows:
        sign = "➕" if r["op_type"] == "Приход" else "➖"
        extra = []
        if r["actor_name"]:
            extra.append(r["actor_name"])
        if r["comment"]:
            extra.append(f"«{r['comment']}»")
        extra_str = f" — {', '.join(extra)}" if extra else ""
        unit_str = f" {r['unit']}" if r["unit"] else ""
        lines.append(f"{sign} `{r['op_date']}` **{r['item_name']}**: {float(r['quantity']):g}{unit_str}{extra_str}")
    await message.answer("\n".join(lines), parse_mode="Markdown")


@dp.message(F.text == "📊 Текущий остаток")
async def show_stock(message: types.Message):
    ws = await ensure_workspace(message)
    if not ws:
        return
    stock_list = db_get_stock(ws["id"])
    if not stock_list:
        await message.answer("В справочнике пока нет позиций. Нажмите «📝 Добавить позицию».")
        return
    lines = [f"📦 **Текущие остатки (#{ws['id']} — {ws['title']}):**\n"]
    for item in stock_list:
        unit_str = f" {item['unit']}" if item["unit"] else ""
        lines.append(f"• **{item['name']}**: {item['balance']:g}{unit_str}")
    await message.answer("\n".join(lines), parse_mode="Markdown")


async def main():
    init_db()
    await dp.start_polling(bot)


if __name__ == "__main__":
    asyncio.run(main())
