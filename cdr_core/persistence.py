"""Safe PostgreSQL persistence for python-telegram-bot conversations."""

import asyncio
import base64
from datetime import date, datetime
import json

try:
    import psycopg
except Exception:
    psycopg = None

from telegram.ext import BasePersistence, PersistenceInput


def _encode(value):
    if isinstance(value, datetime):
        return {"__cdr_type__": "datetime", "value": value.isoformat()}
    if isinstance(value, date):
        return {"__cdr_type__": "date", "value": value.isoformat()}
    if isinstance(value, bytes):
        return {"__cdr_type__": "bytes", "value": base64.b64encode(value).decode("ascii")}
    if isinstance(value, tuple):
        return {"__cdr_type__": "tuple", "value": list(value)}
    if isinstance(value, set):
        return {"__cdr_type__": "set", "value": list(value)}
    if hasattr(value, "to_dict"):
        return {"__cdr_type__": "telegram", "value": value.to_dict()}
    raise TypeError(f"Unsupported persisted type: {type(value).__name__}")


def _decode(value):
    if not isinstance(value, dict) or "__cdr_type__" not in value:
        return value
    kind, raw = value.get("__cdr_type__"), value.get("value")
    if kind == "datetime":
        return datetime.fromisoformat(raw)
    if kind == "date":
        return date.fromisoformat(raw)
    if kind == "bytes":
        return base64.b64decode(raw)
    if kind == "tuple":
        return tuple(raw)
    if kind == "set":
        return set(raw)
    return raw


def _json(value):
    return json.dumps(value, default=_encode, ensure_ascii=False)


class PostgresPersistence(BasePersistence):
    """Stores each user/chat/conversation independently as JSONB.

    No pickle is loaded from disk, and all database I/O is moved away from the
    Telegram event loop.
    """
    def __init__(self, database_url: str, update_interval: float = 20):
        if not database_url or not psycopg:
            raise RuntimeError("PostgreSQL persistence requires DATABASE_URL and psycopg")
        super().__init__(
            store_data=PersistenceInput(bot_data=True, chat_data=True, user_data=True, callback_data=False),
            update_interval=update_interval,
        )
        self.database_url = database_url
        self._initialised = False

    def _connect(self):
        return psycopg.connect(self.database_url)

    def _init(self):
        if self._initialised:
            return
        with self._connect() as conn:
            with conn.cursor() as cur:
                cur.execute("""CREATE TABLE IF NOT EXISTS bot_persistence (
                    scope TEXT NOT NULL, state_key TEXT NOT NULL,
                    state_data JSONB NOT NULL DEFAULT '{}'::jsonb,
                    updated_at TIMESTAMPTZ NOT NULL DEFAULT NOW(),
                    PRIMARY KEY(scope,state_key))""")
            conn.commit()
        self._initialised = True

    def _load_scope(self, scope):
        self._init()
        with self._connect() as conn:
            with conn.cursor() as cur:
                cur.execute("SELECT state_key,state_data FROM bot_persistence WHERE scope=%s", (scope,))
                return {key: json.loads(json.dumps(data), object_hook=_decode) for key, data in cur.fetchall()}

    def _write(self, scope, key, data):
        self._init()
        with self._connect() as conn:
            with conn.cursor() as cur:
                cur.execute("""INSERT INTO bot_persistence(scope,state_key,state_data,updated_at)
                    VALUES(%s,%s,%s::jsonb,NOW()) ON CONFLICT(scope,state_key) DO UPDATE
                    SET state_data=EXCLUDED.state_data,updated_at=NOW()""", (scope, str(key), _json(data)))
            conn.commit()

    def _delete(self, scope, key):
        self._init()
        with self._connect() as conn:
            with conn.cursor() as cur:
                cur.execute("DELETE FROM bot_persistence WHERE scope=%s AND state_key=%s", (scope, str(key)))
            conn.commit()

    async def get_user_data(self):
        rows = await asyncio.to_thread(self._load_scope, "user")
        return {int(key): value for key, value in rows.items()}

    async def get_chat_data(self):
        rows = await asyncio.to_thread(self._load_scope, "chat")
        return {int(key): value for key, value in rows.items()}

    async def get_bot_data(self):
        rows = await asyncio.to_thread(self._load_scope, "bot")
        return rows.get("global", {})

    async def get_callback_data(self):
        return None

    async def get_conversations(self, name):
        rows = await asyncio.to_thread(self._load_scope, f"conversation:{name}")
        conversations = {}
        for key, value in rows.items():
            conversations[tuple(json.loads(key))] = value.get("state")
        return conversations

    async def update_user_data(self, user_id, data):
        await asyncio.to_thread(self._write, "user", user_id, dict(data))

    async def update_chat_data(self, chat_id, data):
        await asyncio.to_thread(self._write, "chat", chat_id, dict(data))

    async def update_bot_data(self, data):
        await asyncio.to_thread(self._write, "bot", "global", dict(data))

    async def update_callback_data(self, data):
        return None

    async def update_conversation(self, name, key, new_state):
        encoded_key = json.dumps(list(key), separators=(",", ":"))
        if new_state is None:
            await asyncio.to_thread(self._delete, f"conversation:{name}", encoded_key)
        else:
            await asyncio.to_thread(self._write, f"conversation:{name}", encoded_key, {"state": new_state})

    async def drop_user_data(self, user_id):
        await asyncio.to_thread(self._delete, "user", user_id)

    async def drop_chat_data(self, chat_id):
        await asyncio.to_thread(self._delete, "chat", chat_id)

    async def refresh_user_data(self, user_id, user_data):
        return None

    async def refresh_chat_data(self, chat_id, chat_data):
        return None

    async def refresh_bot_data(self, bot_data):
        return None

    async def flush(self):
        return None


class ProcessedUpdateStore:
    def __init__(self, database_url: str):
        self.database_url = str(database_url or "").strip()

    @property
    def enabled(self):
        return bool(self.database_url and psycopg)

    def _init(self, cur):
        cur.execute("""CREATE TABLE IF NOT EXISTS bot_processed_updates (
            update_id BIGINT PRIMARY KEY, processed_at TIMESTAMPTZ NOT NULL DEFAULT NOW())""")

    def contains(self, update_id):
        if not self.enabled:
            return False
        with psycopg.connect(self.database_url) as conn:
            with conn.cursor() as cur:
                self._init(cur)
                cur.execute("SELECT 1 FROM bot_processed_updates WHERE update_id=%s", (int(update_id),))
                return bool(cur.fetchone())

    def mark(self, update_id):
        if not self.enabled:
            return
        with psycopg.connect(self.database_url) as conn:
            with conn.cursor() as cur:
                self._init(cur)
                cur.execute("INSERT INTO bot_processed_updates(update_id) VALUES(%s) ON CONFLICT DO NOTHING", (int(update_id),))
                cur.execute("DELETE FROM bot_processed_updates WHERE processed_at < NOW() - INTERVAL '30 days'")
            conn.commit()
