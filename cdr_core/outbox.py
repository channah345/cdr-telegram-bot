"""PostgreSQL notification outbox with retry-safe claiming."""

import json
import os
import socket
from datetime import datetime, timezone

try:
    import psycopg
except Exception:
    psycopg = None


class NotificationOutbox:
    def __init__(self, database_url=""):
        self.database_url = str(database_url or os.getenv("DATABASE_URL", "")).strip()
        self.worker_id = f"{socket.gethostname()}:{os.getpid()}"

    @property
    def enabled(self):
        return bool(self.database_url and psycopg)

    def initialise(self):
        if not self.enabled:
            return False
        with psycopg.connect(self.database_url) as conn:
            with conn.cursor() as cur:
                cur.execute("""CREATE TABLE IF NOT EXISTS platform_notification_outbox (
                    id BIGSERIAL PRIMARY KEY, event_key TEXT NOT NULL UNIQUE, event_type TEXT NOT NULL,
                    payload JSONB NOT NULL DEFAULT '{}'::jsonb, status TEXT NOT NULL DEFAULT 'pending',
                    attempts INTEGER NOT NULL DEFAULT 0, available_at TIMESTAMPTZ NOT NULL DEFAULT NOW(),
                    locked_at TIMESTAMPTZ, locked_by TEXT NOT NULL DEFAULT '', sent_at TIMESTAMPTZ,
                    last_error TEXT NOT NULL DEFAULT '', created_at TIMESTAMPTZ NOT NULL DEFAULT NOW(),
                    updated_at TIMESTAMPTZ NOT NULL DEFAULT NOW())""")
                cur.execute("CREATE INDEX IF NOT EXISTS idx_platform_outbox_ready ON platform_notification_outbox(status,available_at,id)")
            conn.commit()
        return True

    def enqueue(self, event_key, event_type, payload):
        if not self.enabled:
            return False
        self.initialise()
        with psycopg.connect(self.database_url) as conn:
            with conn.cursor() as cur:
                cur.execute("""INSERT INTO platform_notification_outbox(event_key,event_type,payload)
                    VALUES(%s,%s,%s::jsonb) ON CONFLICT(event_key) DO NOTHING RETURNING id""",
                    (str(event_key), str(event_type), json.dumps(payload)))
                created = cur.fetchone()
            conn.commit()
        return bool(created)

    def claim(self, limit=20):
        if not self.enabled:
            return []
        self.initialise()
        with psycopg.connect(self.database_url) as conn:
            with conn.cursor() as cur:
                cur.execute("""WITH ready AS (
                    SELECT id FROM platform_notification_outbox
                    WHERE status IN ('pending','retry') AND available_at <= NOW()
                      AND (locked_at IS NULL OR locked_at < NOW() - INTERVAL '10 minutes')
                    ORDER BY id FOR UPDATE SKIP LOCKED LIMIT %s)
                    UPDATE platform_notification_outbox o SET status='sending',locked_at=NOW(),locked_by=%s,
                    attempts=attempts+1,updated_at=NOW() FROM ready WHERE o.id=ready.id
                    RETURNING o.id,o.event_key,o.event_type,o.payload,o.attempts""", (int(limit), self.worker_id))
                rows = cur.fetchall()
            conn.commit()
        return [{"id":r[0],"event_key":r[1],"event_type":r[2],"payload":r[3] or {},"attempts":r[4]} for r in rows]

    def mark_sent(self, event_id):
        self._finish(event_id, "sent", "")

    def mark_failed(self, event_id, error, attempts=1):
        status = "failed" if int(attempts) >= 8 else "retry"
        delay_minutes = min(60, 2 ** min(int(attempts), 5))
        if not self.enabled:
            return
        with psycopg.connect(self.database_url) as conn:
            with conn.cursor() as cur:
                cur.execute("""UPDATE platform_notification_outbox SET status=%s,last_error=%s,
                    available_at=NOW()+(%s || ' minutes')::interval,locked_at=NULL,locked_by='',updated_at=NOW()
                    WHERE id=%s""", (status, str(error)[:2000], str(delay_minutes), event_id))
            conn.commit()

    def _finish(self, event_id, status, error):
        if not self.enabled:
            return
        with psycopg.connect(self.database_url) as conn:
            with conn.cursor() as cur:
                cur.execute("""UPDATE platform_notification_outbox SET status=%s,last_error=%s,sent_at=%s,
                    locked_at=NULL,locked_by='',updated_at=NOW() WHERE id=%s""",
                    (status, error, datetime.now(timezone.utc), event_id))
            conn.commit()
