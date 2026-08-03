"""Database advisory leases prevent duplicate scheduled work across replicas."""

import os

try:
    import psycopg
except Exception:
    psycopg = None


class SchedulerLease:
    def __init__(self, connection=None):
        self.connection = connection

    @classmethod
    def acquire(cls, name, database_url=""):
        database_url = str(database_url or os.getenv("DATABASE_URL", "")).strip()
        if not database_url or not psycopg:
            return cls(None)
        conn = psycopg.connect(database_url, autocommit=True)
        with conn.cursor() as cur:
            cur.execute("SELECT pg_try_advisory_lock(hashtext(%s))", (f"cdr-scheduler:{name}",))
            if not bool((cur.fetchone() or [False])[0]):
                conn.close()
                return None
        return cls(conn)

    def release(self):
        if self.connection:
            self.connection.close()
            self.connection = None
