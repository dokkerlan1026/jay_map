import os
from datetime import datetime

import psycopg
from psycopg.rows import dict_row


def get_database_url() -> str:
    value = os.getenv("DATABASE_URL", "").strip()
    if not value:
        raise RuntimeError("DATABASE_URL is required")
    return value


def get_db() -> psycopg.Connection:
    return psycopg.connect(get_database_url(), row_factory=dict_row)


def init_db() -> None:
    with get_db() as conn:
        conn.execute("""
            CREATE TABLE IF NOT EXISTS store_daily_state (
                date TEXT NOT NULL,
                store_id INTEGER NOT NULL,
                is_visible BOOLEAN NOT NULL DEFAULT TRUE,
                is_checked_in BOOLEAN NOT NULL DEFAULT FALSE,
                checkin_time TEXT,
                updated_at TEXT NOT NULL,
                PRIMARY KEY (date, store_id)
            )
        """)
        conn.execute("""
            CREATE TABLE IF NOT EXISTS version_history (
                id INTEGER GENERATED ALWAYS AS IDENTITY PRIMARY KEY,
                date TEXT NOT NULL,
                action TEXT,
                version_name TEXT,
                is_named BOOLEAN NOT NULL DEFAULT FALSE,
                active_count INTEGER,
                closed_count INTEGER,
                hidden_ids TEXT NOT NULL DEFAULT '',
                checked_ids TEXT NOT NULL DEFAULT '',
                created_at TEXT NOT NULL
            )
        """)
        conn.execute("""
            CREATE TABLE IF NOT EXISTS sync_meta (
                key TEXT PRIMARY KEY,
                value TEXT
            )
        """)


def healthcheck() -> bool:
    with get_db() as conn:
        return conn.execute("SELECT 1 AS ok").fetchone()["ok"] == 1


# 確保當日資料初始化 (若無則預設全開)；以 sync_meta 標記保證並發下只初始化一次
def ensure_date_initialized(conn, target_date: str, store_count: int) -> None:
    now_string = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    marker = conn.execute(
        """INSERT INTO sync_meta (key, value) VALUES (%s, %s)
           ON CONFLICT (key) DO NOTHING RETURNING key""",
        (f"initialized_{target_date}", now_string),
    ).fetchone()
    if marker is None:
        return
    rows = [
        (target_date, store_id, True, False, "", now_string)
        for store_id in range(1, store_count + 1)
    ]
    with conn.cursor() as cursor:
        cursor.executemany(
            """INSERT INTO store_daily_state
               (date, store_id, is_visible, is_checked_in, checkin_time, updated_at)
               VALUES (%s, %s, %s, %s, %s, %s)
               ON CONFLICT (date, store_id) DO NOTHING""",
            rows,
        )
    conn.execute(
        """INSERT INTO version_history
           (date, action, version_name, is_named, active_count, closed_count,
            hidden_ids, checked_ids, created_at)
           VALUES (%s, %s, %s, %s, %s, %s, %s, %s, %s)""",
        (target_date, f"初始預設 ({store_count}家全部開啟)", "預設全開",
         False, store_count, 0, "", "", now_string),
    )
    conn.execute(
        """INSERT INTO sync_meta (key, value) VALUES (%s, %s)
           ON CONFLICT (key) DO UPDATE SET value = EXCLUDED.value""",
        (f"last_update_{target_date}", now_string),
    )
