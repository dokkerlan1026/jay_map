import pytest

import database


def test_missing_database_url_is_rejected(monkeypatch):
    monkeypatch.delenv("DATABASE_URL", raising=False)
    with pytest.raises(RuntimeError, match="DATABASE_URL is required"):
        database.get_database_url()


def test_init_db_creates_tables(postgres_db):
    rows = postgres_db.execute(
        "SELECT table_name FROM information_schema.tables WHERE table_schema = 'public'"
    ).fetchall()
    names = {r["table_name"] for r in rows}
    assert {"store_daily_state", "version_history", "sync_meta"} <= names


def test_healthcheck(clean_db):
    assert database.healthcheck() is True


def test_date_initialization_is_idempotent(postgres_db):
    database.ensure_date_initialized(postgres_db, "2026-10-23", 3)
    database.ensure_date_initialized(postgres_db, "2026-10-23", 3)
    count = postgres_db.execute(
        "SELECT COUNT(*) AS count FROM store_daily_state WHERE date = %s",
        ("2026-10-23",),
    ).fetchone()["count"]
    histories = postgres_db.execute(
        "SELECT COUNT(*) AS count FROM version_history WHERE date = %s",
        ("2026-10-23",),
    ).fetchone()["count"]
    assert count == 3
    assert histories == 1
