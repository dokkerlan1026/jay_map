import os
import sys

import pytest

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

import database  # noqa: E402


def _require_database_url():
    if not os.getenv("DATABASE_URL"):
        pytest.skip("DATABASE_URL not set; disposable PostgreSQL required")


@pytest.fixture
def clean_db():
    _require_database_url()
    database.init_db()
    with database.get_db() as conn:
        conn.execute("TRUNCATE store_daily_state, version_history, sync_meta RESTART IDENTITY")
    yield


@pytest.fixture
def postgres_db(clean_db):
    with database.get_db() as conn:
        yield conn


@pytest.fixture
def client(clean_db):
    from fastapi.testclient import TestClient

    import app

    with TestClient(app.app) as test_client:
        yield test_client
