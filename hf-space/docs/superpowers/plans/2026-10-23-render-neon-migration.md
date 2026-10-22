# Render + Neon Migration Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Replace the Hugging Face/Gradio/SQLite runtime with a Render-hosted FastAPI service backed by a fresh Neon PostgreSQL database.

**Architecture:** A focused `database.py` module owns psycopg connections, schema creation, health checks, and atomic date initialization. `app.py` keeps the current HTTP contract while using PostgreSQL-native placeholders, booleans, transactions, and `RETURNING`. A repository-root Render Blueprint deploys the `hf-space` subdirectory and receives `DATABASE_URL` as a secret.

**Tech Stack:** Python 3.12, FastAPI, Uvicorn, psycopg 3 binary package, PostgreSQL 16, pytest, Docker, Render, Neon

---

### Task 1: PostgreSQL database module

**Files:**
- Create: `hf-space/database.py`
- Create: `hf-space/tests/conftest.py`
- Create: `hf-space/tests/test_database.py`
- Modify: `hf-space/requirements.txt`

- [ ] **Step 1: Add the failing database tests**

Create tests that set `DATABASE_URL` to the disposable PostgreSQL test database, call `init_db()`, and assert that all three tables exist. Add tests that remove `DATABASE_URL` and expect `RuntimeError("DATABASE_URL is required")`, call `healthcheck()` and expect `True`, and call `ensure_date_initialized(conn, date, store_count)` twice while asserting exactly `store_count` state rows and one initial history row.

```python
def test_missing_database_url_is_rejected(monkeypatch):
    monkeypatch.delenv("DATABASE_URL", raising=False)
    with pytest.raises(RuntimeError, match="DATABASE_URL is required"):
        database.get_database_url()

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
```

- [ ] **Step 2: Run the tests to verify RED**

Run: `$env:DATABASE_URL='postgresql://postgres:postgres@localhost:55432/hualien_test'; pytest tests/test_database.py -v`

Expected: collection fails because `database.py` does not exist.

- [ ] **Step 3: Implement the minimal database module**

Implement:

```python
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

def healthcheck() -> bool:
    with get_db() as conn:
        return conn.execute("SELECT 1 AS ok").fetchone()["ok"] == 1

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
```

Use PostgreSQL identity for `version_history.id`, boolean columns for state flags, `%s` placeholders, `ON CONFLICT`, and `psycopg.rows.dict_row`.

- [ ] **Step 4: Run database tests to verify GREEN**

Run the command from Step 2.

Expected: all database tests pass.

- [ ] **Step 5: Commit the database module**

```powershell
git add hf-space/database.py hf-space/tests hf-space/requirements.txt
git commit -m "feat: add Neon PostgreSQL persistence"
```

### Task 2: Port the FastAPI application from SQLite to PostgreSQL

**Files:**
- Modify: `hf-space/app.py`
- Create: `hf-space/tests/test_app.py`

- [ ] **Step 1: Add failing API tests**

Test the lifespan startup, `/health`, `/`, `/api/stores`, visibility toggle, check-in toggle, named-version save, and version revert against the disposable PostgreSQL database. Assert the existing JSON field names and counts.

```python
def test_store_workflow(client):
    initial = client.get("/api/stores?date=2026-10-23")
    assert initial.status_code == 200
    assert initial.json()["activeCount"] == 500

    hidden = client.post("/api/stores/toggle-visibility", json={
        "date": "2026-10-23", "store_id": 1, "is_visible": False
    })
    assert hidden.status_code == 200

    refreshed = client.get("/api/stores?date=2026-10-23").json()
    assert refreshed["activeCount"] == 499
    assert refreshed["stores"][0]["isVisible"] is False
```

- [ ] **Step 2: Run the API tests to verify RED**

Run: `$env:DATABASE_URL='postgresql://postgres:postgres@localhost:55432/hualien_test'; pytest tests/test_app.py -v`

Expected: failure because `app.py` still imports SQLite and does not expose `/health`.

- [ ] **Step 3: Implement PostgreSQL application access**

Remove `sqlite3`, `DB_PATH`, local `get_db`, local `init_db`, and the Gradio mounting block. Import `get_db`, `init_db`, `healthcheck`, and `ensure_date_initialized` from `database.py`. Initialize the schema from a FastAPI lifespan function.

For every SQL statement:

- Replace `?` with `%s`.
- Pass Python `True`/`False` to boolean columns.
- Replace `conn.executemany(...)` with `conn.cursor().executemany(...)`.
- Replace sync metadata `INSERT OR REPLACE` with `INSERT ... ON CONFLICT (key) DO UPDATE`.
- Replace `cursor.lastrowid` with `INSERT ... RETURNING id` and `fetchone()["id"]`.
- Remove internal `commit()` calls and let the psycopg connection context commit or roll back the whole request.

Add:

```python
@app.get("/health")
def health():
    if not healthcheck():
        raise HTTPException(status_code=503, detail="Database unavailable")
    return {"status": "ok"}
```

Register a psycopg exception handler that returns HTTP 503 JSON without exposing credentials.

```python
@app.exception_handler(psycopg.Error)
async def database_error_handler(_request: Request, _exc: psycopg.Error):
    return JSONResponse(
        status_code=503,
        content={"detail": "Database temporarily unavailable"},
    )
```

- [ ] **Step 4: Run API and database tests to verify GREEN**

Run: `$env:DATABASE_URL='postgresql://postgres:postgres@localhost:55432/hualien_test'; pytest -v`

Expected: all tests pass.

- [ ] **Step 5: Commit the FastAPI migration**

```powershell
git add hf-space/app.py hf-space/tests/test_app.py
git commit -m "refactor: run FastAPI directly on PostgreSQL"
```

### Task 3: Add Render deployment configuration

**Files:**
- Create: `render.yaml`
- Modify: `hf-space/README.md`
- Modify: `hf-space/requirements.txt`

- [ ] **Step 1: Add a failing deployment configuration test**

Add `tests/test_deployment_config.py` that loads the repository-root `render.yaml` and asserts `rootDir == "hf-space"`, the start command uses `$PORT`, health check is `/health`, and `DATABASE_URL` has `sync: false`. Assert requirements contain `psycopg[binary]` and do not contain Gradio, Hugging Face Hub, or SQLite packages.

- [ ] **Step 2: Run the deployment test to verify RED**

Run: `pytest tests/test_deployment_config.py -v`

Expected: failure because `render.yaml` does not exist.

- [ ] **Step 3: Implement the Render Blueprint and dependency list**

Create repository-root `render.yaml`:

```yaml
services:
  - type: web
    name: hualien-store-map
    runtime: python
    rootDir: hf-space
    buildCommand: pip install -r requirements.txt
    startCommand: uvicorn app:app --host 0.0.0.0 --port $PORT
    healthCheckPath: /health
    envVars:
      - key: DATABASE_URL
        sync: false
```

Remove Gradio and Hugging Face Hub requirements. Retain FastAPI, Uvicorn, OpenPyXL, and Pydantic; add pinned-compatible ranges for `psycopg[binary]`, `pytest`, `httpx`, and `PyYAML`.

Rewrite the README with exact Neon creation, connection-string copy, GitHub push, Render Blueprint, secret entry, and first-deploy verification steps. Do not include a real connection string.

- [ ] **Step 4: Run deployment tests to verify GREEN**

Run: `pytest tests/test_deployment_config.py -v`

Expected: all deployment configuration tests pass.

- [ ] **Step 5: Commit deployment configuration**

```powershell
git add render.yaml hf-space/README.md hf-space/requirements.txt hf-space/tests/test_deployment_config.py
git commit -m "chore: configure Render deployment"
```

### Task 4: Full verification

**Files:**
- Verify all modified files

- [ ] **Step 1: Run the full test suite**

Run: `$env:DATABASE_URL='postgresql://postgres:postgres@localhost:55432/hualien_test'; pytest -v`

Expected: all tests pass with zero failures.

- [ ] **Step 2: Run static verification**

Run: `python -m py_compile app.py database.py`

Expected: exit code 0.

Run: `rg -n "sqlite3|gradio|huggingface_hub|HfFolder" app.py database.py requirements.txt`

Expected: no matches.

- [ ] **Step 3: Run the Render-equivalent smoke test**

Start: `$env:PORT='18080'; uvicorn app:app --host 127.0.0.1 --port $env:PORT`

Request `/health` and `/`; expect HTTP 200 from both, then stop Uvicorn.

- [ ] **Step 4: Review repository state**

Run: `git status --short` and `git diff --check`.

Expected: only intentional migration files are modified; no database files, caches, credentials, or test artifacts are present.

- [ ] **Step 5: Commit any verification-only corrections**

If verification required a source correction, commit only that correction with a focused message. Otherwise do not create an empty commit.
