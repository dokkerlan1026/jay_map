import os
import json
from datetime import datetime, timedelta
from typing import List, Optional
from io import BytesIO

from contextlib import asynccontextmanager

import psycopg
from fastapi import FastAPI, HTTPException, Request
from fastapi.responses import HTMLResponse, JSONResponse, StreamingResponse
from fastapi.staticfiles import StaticFiles
from pydantic import BaseModel
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side

from database import ensure_date_initialized as db_ensure_date_initialized
from database import get_db, healthcheck, init_db

# ==========================================
# 基礎設定與資料載入
# ==========================================
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
DATA_FILE = os.path.join(BASE_DIR, "stores_data.json")

# 載入基礎 500 筆店家資料
with open(DATA_FILE, "r", encoding="utf-8") as f:
    DEFAULT_STORES = json.load(f)

# 依店名精確校準覆寫
PRECISE_FIXES = {
    "store-1":   {"lat": 23.973890, "lng": 121.578260, "subGroup": "主要名單-西區"},
    "store-5":   {"lat": 23.978850, "lng": 121.588100, "subGroup": "非488名單"},
    "store-66":  {"lat": 23.978500, "lng": 121.608800, "subGroup": "主要名單-東區"},
    "store-142": {"lat": 23.971250, "lng": 121.600350, "subGroup": "非488名單"},
    "store-143": {"lat": 23.969313, "lng": 121.595063, "subGroup": "主要名單-東區"},
    "store-145": {"lat": 23.969420, "lng": 121.595150, "subGroup": "主要名單-東區"},
    "store-191": {"lat": 23.970578, "lng": 121.596152, "subGroup": "主要名單-東區"},
    "store-266": {"lat": 23.974800, "lng": 121.609200, "subGroup": "主要名單-東區"},
    "store-288": {"lat": 23.986200, "lng": 121.599300, "subGroup": "主要名單-西區"}
}

for s in DEFAULT_STORES:
    if s["id"] in PRECISE_FIXES:
        s["lat"] = PRECISE_FIXES[s["id"]]["lat"]
        s["lng"] = PRECISE_FIXES[s["id"]]["lng"]
        s["subGroup"] = PRECISE_FIXES[s["id"]]["subGroup"]

# ==========================================
# 應用程式與資料庫初始化
# ==========================================
@asynccontextmanager
async def lifespan(_app: FastAPI):
    init_db()
    yield

app = FastAPI(title="花蓮店家排程地圖", version="3.0.0", lifespan=lifespan)

@app.exception_handler(psycopg.Error)
async def database_error_handler(_request: Request, _exc: psycopg.Error):
    return JSONResponse(
        status_code=503,
        content={"detail": "Database temporarily unavailable"},
    )

@app.get("/health")
def health():
    if not healthcheck():
        raise HTTPException(status_code=503, detail="Database unavailable")
    return {"status": "ok"}

def get_now_str():
    return datetime.now().strftime("%Y-%m-%d %H:%M:%S")

def get_today_str():
    return datetime.now().strftime("%Y-%m-%d")

def get_yesterday_str(base_date: str):
    dt = datetime.strptime(base_date, "%Y-%m-%d")
    return (dt - timedelta(days=1)).strftime("%Y-%m-%d")

def update_sync_time(conn, target_date: str):
    now_str = get_now_str()
    conn.execute(
        """INSERT INTO sync_meta (key, value) VALUES (%s, %s)
           ON CONFLICT (key) DO UPDATE SET value = EXCLUDED.value""",
        (f"last_update_{target_date}", now_str),
    )
    return now_str

def get_sync_time(conn, target_date: str) -> str:
    row = conn.execute("SELECT value FROM sync_meta WHERE key = %s", (f"last_update_{target_date}",)).fetchone()
    return row["value"] if row else ""

def ensure_date_initialized(conn, target_date: str):
    db_ensure_date_initialized(conn, target_date, len(DEFAULT_STORES))

# ==========================================
# Pydantic 請求模型
# ==========================================
class VisibilityToggleRequest(BaseModel):
    date: str
    store_id: int
    is_visible: bool

class BatchVisibilityRequest(BaseModel):
    date: str
    hidden_ids: List[int]
    action_note: Optional[str] = "批次設定"

class CheckinToggleRequest(BaseModel):
    date: str
    store_id: int
    is_checked_in: bool

class SaveNamedVersionRequest(BaseModel):
    date: str
    version_name: str
    hidden_ids: List[int]
    checked_ids: List[int]

class RevertVersionRequest(BaseModel):
    date: str
    version_id: int

# ==========================================
# REST API 端點
# ==========================================

@app.get("/api/stores")
def get_stores(date: Optional[str] = None):
    target_date = date or get_today_str()
    with get_db() as conn:
        ensure_date_initialized(conn, target_date)
        rows = conn.execute("""
            SELECT store_id, is_visible, is_checked_in, checkin_time, updated_at
            FROM store_daily_state
            WHERE date = %s
            ORDER BY store_id ASC
        """, (target_date,)).fetchall()
        
        state_map = {r["store_id"]: r for r in rows}
        sync_time = get_sync_time(conn, target_date)
        
        result_stores = []
        active_count = 0
        checked_count = 0
        
        for idx, s in enumerate(DEFAULT_STORES):
            num = idx + 1
            st = state_map.get(num)
            is_vis = bool(st["is_visible"]) if st else True
            is_chk = bool(st["is_checked_in"]) if st else False
            chk_time = st["checkin_time"] if st else ""
            
            if is_vis: active_count += 1
            if is_chk: checked_count += 1
            
            result_stores.append({
                **s,
                "isVisible": is_vis,
                "isCheckedIn": is_chk,
                "checkinTime": chk_time
            })
            
        return {
            "success": True,
            "date": target_date,
            "stores": result_stores,
            "totalCount": len(DEFAULT_STORES),
            "activeCount": active_count,
            "closedCount": len(DEFAULT_STORES) - active_count,
            "checkedCount": checked_count,
            "serverTime": sync_time
        }

@app.get("/api/stores/poll")
def poll_collaborative_state(date: Optional[str] = None, client_time: Optional[str] = ""):
    target_date = date or get_today_str()
    with get_db() as conn:
        sync_time = get_sync_time(conn, target_date)
        if client_time and sync_time and client_time == sync_time:
            return {"success": True, "hasUpdate": False, "serverTime": sync_time}
            
        rows = conn.execute("""
            SELECT store_id, is_visible, is_checked_in
            FROM store_daily_state
            WHERE date = %s
        """, (target_date,)).fetchall()
        
        hidden_ids = [r["store_id"] for r in rows if not r["is_visible"]]
        checked_ids = [r["store_id"] for r in rows if r["is_checked_in"]]
        active_count = len(DEFAULT_STORES) - len(hidden_ids)
        
        return {
            "success": True,
            "hasUpdate": True,
            "date": target_date,
            "serverTime": sync_time,
            "hiddenIds": hidden_ids,
            "checkedInIds": checked_ids,
            "activeCount": active_count,
            "closedCount": len(hidden_ids),
            "checkedCount": len(checked_ids)
        }

@app.post("/api/stores/toggle-visibility")
def toggle_visibility(req: VisibilityToggleRequest):
    with get_db() as conn:
        now_str = get_now_str()
        conn.execute("""
            UPDATE store_daily_state
            SET is_visible = %s, updated_at = %s
            WHERE date = %s AND store_id = %s
        """, (req.is_visible, now_str, req.date, req.store_id))
        
        # 記錄狀態變更並更新時間戳記
        sync_time = update_sync_time(conn, req.date)
        return {"success": True, "serverTime": sync_time}

@app.post("/api/stores/batch-visibility")
def batch_visibility(req: BatchVisibilityRequest):
    with get_db() as conn:
        now_str = get_now_str()
        hidden_set = set(req.hidden_ids)
        
        # 批次更新
        updates = [((i + 1) not in hidden_set, now_str, req.date, i + 1) for i in range(len(DEFAULT_STORES))]
        conn.cursor().executemany("""
            UPDATE store_daily_state
            SET is_visible = %s, updated_at = %s
            WHERE date = %s AND store_id = %s
        """, updates)
        
        active_count = len(DEFAULT_STORES) - len(hidden_set)
        # 寫入一筆歷史存檔
        hidden_str = ",".join(map(str, req.hidden_ids))
        conn.execute("""
            INSERT INTO version_history (date, action, version_name, is_named, active_count, closed_count, hidden_ids, checked_ids, created_at)
            VALUES (%s, %s, %s, %s, %s, %s, %s, %s, %s)
        """, (req.date, req.action_note, "", False, active_count, len(hidden_set), hidden_str, "", now_str))
        
        sync_time = update_sync_time(conn, req.date)
        return {"success": True, "activeCount": active_count, "serverTime": sync_time}

@app.post("/api/stores/toggle-checkin")
def toggle_checkin(req: CheckinToggleRequest):
    with get_db() as conn:
        now_str = get_now_str()
        chk_time = datetime.now().strftime("%H:%M") if req.is_checked_in else ""
        conn.execute("""
            UPDATE store_daily_state
            SET is_checked_in = %s, checkin_time = %s, updated_at = %s
            WHERE date = %s AND store_id = %s
        """, (req.is_checked_in, chk_time, now_str, req.date, req.store_id))
        
        sync_time = update_sync_time(conn, req.date)
        return {"success": True, "checkinTime": chk_time, "serverTime": sync_time}

@app.post("/api/stores/load-all-open")
def load_all_open(date: Optional[str] = None):
    target_date = date or get_today_str()
    with get_db() as conn:
        now_str = get_now_str()
        conn.execute("""
            UPDATE store_daily_state
            SET is_visible = TRUE, updated_at = %s
            WHERE date = %s
        """, (now_str, target_date))
        
        conn.execute("""
            INSERT INTO version_history (date, action, version_name, is_named, active_count, closed_count, hidden_ids, checked_ids, created_at)
            VALUES (%s, %s, %s, %s, %s, %s, %s, %s, %s)
        """, (target_date, "☀️ 載入全開啟 (500家全部顯示)", "", False, len(DEFAULT_STORES), 0, "", "", now_str))
        
        sync_time = update_sync_time(conn, target_date)
        return {"success": True, "activeCount": len(DEFAULT_STORES), "closedCount": 0, "serverTime": sync_time}

@app.post("/api/stores/load-yesterday")
def load_yesterday(date: Optional[str] = None):
    target_date = date or get_today_str()
    yesterday = get_yesterday_str(target_date)
    
    with get_db() as conn:
        # 尋找昨天最後一次存檔紀錄
        last_hist = conn.execute("""
            SELECT * FROM version_history
            WHERE date = %s
            ORDER BY id DESC
            LIMIT 1
        """, (yesterday,)).fetchone()
        
        if not last_hist:
            # 若歷史表沒有，直接讀取昨天的 daily_state
            y_rows = conn.execute("""
                SELECT store_id, is_visible, is_checked_in
                FROM store_daily_state
                WHERE date = %s
            """, (yesterday,)).fetchall()
            if not y_rows:
                raise HTTPException(status_code=404, detail=f"找不到昨天 ({yesterday}) 的任何存檔或排程紀錄！")
            hidden_ids = [r["store_id"] for r in y_rows if not r["is_visible"]]
            checked_ids = [r["store_id"] for r in y_rows if r["is_checked_in"]]
        else:
            hidden_ids = [int(x) for x in last_hist["hidden_ids"].split(",") if x.strip()]
            checked_ids = [int(x) for x in last_hist["checked_ids"].split(",") if x.strip()]
            
        hidden_set = set(hidden_ids)
        checked_set = set(checked_ids)
        now_str = get_now_str()
        
        updates = []
        for i in range(len(DEFAULT_STORES)):
            s_id = i + 1
            vis = s_id not in hidden_set
            chk = s_id in checked_set
            updates.append((vis, chk, now_str, target_date, s_id))
            
        conn.cursor().executemany("""
            UPDATE store_daily_state
            SET is_visible = %s, is_checked_in = %s, updated_at = %s
            WHERE date = %s AND store_id = %s
        """, updates)
        
        active_count = len(DEFAULT_STORES) - len(hidden_ids)
        conn.execute("""
            INSERT INTO version_history (date, action, version_name, is_named, active_count, closed_count, hidden_ids, checked_ids, created_at)
            VALUES (%s, %s, %s, %s, %s, %s, %s, %s, %s)
        """, (target_date, f"↩️ 回復昨日最後紀錄 ({yesterday})", "", False, active_count, len(hidden_ids), last_hist["hidden_ids"] if last_hist else "", "", now_str))
        
        sync_time = update_sync_time(conn, target_date)
        return {
            "success": True,
            "yesterdayDate": yesterday,
            "activeCount": active_count,
            "closedCount": len(hidden_ids),
            "serverTime": sync_time
        }

@app.get("/api/versions")
def get_versions(date: Optional[str] = None):
    target_date = date or get_today_str()
    with get_db() as conn:
        rows = conn.execute("""
            SELECT * FROM version_history
            ORDER BY id DESC
            LIMIT 40
        """).fetchall()
        
        all_list = []
        named_list = []
        
        for r in rows:
            is_named = bool(r["is_named"]) or bool(r["version_name"])
            item = {
                "id": r["id"],
                "date": r["date"],
                "action": r["action"],
                "versionName": r["version_name"] or r["action"],
                "isNamed": is_named,
                "activeCount": r["active_count"],
                "closedCount": r["closed_count"],
                "timestamp": r["created_at"]
            }
            all_list.append(item)
            if is_named:
                named_list.append(item)
                
        return {
            "success": True,
            "list": all_list,
            "namedList": named_list
        }

@app.post("/api/versions/save")
def save_named_version(req: SaveNamedVersionRequest):
    with get_db() as conn:
        now_str = get_now_str()
        hidden_str = ",".join(map(str, req.hidden_ids))
        checked_str = ",".join(map(str, req.checked_ids))
        active_count = len(DEFAULT_STORES) - len(req.hidden_ids)
        closed_count = len(req.hidden_ids)
        
        version_id = conn.execute("""
            INSERT INTO version_history (date, action, version_name, is_named, active_count, closed_count, hidden_ids, checked_ids, created_at)
            VALUES (%s, %s, %s, %s, %s, %s, %s, %s, %s)
            RETURNING id
        """, (req.date, f"🏷️ 自訂版本：{req.version_name}", req.version_name, True, active_count, closed_count, hidden_str, checked_str, now_str)).fetchone()["id"]
        
        sync_time = update_sync_time(conn, req.date)
        return {
            "success": True,
            "versionId": version_id,
            "versionName": req.version_name,
            "activeCount": active_count,
            "closedCount": closed_count,
            "serverTime": sync_time
        }

@app.post("/api/versions/revert")
def revert_version(req: RevertVersionRequest):
    with get_db() as conn:
        hist = conn.execute("SELECT * FROM version_history WHERE id = %s", (req.version_id,)).fetchone()
        if not hist:
            raise HTTPException(status_code=404, detail="找不到指定的歷史版本！")
            
        hidden_ids = [int(x) for x in hist["hidden_ids"].split(",") if x.strip()]
        checked_ids = [int(x) for x in hist["checked_ids"].split(",") if x.strip()]
        hidden_set = set(hidden_ids)
        checked_set = set(checked_ids)
        now_str = get_now_str()
        
        updates = []
        for i in range(len(DEFAULT_STORES)):
            s_id = i + 1
            vis = s_id not in hidden_set
            chk = s_id in checked_set
            updates.append((vis, chk, now_str, req.date, s_id))
            
        conn.cursor().executemany("""
            UPDATE store_daily_state
            SET is_visible = %s, is_checked_in = %s, updated_at = %s
            WHERE date = %s AND store_id = %s
        """, updates)
        
        active_count = len(DEFAULT_STORES) - len(hidden_ids)
        v_title = hist["version_name"] or hist["action"]
        conn.execute("""
            INSERT INTO version_history (date, action, version_name, is_named, active_count, closed_count, hidden_ids, checked_ids, created_at)
            VALUES (%s, %s, %s, %s, %s, %s, %s, %s, %s)
        """, (req.date, f"讀取版本 [{v_title}]", "", False, active_count, len(hidden_ids), hist["hidden_ids"], hist["checked_ids"], now_str))
        
        sync_time = update_sync_time(conn, req.date)
        return {
            "success": True,
            "versionName": v_title,
            "activeCount": active_count,
            "closedCount": len(hidden_ids),
            "serverTime": sync_time
        }

# ==========================================
# 📊 一鍵匯出標準 Excel (.xlsx) 功能
# ==========================================
@app.get("/api/export-excel")
def export_excel(date: Optional[str] = None):
    target_date = date or get_today_str()
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = f"排程報表_{target_date}"

    # 表頭樣式
    header_fill = PatternFill(start_color="C5221F", end_color="C5221F", fill_type="solid")
    header_font = Font(name="微軟正黑體", size=11, bold=True, color="FFFFFF")
    data_font = Font(name="微軟正黑體", size=10)
    align_center = Alignment(horizontal="center", vertical="center")
    align_left = Alignment(horizontal="left", vertical="center")
    thin_border = Border(
        left=Side(style='thin', color='DDDDDD'),
        right=Side(style='thin', color='DDDDDD'),
        top=Side(style='thin', color='DDDDDD'),
        bottom=Side(style='thin', color='DDDDDD')
    )

    headers = ["序號", "顯示狀態", "店名", "地址", "分組分類", "區域", "經緯度坐標", "打卡狀態", "打卡時間"]
    ws.append(headers)

    for col_idx in range(1, len(headers) + 1):
        cell = ws.cell(row=1, column=col_idx)
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = align_center

    with get_db() as conn:
        ensure_date_initialized(conn, target_date)
        rows = conn.execute("""
            SELECT store_id, is_visible, is_checked_in, checkin_time
            FROM store_daily_state
            WHERE date = %s
            ORDER BY store_id ASC
        """, (target_date,)).fetchall()
        state_map = {r["store_id"]: r for r in rows}

    for idx, s in enumerate(DEFAULT_STORES):
        num = idx + 1
        st = state_map.get(num)
        is_vis = bool(st["is_visible"]) if st else True
        is_chk = bool(st["is_checked_in"]) if st else False
        chk_time = st["checkin_time"] if st else ""

        row_data = [
            num,
            "開啟顯示" if is_vis else "隱藏關閉",
            s["name"],
            s["address"],
            s["subGroup"],
            s["region"],
            f"{s['lat']}, {s['lng']}",
            "✅ 已打卡" if is_chk else "未打卡",
            chk_time
        ]
        ws.append(row_data)
        cur_row = idx + 2
        
        # 標註背景顏色
        status_fill = PatternFill(start_color="E6F4EA" if is_vis else "FCE8E6", fill_type="solid")
        ws.cell(row=cur_row, column=2).fill = status_fill

        for col_idx in range(1, len(headers) + 1):
            c = ws.cell(row=cur_row, column=col_idx)
            c.font = data_font
            c.border = thin_border
            if col_idx in (1, 2, 6, 8, 9):
                c.alignment = align_center
            else:
                c.alignment = align_left

    # 自動調整欄寬
    col_widths = [8, 12, 26, 34, 18, 12, 22, 12, 12]
    for i, w in enumerate(col_widths, 1):
        ws.column_dimensions[openpyxl.utils.get_column_letter(i)].width = w

    output = BytesIO()
    wb.save(output)
    output.seek(0)

    filename = f"hualien_stores_{target_date}.xlsx"
    return StreamingResponse(
        output,
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers={"Content-Disposition": f"attachment; filename={filename}"}
    )

# 掛載靜態文件
STATIC_DIR = os.path.join(BASE_DIR, "static")
os.makedirs(STATIC_DIR, exist_ok=True)
app.mount("/static", StaticFiles(directory=STATIC_DIR), name="static")

@app.get("/", response_class=HTMLResponse)
def index_page():
    index_path = os.path.join(STATIC_DIR, "index.html")
    if os.path.exists(index_path):
        with open(index_path, "r", encoding="utf-8") as f:
            return f.read()
    return "<h1>花蓮店家排程地圖建置中...</h1>"

if __name__ == "__main__":
    import uvicorn
    uvicorn.run(app, host="0.0.0.0", port=int(os.getenv("PORT", "8000")))
