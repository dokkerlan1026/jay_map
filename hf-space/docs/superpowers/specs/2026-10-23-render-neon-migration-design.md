# Render + Neon 遷移設計

## 目標

將目前以 Hugging Face Gradio SDK 與本機 SQLite 執行的花蓮店家排程網站，遷移為由 Render 執行的原生 FastAPI 應用程式，並使用 Neon PostgreSQL 永久保存多人排程、打卡與版本紀錄。Neon 從全新空資料庫開始，不遷移既有 SQLite 資料。

## 範圍

- 保留現有網頁畫面、REST API 路徑與回應格式。
- 移除 Gradio 掛載層與 Hugging Face 專用 metadata／相依套件。
- 以 `psycopg` 直接存取 Neon PostgreSQL，不導入 ORM。
- 新增 Render Blueprint 與健康檢查端點。
- 透過 GitHub `main` 分支觸發 Render 自動部署。
- 不包含舊 SQLite 資料搬移、使用者登入或權限系統。

## 架構

Render 執行原生 FastAPI 應用程式，由 FastAPI 提供 `static/index.html`、靜態檔案與既有 API。Render Secret 環境變數 `DATABASE_URL` 存放 Neon 連線字串。新增 `database.py` 封裝 PostgreSQL 連線、交易、資料表初始化與查詢結果處理；`app.py` 保留 HTTP 與業務邏輯，透過該模組存取資料。

應用程式啟動時驗證 `DATABASE_URL` 並建立以下資料表：

- `store_daily_state`：每日期、每店家的顯示及打卡狀態。
- `version_history`：自動與具名版本歷程。
- `sync_meta`：每日期的最後同步時間。

第一次讀取某日期時，系統原子化建立該日期的 500 筆預設狀態與初始版本。並行請求使用 PostgreSQL conflict handling，避免重複初始化。

## 資料流程

1. 使用者開啟 Render 網址，FastAPI 回傳現有前端。
2. 前端呼叫現有 `/api/stores`、版本、切換與 Excel 匯出端點。
3. API 透過 `database.py` 建立短生命週期 PostgreSQL 連線。
4. 寫入操作在 transaction 中完成；成功時提交，例外時回滾。
5. Render 重啟或重新部署後仍連回同一個 Neon 資料庫。
6. GitHub `main` 分支有新 commit 時，Render 自動建置與部署。

## PostgreSQL 相容調整

- SQLite `?` placeholder 改為 psycopg `%s` placeholder。
- `INSERT OR REPLACE` 改為 `INSERT ... ON CONFLICT ... DO UPDATE`。
- 自動遞增主鍵改為 PostgreSQL identity column。
- 查詢結果使用 dict row factory，維持現有 `row["column"]` 存取方式。
- 布林欄位保留目前 API 語意，資料庫內使用 PostgreSQL boolean。
- 具名版本新增後透過 `RETURNING id` 取得版本 ID。

## 錯誤處理與安全

- `DATABASE_URL` 僅設定於 Render Secret，不寫入 repository、設定檔或日誌。
- 所有 SQL 使用參數化查詢。
- 缺少或無法連接 `DATABASE_URL` 時啟動失敗並輸出明確原因，不退回本機 SQLite。
- 資料庫暫時不可用時，API 回傳 HTTP 503；交易自動回滾。
- 新增 `/health`，實際執行輕量資料庫查詢；正常時回傳 200，異常時回傳 503。

## 部署設定

- `requirements.txt` 移除 Gradio 與 Hugging Face Hub，新增 `psycopg[binary]` 與測試依賴。
- `render.yaml` 定義 Python Web Service、安裝命令、啟動命令與 `/health` 健康檢查。
- 啟動命令使用 Render 提供的 `$PORT`：`uvicorn app:app --host 0.0.0.0 --port $PORT`。
- `README.md` 改為 Render + Neon 設定說明，包含 Neon 建立、`DATABASE_URL` Secret 與 GitHub／Render 串接步驟。

## 測試策略

- 資料庫單元／整合測試覆蓋建表、日期初始化、同步時間 upsert 與版本 ID 回傳。
- API 測試覆蓋首頁、健康檢查、店家讀取、顯示切換、打卡與版本還原。
- 測試缺少 `DATABASE_URL` 時的明確失敗行為。
- 執行 Python 編譯、完整 pytest，以及與 Render 相同的 Uvicorn 啟動煙霧測試。

## 完成條件

- 專案不再匯入或依賴 Gradio、Hugging Face Hub 或 SQLite。
- 所有現有前端功能與 API 回應格式維持可用。
- 新 Neon 資料庫可自動建表並初始化指定日期。
- Render 重啟後資料仍存在 Neon。
- Secrets 不進入 Git，所有測試與啟動驗證通過。
