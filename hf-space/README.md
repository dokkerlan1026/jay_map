# 花蓮店家每日排程地圖 (Render + Neon 版)

專為花蓮 500 筆店家設計的高效能每日排程、地圖定位、打卡與多人即時共同編輯系統。
後端為 FastAPI，資料存放於 Neon PostgreSQL，部署於 Render。

### 🌟 特色功能
- 📍 **500 筆店家精確定位**：店名真實 POI 地標對齊、導航與搜尋一鍵直達。
- 🟢🔴🟡 **三區彩色標籤**：東區（綠）、西區（紅）、非488名單（黃），支援「紅綠同時顯示」與快速切換。
- 👥 **全員即時共同編輯**：多人同時在線打卡、開關，秒級自動同步，不再互踩覆蓋。
- 💾 **自訂具名版本與版本歷程**：支援儲存命名版本（如「下午路線」），全體同仁一鍵載入共享。
- ☀️ **隔天全開啟 / ↩️ 回到昨天**：一鍵快速切換排程。
- 📊 **直接下載 Excel (.xlsx)**：隨時下載當前所有店家的開關、地址與打卡時間報表。

## 🚀 部署步驟

### 1. 建立 Neon 資料庫
1. 到 <https://neon.tech> 註冊並建立新 Project（Region 建議選 Singapore）。
2. 在 Project Dashboard 點 **Connect**，複製 **Pooled connection** 的連線字串，格式如：
   `postgresql://<user>:<password>@<host>-pooler.<region>.aws.neon.tech/<db>?sslmode=require`
3. 資料表會在服務第一次啟動時自動建立，不需手動執行 SQL。

### 2. 推送程式碼到 GitHub
```bash
git push origin main
```

### 3. 在 Render 建立 Blueprint
1. 到 <https://dashboard.render.com> → **New** → **Blueprint**，選擇此 GitHub repository。
2. Render 會讀取 repository 根目錄的 `render.yaml`（服務根目錄為 `hf-space`）。
3. 提示輸入 `DATABASE_URL` 時，貼上步驟 1 複製的 Neon 連線字串（屬於機密，切勿提交到 Git）。
4. 點 **Apply** 開始第一次部署。

### 4. 驗證部署
- 開啟 `https://<你的服務>.onrender.com/health`，應回傳 `{"status": "ok"}`。
- 開啟 `https://<你的服務>.onrender.com/` 應顯示地圖。

## 🧪 本機開發與測試
```bash
docker run -d --name hualien-neon-test -e POSTGRES_PASSWORD=postgres -e POSTGRES_DB=hualien_test -p 55432:5432 postgres:16-alpine
```
```bash
pip install -r requirements.txt
```
PowerShell：
```powershell
$env:DATABASE_URL='postgresql://postgres:postgres@localhost:55432/hualien_test'; pytest -v
```
啟動伺服器：`uvicorn app:app --reload`
