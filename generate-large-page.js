import fs from "node:fs";
import path from "node:path";
import { fileURLToPath } from "node:url";

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);

const jsonPath = path.join(__dirname, "large_stores.json");
const stores = JSON.parse(fs.readFileSync(jsonPath, "utf8"));

const eastStores = stores.filter(s => s.subGroup === "主要名單-東區");
const westStores = stores.filter(s => s.subGroup === "主要名單-西區");
const diffStores = stores.filter(s => s.subGroup === "非488名單");

const html = `<!doctype html>
<html lang="zh-Hant">
  <head>
    <meta charset="UTF-8" />
    <meta name="viewport" content="width=device-width, initial-scale=1.0" />
    <meta name="description" content="花蓮店家地圖 - Google 我的地圖風格，支援紅綠同時顯示、黃色非488標記、單點勾選與店家打卡功能。" />
    <title>jay的地圖 - 花蓮店家地圖 (紅綠雙色+黃色非488 & 一鍵打卡)</title>
    <link rel="preconnect" href="https://fonts.googleapis.com" />
    <link rel="preconnect" href="https://fonts.gstatic.com" crossorigin />
    <link href="https://fonts.googleapis.com/css2?family=Noto+Sans+TC:wght@400;500;600;700;800&family=Roboto:wght@400;500;700&display=swap" rel="stylesheet" />
    <link rel="stylesheet" href="https://unpkg.com/leaflet@1.9.4/dist/leaflet.css" />
    <style>
      :root {
        --google-red: #c5221f;
        --google-hover: #f1f3f4;
        --color-east: #2e7d32;       /* 綠色 */
        --color-west: #d32f2f;       /* 紅色 */
        --color-diff: #f59e0b;       /* 黃色/金黃色 */
        --border-color: #dadce0;
        --text-primary: #202124;
        --text-secondary: #5f6368;
        --sidebar-width: 400px;
      }
      * { box-sizing: border-box; margin: 0; padding: 0; }
      body {
        font-family: "Noto Sans TC", "Roboto", system-ui, sans-serif;
        background: #f8f9fa;
        color: var(--text-primary);
        height: 100vh;
        display: flex;
        flex-direction: column;
        overflow: hidden;
      }

      /* 頂部 Google 我的地圖經典紅色工具列 */
      header {
        background: var(--google-red);
        color: white;
        height: 56px;
        padding: 0 16px;
        display: flex;
        align-items: center;
        justify-content: space-between;
        gap: 12px;
        box-shadow: 0 2px 4px rgba(0,0,0,0.2);
        z-index: 1000;
        flex-shrink: 0;
      }
      .header-left {
        display: flex;
        align-items: center;
        gap: 12px;
      }
      .menu-btn {
        background: none;
        border: none;
        color: white;
        width: 36px;
        height: 36px;
        border-radius: 50%;
        display: flex;
        align-items: center;
        justify-content: center;
        cursor: pointer;
        transition: background 150ms;
      }
      .menu-btn:hover { background: rgba(255,255,255,0.18); }
      .brand-box {
        display: flex;
        flex-direction: column;
      }
      .brand-title {
        font-size: 1.15rem;
        font-weight: 700;
        letter-spacing: 0.3px;
        display: flex;
        align-items: center;
        gap: 8px;
      }
      .brand-sub {
        font-size: 0.75rem;
        opacity: 0.88;
      }
      .header-right {
        display: flex;
        align-items: center;
        gap: 8px;
      }
      .header-btn {
        background: rgba(255,255,255,0.18);
        color: white;
        border: 1px solid rgba(255,255,255,0.3);
        padding: 6px 10px;
        border-radius: 4px;
        font-size: 0.78rem;
        font-weight: 600;
        text-decoration: none;
        display: inline-flex;
        align-items: center;
        gap: 4px;
        cursor: pointer;
        transition: all 140ms ease;
      }
      .header-btn:hover {
        background: white;
        color: var(--google-red);
      }

      /* 主版面：左側欄 + 右側地圖 */
      .layout-body {
        flex: 1;
        display: flex;
        position: relative;
        overflow: hidden;
      }

      /* 左側抽屜 / 面板 */
      .sidebar {
        width: var(--sidebar-width);
        background: white;
        border-right: 1px solid var(--border-color);
        display: flex;
        flex-direction: column;
        height: 100%;
        z-index: 500;
        box-shadow: 2px 0 6px rgba(0,0,0,0.06);
        transition: transform 240ms ease;
      }
      .sidebar.collapsed {
        transform: translateX(calc(-1 * var(--sidebar-width)));
      }

      /* 統計與快捷按鈕區 */
      .panel-intro {
        padding: 10px 16px;
        border-bottom: 1px solid var(--border-color);
        background: #fafafa;
      }
      .intro-meta {
        font-size: 0.78rem;
        color: var(--text-secondary);
        display: flex;
        align-items: center;
        justify-content: space-between;
        margin-bottom: 8px;
      }
      .quick-actions {
        display: flex;
        gap: 4px;
        flex-wrap: wrap;
      }
      .btn-chip {
        padding: 4px 8px;
        border: 1px solid var(--border-color);
        background: white;
        border-radius: 12px;
        font-size: 0.72rem;
        font-weight: 600;
        color: var(--text-secondary);
        cursor: pointer;
        transition: all 120ms;
        display: inline-flex;
        align-items: center;
        gap: 4px;
      }
      .btn-chip:hover {
        background: var(--google-hover);
        color: var(--text-primary);
        border-color: #bdc1c6;
      }
      .btn-chip.highlight-redgreen {
        background: #e8f5e9;
        color: #1b5e20;
        border-color: #a5d6a7;
        font-weight: 700;
      }
      .btn-chip.highlight-redgreen:hover {
        background: #c8e6c9;
      }

      /* 搜尋列 */
      .search-box {
        padding: 8px 16px;
        border-bottom: 1px solid var(--border-color);
        background: #ffffff;
      }
      .search-input-wrap {
        position: relative;
        display: flex;
        align-items: center;
      }
      .search-input {
        width: 100%;
        padding: 7px 30px 7px 10px;
        border: 1px solid var(--border-color);
        border-radius: 6px;
        font-size: 0.86rem;
        font-family: inherit;
        outline: none;
      }
      .search-input:focus { border-color: #1a73e8; }
      .search-clear {
        position: absolute;
        right: 8px;
        background: none;
        border: none;
        font-size: 16px;
        color: #9aa0a6;
        cursor: pointer;
        display: none;
      }

      /* 圖層清單（樹狀結構） */
      .layer-tree {
        flex: 1;
        overflow-y: auto;
      }

      /* 單一圖層群組 */
      .layer-group {
        border-bottom: 1px solid var(--border-color);
      }
      .group-header {
        display: flex;
        align-items: center;
        padding: 10px 12px 10px 16px;
        background: #ffffff;
        cursor: pointer;
        user-select: none;
        transition: background 120ms;
        gap: 10px;
      }
      .group-header:hover {
        background: #f1f3f4;
      }
      .group-checkbox {
        width: 18px;
        height: 18px;
        cursor: pointer;
        accent-color: var(--google-red);
      }
      .group-icon {
        width: 12px;
        height: 12px;
        border-radius: 50%;
        flex-shrink: 0;
      }
      .group-icon.east { background: var(--color-east); }
      .group-icon.west { background: var(--color-west); }
      .group-icon.diff { background: var(--color-diff); }

      .group-title-wrap {
        flex: 1;
        display: flex;
        flex-direction: column;
      }
      .group-title {
        font-size: 0.9rem;
        font-weight: 700;
        color: var(--text-primary);
        display: flex;
        align-items: center;
        gap: 6px;
      }
      .group-stats {
        font-size: 0.72rem;
        color: var(--text-secondary);
      }
      .arrow-btn {
        width: 24px;
        height: 24px;
        display: flex;
        align-items: center;
        justify-content: center;
        font-size: 0.8rem;
        color: var(--text-secondary);
        transition: transform 180ms ease;
      }
      .layer-group.expanded .arrow-btn {
        transform: rotate(180deg);
      }

      .item-list {
        display: none;
        background: #ffffff;
      }
      .layer-group.expanded .item-list {
        display: block;
      }

      /* 單一店家項目（支援打卡按鈕與獨立勾選） */
      .store-row {
        display: flex;
        align-items: center;
        padding: 8px 12px 8px 16px;
        border-top: 1px solid #f1f3f4;
        cursor: pointer;
        transition: background 100ms;
        gap: 8px;
      }
      .store-row:hover {
        background: #f8f9fa;
      }
      .store-row.active {
        background: #e8f0fe;
      }
      .store-row.disabled {
        opacity: 0.42;
      }
      .item-checkbox {
        width: 16px;
        height: 16px;
        cursor: pointer;
        accent-color: var(--google-red);
        flex-shrink: 0;
      }
      .pin-svg {
        width: 14px;
        height: 18px;
        flex-shrink: 0;
      }
      .store-info {
        flex: 1;
        min-width: 0;
      }
      .store-name-text {
        font-size: 0.86rem;
        font-weight: 500;
        color: var(--text-primary);
        white-space: nowrap;
        overflow: hidden;
        text-overflow: ellipsis;
      }
      .store-addr-text {
        font-size: 0.72rem;
        color: var(--text-secondary);
        white-space: nowrap;
        overflow: hidden;
        text-overflow: ellipsis;
      }

      /* 打卡按鈕（截圖手繪標註樣式） */
      .checkin-btn {
        padding: 4px 8px;
        font-size: 0.72rem;
        font-weight: 600;
        border-radius: 4px;
        border: 1px solid #d1d5db;
        background: #ffffff;
        color: #4b5563;
        cursor: pointer;
        flex-shrink: 0;
        transition: all 140ms ease;
        display: inline-flex;
        align-items: center;
        gap: 3px;
        user-select: none;
      }
      .checkin-btn:hover {
        background: #f3f4f6;
        border-color: #9ca3af;
      }
      /* 已打卡樣式 */
      .checkin-btn.checked-in {
        background: #16a34a;
        color: #ffffff;
        border-color: #15803d;
        box-shadow: 0 1px 3px rgba(22,163,74,0.3);
      }
      .checkin-btn.checked-in:hover {
        background: #15803d;
      }

      /* 側邊欄收合切換浮動按鈕 */
      .toggle-sidebar-btn {
        position: absolute;
        top: 14px;
        left: calc(var(--sidebar-width) + 12px);
        z-index: 999;
        background: white;
        border: 1px solid var(--border-color);
        box-shadow: 0 1px 4px rgba(0,0,0,0.18);
        border-radius: 4px;
        padding: 6px 10px;
        font-size: 0.8rem;
        font-weight: 600;
        cursor: pointer;
        display: flex;
        align-items: center;
        gap: 6px;
        transition: all 200ms ease;
      }
      .sidebar.collapsed ~ .toggle-sidebar-btn {
        left: 14px;
      }

      /* 右側 Leaflet 地圖容器 */
      #map {
        flex: 1;
        height: 100%;
        background: #f2efe9;
      }

      /* 自訂 Google 經典水滴圖釘 Marker */
      .gmarker-wrap {
        background: transparent;
        border: none;
      }
      .gmarker-svg {
        filter: drop-shadow(0 2px 3px rgba(0,0,0,0.35));
        transition: transform 120ms ease;
      }
      .gmarker-svg:hover {
        transform: scale(1.2) translateY(-2px);
      }

      /* Popup 美化 */
      .leaflet-popup-content-wrapper {
        border-radius: 8px;
        box-shadow: 0 2px 10px rgba(0,0,0,0.22);
        padding: 0;
        overflow: hidden;
      }
      .leaflet-popup-content {
        margin: 0;
        padding: 14px 16px;
        min-width: 250px;
        max-width: 320px;
        font-family: inherit;
      }
      .pop-header {
        display: flex;
        align-items: center;
        justify-content: space-between;
        margin-bottom: 6px;
      }
      .pop-badge {
        font-size: 0.7rem;
        font-weight: 700;
        padding: 2px 8px;
        border-radius: 12px;
      }
      .pop-badge.east { background: #e8f5e9; color: var(--color-east); }
      .pop-badge.west { background: #ffebee; color: var(--color-west); }
      .pop-badge.diff { background: #fef3c7; color: #b45309; }

      .pop-title {
        font-size: 1rem;
        font-weight: 700;
        color: var(--text-primary);
        margin-bottom: 4px;
      }
      .pop-addr {
        font-size: 0.8rem;
        color: var(--text-secondary);
        margin-bottom: 12px;
        line-height: 1.4;
      }
      .pop-actions {
        display: flex;
        gap: 6px;
        flex-wrap: wrap;
      }
      .pop-btn {
        padding: 6px 10px;
        border-radius: 4px;
        font-size: 0.78rem;
        font-weight: 600;
        cursor: pointer;
        text-decoration: none;
        display: inline-flex;
        align-items: center;
        gap: 4px;
      }
      .pop-btn-nav {
        background: #1a73e8;
        color: white !important;
        border: none;
        flex: 1;
        justify-content: center;
      }
      .pop-btn-nav:hover { background: #1557b0; }
      .pop-btn-checkin {
        background: #ffffff;
        border: 1px solid #d1d5db;
        color: #374151;
      }
      .pop-btn-checkin.checked-in {
        background: #16a34a;
        color: white;
        border-color: #15803d;
      }

      /* 打卡通知 Toast */
      .toast {
        position: fixed;
        bottom: 24px;
        left: 50%;
        transform: translateX(-50%);
        background: #1f2937;
        color: #ffffff;
        padding: 10px 18px;
        border-radius: 20px;
        font-size: 0.84rem;
        font-weight: 500;
        z-index: 2000;
        box-shadow: 0 4px 12px rgba(0,0,0,0.25);
        display: none;
        animation: fadeIn 150ms ease-in-out;
      }
      @keyframes fadeIn {
        from { opacity: 0; transform: translate(-50%, 10px); }
        to { opacity: 1; transform: translate(-50%, 0); }
      }

      @media (max-width: 768px) {
        .layout-body { flex-direction: column-reverse; }
        .sidebar { width: 100%; height: 50vh; }
        #map { width: 100%; height: 50vh; }
        .toggle-sidebar-btn { display: none; }
      }
    </style>
  </head>
  <body>
    <header>
      <div class="header-left">
        <button class="menu-btn" id="menuToggleBtn" title="展開/收合側邊欄">
          <svg width="22" height="22" viewBox="0 0 24 24" fill="currentColor">
            <path d="M3 18h18v-2H3v2zm0-5h18v-2H3v2zm0-7v2h18V6H3z"/>
          </svg>
        </button>
        <div class="brand-box">
          <div class="brand-title">
            <span>jay的地圖</span>
            <span style="font-size:0.75rem;background:rgba(255,255,255,0.22);padding:2px 8px;border-radius:12px;">花蓮店家地圖</span>
          </div>
          <div class="brand-sub">主要名單拆為東區 (綠) / 西區 (紅) 與非488名單 (黃)</div>
        </div>
      </div>
      <div class="header-right">
        <a href="./hualien_large_stores.kml" download="hualien_large_stores.kml" class="header-btn" title="下載完整 KML">
          <span>↓ KML</span>
        </a>
        <a href="./hualien_google_mymaps.csv" download="hualien_google_mymaps.csv" class="header-btn" title="下載 Google 地圖專用 CSV">
          <span>↓ CSV</span>
        </a>
        <a href="./index.html" class="header-btn" title="示範路線規劃版">
          <span>← 路線規劃</span>
        </a>
      </div>
    </header>

    <div class="layout-body">
      <aside class="sidebar" id="sidebar">
        <!-- 頂部資訊與快速操作按鈕 -->
        <div class="panel-intro">
          <div class="intro-meta">
            <span id="activeCountLabel">目前顯示：500 / 500 處店家</span>
            <span id="checkinStatsLabel" style="color:#16a34a;font-weight:700;">已打卡：0 家</span>
          </div>
          <div class="quick-actions">
            <button class="btn-chip" id="btnShowAll">全部顯示</button>
            <button class="btn-chip" id="btnHideAll">全部隱藏</button>
            <!-- 🌟 使用者指定新增：紅綠同時顯示選項 -->
            <button class="btn-chip highlight-redgreen" id="btnShowRedGreen" title="僅顯示東區(綠)與西區(紅)，隱藏非488名單">
              <span>🟢🔴 紅綠同時顯示</span>
            </button>
            <button class="btn-chip" id="btnOnlyEast">僅東區 (綠)</button>
            <button class="btn-chip" id="btnOnlyWest">僅西區 (紅)</button>
            <button class="btn-chip" id="btnOnlyDiff" style="color:#b45309;">僅非488 (黃)</button>
          </div>
        </div>

        <!-- 搜尋輸入框 -->
        <div class="search-box">
          <div class="search-input-wrap">
            <input type="text" id="searchInput" class="search-input" placeholder="搜尋店名、路名、地址…" />
            <button id="searchClearBtn" class="search-clear">×</button>
          </div>
        </div>

        <!-- 圖層清單 -->
        <div class="layer-tree">
          <!-- 1. 東區主要名單 (綠色) -->
          <div class="layer-group expanded" id="group-east">
            <div class="group-header">
              <input type="checkbox" class="group-checkbox" id="chk-group-east" checked title="全選/取消東區" />
              <div class="group-icon east"></div>
              <div class="group-title-wrap" onclick="toggleGroupExpand('group-east')">
                <span class="group-title">🟢 主要名單-東區 (界線以東)</span>
                <span class="group-stats" id="stats-east">顯示 ${eastStores.length} / ${eastStores.length} 筆</span>
              </div>
              <div class="arrow-btn" onclick="toggleGroupExpand('group-east')">▼</div>
            </div>
            <div class="item-list" id="list-east"></div>
          </div>

          <!-- 2. 西區主要名單 (紅色) -->
          <div class="layer-group expanded" id="group-west">
            <div class="group-header">
              <input type="checkbox" class="group-checkbox" id="chk-group-west" checked title="全選/取消西區" />
              <div class="group-icon west"></div>
              <div class="group-title-wrap" onclick="toggleGroupExpand('group-west')">
                <span class="group-title">🔴 主要名單-西區 (界線以西)</span>
                <span class="group-stats" id="stats-west">顯示 ${westStores.length} / ${westStores.length} 筆</span>
              </div>
              <div class="arrow-btn" onclick="toggleGroupExpand('group-west')">▼</div>
            </div>
            <div class="item-list" id="list-west"></div>
          </div>

          <!-- 3. 非 488 店家名單 (黃色) -->
          <div class="layer-group expanded" id="group-diff">
            <div class="group-header">
              <input type="checkbox" class="group-checkbox" id="chk-group-diff" checked title="全選/取消非488名單" />
              <div class="group-icon diff"></div>
              <div class="group-title-wrap" onclick="toggleGroupExpand('group-diff')">
                <span class="group-title">🟡 非488名單 (黃色標記)</span>
                <span class="group-stats" id="stats-diff">顯示 ${diffStores.length} / ${diffStores.length} 筆</span>
              </div>
              <div class="arrow-btn" onclick="toggleGroupExpand('group-diff')">▼</div>
            </div>
            <div class="item-list" id="list-diff"></div>
          </div>
        </div>
      </aside>

      <!-- 側邊欄浮動收合切換鈕 -->
      <button class="toggle-sidebar-btn" id="floatingToggleBtn">
        <span id="toggleIcon">◀</span>
        <span id="toggleText">收合清單</span>
      </button>

      <div id="map"></div>
    </div>

    <!-- 打卡提示 Toast -->
    <div id="toastBox" class="toast">已完成打卡！</div>

    <script src="https://unpkg.com/leaflet@1.9.4/dist/leaflet.js"></script>
    <script>
      const stores = ${JSON.stringify(stores)};

      // 初始化 Leaflet 地圖
      const map = L.map('map', { zoomControl: true }).setView([23.985, 121.605], 13);
      L.tileLayer('https://{s}.tile.openstreetmap.org/{z}/{x}/{y}.png', {
        maxZoom: 19,
        attribution: '&copy; OpenStreetMap & Google My Maps Style'
      }).addTo(map);

      // 林森路 + 德安一街分界輔助虛線
      const boundaryLine = L.polyline([
        [23.960, 121.581],
        [23.970, 121.5925],
        [23.983, 121.6070],
        [23.995, 121.620]
      ], {
        color: '#6b7280',
        weight: 2,
        dashArray: '6, 6',
        opacity: 0.6
      }).addTo(map);
      boundaryLine.bindTooltip('← 界線以西 (西區紅) ｜ 林森路+德安一街分界 ｜ 界線以東 (東區綠) →', { sticky: true });

      // 水滴型 SVG Marker 生成函式
      function createPinIcon(colorHex) {
        const svg = \`
          <svg class="gmarker-svg" width="22" height="30" viewBox="0 0 24 34" fill="none" xmlns="http://www.w3.org/2000/svg">
            <path d="M12 0C5.372 0 0 5.373 0 12C0 19.5 12 34 12 34C12 34 24 19.5 24 12C24 5.373 18.628 0 12 0Z" fill="\${colorHex}" stroke="#ffffff" stroke-width="1.5"/>
            <circle cx="12" cy="11" r="4.5" fill="#ffffff"/>
          </svg>
        \`;
        return L.divIcon({
          className: 'gmarker-wrap',
          html: svg,
          iconSize: [22, 30],
          iconAnchor: [11, 30],
          popupAnchor: [0, -28]
        });
      }

      const iconEast = createPinIcon('#2e7d32'); // 綠色
      const iconWest = createPinIcon('#d32f2f'); // 紅色
      const iconDiff = createPinIcon('#f59e0b'); // 黃色

      function getSmallPinSvg(colorHex) {
        return \`<svg class="pin-svg" viewBox="0 0 24 34" fill="none"><path d="M12 0C5.372 0 0 5.373 0 12C0 19.5 12 34 12 34C12 34 24 19.5 24 12C24 5.373 18.628 0 12 0Z" fill="\${colorHex}" stroke="#ffffff" stroke-width="1.5"/><circle cx="12" cy="11" r="4" fill="#ffffff"/></svg>\`;
      }

      // 打卡持久化儲存 (localStorage)
      const CHECKIN_STORAGE_KEY = 'hualien_store_checkins';
      let checkinSet = new Set();
      try {
        const saved = localStorage.getItem(CHECKIN_STORAGE_KEY);
        if (saved) {
          checkinSet = new Set(JSON.parse(saved));
        }
      } catch (e) {
        console.warn('LocalStorage unavailable', e);
      }

      function saveCheckins() {
        try {
          localStorage.setItem(CHECKIN_STORAGE_KEY, JSON.stringify(Array.from(checkinSet)));
        } catch (e) {}
        updateCheckinStats();
      }

      function showToast(text) {
        const toast = document.getElementById('toastBox');
        toast.textContent = text;
        toast.style.display = 'block';
        setTimeout(() => { toast.style.display = 'none'; }, 2000);
      }

      const storeStates = new Map();

      function getGoogleMapsNavUrl(s) {
        return 'https://www.google.com/maps/dir/?api=1&destination=' + s.lat + ',' + s.lng;
      }

      // 建立地圖 Marker 與彈出視窗
      stores.forEach((store, index) => {
        let icon = iconEast;
        let badgeClass = 'east';
        let badgeText = '🟢 主要名單 - 界線以東';
        if (store.subGroup === '主要名單-西區') {
          icon = iconWest;
          badgeClass = 'west';
          badgeText = '🔴 主要名單 - 界線以西';
        } else if (store.subGroup === '非488名單') {
          icon = iconDiff;
          badgeClass = 'diff';
          badgeText = '🟡 非 488 名單店家';
        }

        const marker = L.marker([store.lat, store.lng], { icon: icon });

        function updatePopupContent() {
          const isCheckedIn = checkinSet.has(store.id);
          const popupDiv = document.createElement('div');
          popupDiv.innerHTML = \`
            <div class="pop-header">
              <span class="pop-badge \${badgeClass}">\${badgeText}</span>
              <span style="font-size:0.75rem;color:#888;">#\${index + 1}</span>
            </div>
            <div class="pop-title">\${store.name}</div>
            <div class="pop-addr">\${store.address}\${store.estimated ? '<br/><span style="color:#b45309;font-size:11px;">※ 位置為地址估算</span>' : ''}</div>
            <div class="pop-actions">
              <a class="pop-btn pop-btn-nav" href="\${getGoogleMapsNavUrl(store)}" target="_blank" rel="noopener">Google 導航 ↗</a>
              <button class="pop-btn pop-btn-checkin \${isCheckedIn ? 'checked-in' : ''}" id="pop-checkin-\${store.id}">
                \${isCheckedIn ? '✅ 已打卡' : '📍 打卡'}
              </button>
            </div>
          \`;

          const btn = popupDiv.querySelector('#pop-checkin-' + store.id);
          btn.addEventListener('click', (e) => {
            e.stopPropagation();
            toggleCheckin(store.id);
            updatePopupContent();
            marker.setPopupContent(popupDiv);
          });

          return popupDiv;
        }

        marker.bindPopup(updatePopupContent());
        marker.on('click', () => {
          marker.setPopupContent(updatePopupContent());
          highlightStoreRow(store.id);
        });

        marker.addTo(map);

        storeStates.set(store.id, {
          store,
          marker,
          isVisible: true,
          rowElement: null,
          chkElement: null,
          checkinBtn: null
        });
      });

      // 渲染左側店家列表清單（帶獨立 Checkbox 與 打卡按鈕）
      const listEast = document.getElementById('list-east');
      const listWest = document.getElementById('list-west');
      const listDiff = document.getElementById('list-diff');

      stores.forEach((store, index) => {
        const state = storeStates.get(store.id);
        const row = document.createElement('div');
        row.className = 'store-row';
        row.id = 'store-row-' + store.id;

        const pinColor = store.subGroup === '主要名單-東區' ? '#2e7d32' : (store.subGroup === '主要名單-西區' ? '#d32f2f' : '#f59e0b');
        const isCheckedIn = checkinSet.has(store.id);

        row.innerHTML = \`
          <input type="checkbox" class="item-checkbox" id="chk-store-\${store.id}" checked />
          \${getSmallPinSvg(pinColor)}
          <div class="store-info">
            <div class="store-name-text" title="\${store.name}">\${index + 1}. \${store.name}</div>
            <div class="store-addr-text" title="\${store.address}">\${store.address}</div>
          </div>
          <button class="checkin-btn \${isCheckedIn ? 'checked-in' : ''}" id="btn-chk-\${store.id}" title="點擊打卡">
            \${isCheckedIn ? '✅ 已打卡' : '📍 打卡'}
          </button>
        \`;

        const chk = row.querySelector('.item-checkbox');
        const btnCheckin = row.querySelector('#btn-chk-' + store.id);
        state.rowElement = row;
        state.chkElement = chk;
        state.checkinBtn = btnCheckin;

        // 單獨 Checkbox 切換
        chk.addEventListener('click', (e) => {
          e.stopPropagation();
          setStoreVisibility(store.id, chk.checked);
        });

        // 打卡按鈕點擊事件（切換打卡 / 已打卡）
        btnCheckin.addEventListener('click', (e) => {
          e.stopPropagation();
          toggleCheckin(store.id);
        });

        // 點擊整列聚焦地圖
        row.addEventListener('click', () => {
          focusStoreOnMap(store.id);
        });

        if (store.subGroup === '主要名單-東區') {
          listEast.appendChild(row);
        } else if (store.subGroup === '主要名單-西區') {
          listWest.appendChild(row);
        } else {
          listDiff.appendChild(row);
        }
      });

      // 打卡核心切換函式
      function toggleCheckin(storeId) {
        const state = storeStates.get(storeId);
        if (!state) return;

        if (checkinSet.has(storeId)) {
          checkinSet.delete(storeId);
          state.checkinBtn.classList.remove('checked-in');
          state.checkinBtn.innerHTML = '📍 打卡';
          showToast('已取消【' + state.store.name + '】打卡');
        } else {
          checkinSet.add(storeId);
          state.checkinBtn.classList.add('checked-in');
          state.checkinBtn.innerHTML = '✅ 已打卡';
          showToast('✅ 已完成【' + state.store.name + '】打卡！');
        }
        saveCheckins();
      }

      function updateCheckinStats() {
        const label = document.getElementById('checkinStatsLabel');
        if (label) {
          label.textContent = '已打卡：' + checkinSet.size + ' 家';
        }
      }

      function setStoreVisibility(storeId, isVisible) {
        const state = storeStates.get(storeId);
        if (!state) return;
        state.isVisible = isVisible;
        state.chkElement.checked = isVisible;
        if (isVisible) {
          state.marker.addTo(map);
          state.rowElement.classList.remove('disabled');
        } else {
          map.removeLayer(state.marker);
          state.rowElement.classList.add('disabled');
        }
        updateGroupCheckboxes();
        updateStats();
      }

      function focusStoreOnMap(storeId) {
        const state = storeStates.get(storeId);
        if (!state) return;
        if (!state.isVisible) {
          setStoreVisibility(storeId, true);
        }
        map.flyTo([state.store.lat, state.store.lng], 16, { duration: 0.5 });
        state.marker.openPopup();
        highlightStoreRow(storeId);
      }

      let currentActiveRow = null;
      function highlightStoreRow(storeId) {
        if (currentActiveRow) {
          currentActiveRow.classList.remove('active');
        }
        const state = storeStates.get(storeId);
        if (state && state.rowElement) {
          state.rowElement.classList.add('active');
          state.rowElement.scrollIntoView({ block: 'nearest', behavior: 'smooth' });
          currentActiveRow = state.rowElement;
        }
      }

      function setGroupVisibility(groupName, isVisible) {
        stores.filter(s => s.subGroup === groupName).forEach(s => {
          const state = storeStates.get(s.id);
          state.isVisible = isVisible;
          state.chkElement.checked = isVisible;
          if (isVisible) {
            state.marker.addTo(map);
            state.rowElement.classList.remove('disabled');
          } else {
            map.removeLayer(state.marker);
            state.rowElement.classList.add('disabled');
          }
        });
        updateGroupCheckboxes();
        updateStats();
      }

      const chkGroupEast = document.getElementById('chk-group-east');
      const chkGroupWest = document.getElementById('chk-group-west');
      const chkGroupDiff = document.getElementById('chk-group-diff');

      chkGroupEast.addEventListener('change', () => setGroupVisibility('主要名單-東區', chkGroupEast.checked));
      chkGroupWest.addEventListener('change', () => setGroupVisibility('主要名單-西區', chkGroupWest.checked));
      chkGroupDiff.addEventListener('change', () => setGroupVisibility('非488名單', chkGroupDiff.checked));

      function updateGroupCheckboxes() {
        function checkStatus(groupName, groupChk) {
          const groupStores = stores.filter(s => s.subGroup === groupName);
          const visibleCount = groupStores.filter(s => storeStates.get(s.id).isVisible).length;
          if (visibleCount === groupStores.length) {
            groupChk.checked = true;
            groupChk.indeterminate = false;
          } else if (visibleCount === 0) {
            groupChk.checked = false;
            groupChk.indeterminate = false;
          } else {
            groupChk.checked = false;
            groupChk.indeterminate = true;
          }
          return { visibleCount, total: groupStores.length };
        }

        const eastStats = checkStatus('主要名單-東區', chkGroupEast);
        const westStats = checkStatus('主要名單-西區', chkGroupWest);
        const diffStats = checkStatus('非488名單', chkGroupDiff);

        document.getElementById('stats-east').textContent = \`顯示 \${eastStats.visibleCount} / \${eastStats.total} 筆\`;
        document.getElementById('stats-west').textContent = \`顯示 \${westStats.visibleCount} / \${westStats.total} 筆\`;
        document.getElementById('stats-diff').textContent = \`顯示 \${diffStats.visibleCount} / \${diffStats.total} 筆\`;
      }

      function updateStats() {
        let visibleCount = 0;
        storeStates.forEach(st => {
          if (st.isVisible) visibleCount++;
        });
        document.getElementById('activeCountLabel').textContent = \`目前顯示：\${visibleCount} / \${stores.length} 處店家\`;
      }

      // 快捷操作按鈕事件
      document.getElementById('btnShowAll').addEventListener('click', () => {
        stores.forEach(s => setStoreVisibility(s.id, true));
      });
      document.getElementById('btnHideAll').addEventListener('click', () => {
        stores.forEach(s => setStoreVisibility(s.id, false));
      });

      // 🌟 使用者指定：紅綠同時顯示（主要388筆，包含東區綠+西區紅，隱藏非488黃）
      document.getElementById('btnShowRedGreen').addEventListener('click', () => {
        stores.forEach(s => {
          const isRedOrGreen = (s.subGroup === '主要名單-東區' || s.subGroup === '主要名單-西區');
          setStoreVisibility(s.id, isRedOrGreen);
        });
        showToast('已切換為：🟢🔴 紅綠主要名單同時顯示 (388筆)');
      });

      document.getElementById('btnOnlyEast').addEventListener('click', () => {
        stores.forEach(s => setStoreVisibility(s.id, s.subGroup === '主要名單-東區'));
      });
      document.getElementById('btnOnlyWest').addEventListener('click', () => {
        stores.forEach(s => setStoreVisibility(s.id, s.subGroup === '主要名單-西區'));
      });
      document.getElementById('btnOnlyDiff').addEventListener('click', () => {
        stores.forEach(s => setStoreVisibility(s.id, s.subGroup === '非488名單'));
      });

      window.toggleGroupExpand = function(groupId) {
        const el = document.getElementById(groupId);
        if (el) el.classList.toggle('expanded');
      };

      const sidebar = document.getElementById('sidebar');
      const floatingToggleBtn = document.getElementById('floatingToggleBtn');
      const menuToggleBtn = document.getElementById('menuToggleBtn');
      const toggleIcon = document.getElementById('toggleIcon');
      const toggleText = document.getElementById('toggleText');

      function toggleSidebar() {
        const isCollapsed = sidebar.classList.toggle('collapsed');
        if (isCollapsed) {
          toggleIcon.textContent = '▶';
          toggleText.textContent = '展開清單';
        } else {
          toggleIcon.textContent = '◀';
          toggleText.textContent = '收合清單';
        }
        setTimeout(() => map.invalidateSize(), 260);
      }
      floatingToggleBtn.addEventListener('click', toggleSidebar);
      menuToggleBtn.addEventListener('click', toggleSidebar);

      const searchInput = document.getElementById('searchInput');
      const searchClearBtn = document.getElementById('searchClearBtn');

      searchInput.addEventListener('input', () => {
        const q = searchInput.value.trim().toLowerCase();
        searchClearBtn.style.display = q ? 'block' : 'none';

        stores.forEach(store => {
          const state = storeStates.get(store.id);
          const matches = !q || store.name.toLowerCase().includes(q) || store.address.toLowerCase().includes(q);
          state.rowElement.style.display = matches ? 'flex' : 'none';
        });
      });

      searchClearBtn.addEventListener('click', () => {
        searchInput.value = '';
        searchClearBtn.style.display = 'none';
        stores.forEach(store => {
          storeStates.get(store.id).rowElement.style.display = 'flex';
        });
      });

      // 初始化
      updateGroupCheckboxes();
      updateStats();
      updateCheckinStats();
    </script>
  </body>
</html>
`;

fs.writeFileSync(path.join(__dirname, "large.html"), html, "utf8");
fs.writeFileSync(path.join(__dirname, "dist", "large.html"), html, "utf8");
console.log("全新大型網頁已生成 (含紅綠同時顯示、黃色標記、打卡按鈕)！");
