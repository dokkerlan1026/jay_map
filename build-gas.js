import fs from "node:fs";
import path from "node:path";
import { fileURLToPath } from "node:url";

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);

const jsonPath = path.join(__dirname, "large_stores.json");
const stores = JSON.parse(fs.readFileSync(jsonPath, "utf8"));

// 確保 gas 目錄存在
const gasDir = path.join(__dirname, "gas");
if (!fs.existsSync(gasDir)) {
  fs.mkdirSync(gasDir, { recursive: true });
}

// 1. 生成 gas/Code.gs（包含完整 500 筆 JSON 資料與 doGet、setupSheetWithData）
const codeGsTemplate = fs.readFileSync(path.join(gasDir, "Code.gs"), "utf8");
const fullCodeGs = codeGsTemplate.replace(
  "STORE_DATA_PLACEHOLDER",
  JSON.stringify(stores, null, 2)
);
fs.writeFileSync(path.join(gasDir, "Code.gs"), fullCodeGs, "utf8");
console.log("gas/Code.gs 已成功生成！");

// 2. 生成 gas/Index.html（Google Apps Script 專用網頁樣板）
const eastStores = stores.filter(s => s.subGroup === "主要名單-東區");
const westStores = stores.filter(s => s.subGroup === "主要名單-西區");
const diffStores = stores.filter(s => s.subGroup === "非488名單");

const indexHtml = `<!doctype html>
<html lang="zh-Hant">
  <head>
    <meta charset="UTF-8" />
    <meta name="viewport" content="width=device-width, initial-scale=1.0" />
    <title>jay的地圖 - 花蓮店家地圖</title>
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
        --color-diff: #e65100;       /* 橙色 */
        --border-color: #dadce0;
        --text-primary: #202124;
        --text-secondary: #5f6368;
        --sidebar-width: 380px;
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
        gap: 16px;
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
      .gas-badge {
        font-size: 0.7rem;
        background: rgba(255,255,255,0.25);
        padding: 2px 8px;
        border-radius: 10px;
        font-weight: 600;
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

      /* 地圖資訊與快速操作區 */
      .panel-intro {
        padding: 12px 16px;
        border-bottom: 1px solid var(--border-color);
        background: #fafafa;
      }
      .intro-meta {
        font-size: 0.8rem;
        color: var(--text-secondary);
        display: flex;
        align-items: center;
        justify-content: space-between;
        margin-bottom: 8px;
      }
      .quick-actions {
        display: flex;
        gap: 6px;
        flex-wrap: wrap;
      }
      .btn-chip {
        padding: 4px 8px;
        border: 1px solid var(--border-color);
        background: white;
        border-radius: 14px;
        font-size: 0.74rem;
        font-weight: 600;
        color: var(--text-secondary);
        cursor: pointer;
        transition: all 120ms;
      }
      .btn-chip:hover {
        background: var(--google-hover);
        color: var(--text-primary);
        border-color: #bdc1c6;
      }
      .btn-chip.primary {
        background: #e8f0fe;
        color: #1a73e8;
        border-color: #d2e3fc;
      }

      /* 搜尋列 */
      .search-box {
        padding: 8px 16px 12px;
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
        padding: 8px 32px 8px 10px;
        border: 1px solid var(--border-color);
        border-radius: 6px;
        font-size: 0.88rem;
        font-family: inherit;
        outline: none;
        transition: border-color 150ms;
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

      /* 單一圖層群組（東區 / 西區 / 非488） */
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
        font-size: 0.92rem;
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

      /* 群組底下的店家列表 */
      .item-list {
        display: none;
        background: #ffffff;
      }
      .layer-group.expanded .item-list {
        display: block;
      }

      /* 單一店家項目（支援獨立勾選） */
      .store-row {
        display: flex;
        align-items: center;
        padding: 8px 12px 8px 24px;
        border-top: 1px solid #f1f3f4;
        cursor: pointer;
        transition: background 100ms;
        gap: 10px;
      }
      .store-row:hover {
        background: #f8f9fa;
      }
      .store-row.active {
        background: #e8f0fe;
      }
      .store-row.disabled {
        opacity: 0.45;
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

      /* 側邊欄收合展開切換浮動按鈕 */
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
      .toggle-sidebar-btn:hover {
        background: #f1f3f4;
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

      /* Leaflet Popup 美化 */
      .leaflet-popup-content-wrapper {
        border-radius: 8px;
        box-shadow: 0 2px 10px rgba(0,0,0,0.22);
        padding: 0;
        overflow: hidden;
      }
      .leaflet-popup-content {
        margin: 0;
        padding: 14px 16px;
        min-width: 240px;
        max-width: 300px;
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
      .pop-badge.diff { background: #fff3e0; color: var(--color-diff); }

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
        gap: 8px;
      }
      .pop-nav-btn {
        flex: 1;
        text-align: center;
        padding: 6px 10px;
        background: #1a73e8;
        color: white !important;
        text-decoration: none;
        border-radius: 4px;
        font-size: 0.8rem;
        font-weight: 600;
      }
      .pop-nav-btn:hover { background: #1557b0; }
      .pop-toggle-btn {
        padding: 6px 10px;
        background: #f1f3f4;
        border: 1px solid var(--border-color);
        border-radius: 4px;
        font-size: 0.78rem;
        font-weight: 600;
        cursor: pointer;
        color: var(--text-secondary);
      }
      .pop-toggle-btn:hover { background: #e8eaed; color: var(--text-primary); }

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
            <span class="gas-badge">Google Apps Script Web App</span>
          </div>
          <div class="brand-sub">花蓮店家地圖 (東區綠色 / 西區紅色 / 非488橙色)</div>
        </div>
      </div>
    </header>

    <div class="layout-body">
      <aside class="sidebar" id="sidebar">
        <!-- 頂部資訊與快速操作按鈕 -->
        <div class="panel-intro">
          <div class="intro-meta">
            <span id="activeCountLabel">目前顯示：載入中...</span>
            <span>以林森路+德安一街為界</span>
          </div>
          <div class="quick-actions">
            <button class="btn-chip primary" id="btnShowAll">全部顯示</button>
            <button class="btn-chip" id="btnHideAll">全部隱藏</button>
            <button class="btn-chip" id="btnOnlyEast">僅看東區 (綠)</button>
            <button class="btn-chip" id="btnOnlyWest">僅看西區 (紅)</button>
            <button class="btn-chip" id="btnOnlyDiff">僅看非488 (橙)</button>
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
                <span class="group-stats" id="stats-east">計算中...</span>
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
                <span class="group-stats" id="stats-west">計算中...</span>
              </div>
              <div class="arrow-btn" onclick="toggleGroupExpand('group-west')">▼</div>
            </div>
            <div class="item-list" id="list-west"></div>
          </div>

          <!-- 3. 非 488 店家名單 (橙色) -->
          <div class="layer-group expanded" id="group-diff">
            <div class="group-header">
              <input type="checkbox" class="group-checkbox" id="chk-group-diff" checked title="全選/取消非488名單" />
              <div class="group-icon diff"></div>
              <div class="group-title-wrap" onclick="toggleGroupExpand('group-diff')">
                <span class="group-title">🟠 非488名單 (亮橙色標記)</span>
                <span class="group-stats" id="stats-diff">計算中...</span>
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

    <script src="https://unpkg.com/leaflet@1.9.4/dist/leaflet.js"></script>
    <script>
      // 從 Apps Script 後端樣板注入資料，或 Fallback 預設資料
      let stores = [];
      try {
        stores = JSON.parse(<?= storesJson ?>);
      } catch (e) {
        stores = ${JSON.stringify(stores)};
      }

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
      const iconDiff = createPinIcon('#e65100'); // 橙色

      function getSmallPinSvg(colorHex) {
        return \`<svg class="pin-svg" viewBox="0 0 24 34" fill="none"><path d="M12 0C5.372 0 0 5.373 0 12C0 19.5 12 34 12 34C12 34 24 19.5 24 12C24 5.373 18.628 0 12 0Z" fill="\${colorHex}" stroke="#ffffff" stroke-width="1.5"/><circle cx="12" cy="11" r="4" fill="#ffffff"/></svg>\`;
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
          badgeText = '🟠 非 488 名單店家';
        }

        const marker = L.marker([store.lat, store.lng], { icon: icon });

        const popupDiv = document.createElement('div');
        popupDiv.innerHTML = \`
          <div class="pop-header">
            <span class="pop-badge \${badgeClass}">\${badgeText}</span>
            <span style="font-size:0.75rem;color:#888;">#\${index + 1}</span>
          </div>
          <div class="pop-title">\${store.name}</div>
          <div class="pop-addr">\${store.address}\${store.estimated ? '<br/><span style="color:#b45309;font-size:11px;">※ 位置為地址估算</span>' : ''}</div>
          <div class="pop-actions">
            <a class="pop-nav-btn" href="\${getGoogleMapsNavUrl(store)}" target="_blank" rel="noopener">Google 導航 ↗</a>
            <button class="pop-toggle-btn" onclick="toggleSingleStoreVisible(\${store.id}, false)">隱藏此點</button>
          </div>
        \`;
        marker.bindPopup(popupDiv);
        marker.on('click', () => highlightStoreRow(store.id));
        marker.addTo(map);

        storeStates.set(store.id, {
          store,
          marker,
          isVisible: true,
          rowElement: null,
          chkElement: null
        });
      });

      // 渲染左側店家列表清單
      const listEast = document.getElementById('list-east');
      const listWest = document.getElementById('list-west');
      const listDiff = document.getElementById('list-diff');

      stores.forEach((store, index) => {
        const state = storeStates.get(store.id);
        const row = document.createElement('div');
        row.className = 'store-row';
        row.id = 'store-row-' + store.id;

        const pinColor = store.subGroup === '主要名單-東區' ? '#2e7d32' : (store.subGroup === '主要名單-西區' ? '#d32f2f' : '#e65100');

        row.innerHTML = \`
          <input type="checkbox" class="item-checkbox" id="chk-store-\${store.id}" checked />
          \${getSmallPinSvg(pinColor)}
          <div class="store-info">
            <div class="store-name-text" title="\${store.name}">\${index + 1}. \${store.name}</div>
            <div class="store-addr-text" title="\${store.address}">\${store.address}</div>
          </div>
        \`;

        const chk = row.querySelector('.item-checkbox');
        state.rowElement = row;
        state.chkElement = chk;

        chk.addEventListener('click', (e) => {
          e.stopPropagation();
          setStoreVisibility(store.id, chk.checked);
        });

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

      window.toggleSingleStoreVisible = function(storeId, forceState) {
        const state = storeStates.get(storeId);
        if (state) {
          const next = forceState !== undefined ? forceState : !state.isVisible;
          setStoreVisibility(storeId, next);
        }
      };

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

      document.getElementById('btnShowAll').addEventListener('click', () => {
        stores.forEach(s => setStoreVisibility(s.id, true));
      });
      document.getElementById('btnHideAll').addEventListener('click', () => {
        stores.forEach(s => setStoreVisibility(s.id, false));
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

      updateGroupCheckboxes();
      updateStats();
    </script>
  </body>
</html>
`;

fs.writeFileSync(path.join(gasDir, "Index.html"), indexHtml, "utf8");
console.log("gas/Index.html 已成功生成！");
