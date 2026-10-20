import fs from "node:fs";
import path from "node:path";
import { fileURLToPath } from "node:url";

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);

const jsonPath = path.join(__dirname, "large_stores.json");
const stores = JSON.parse(fs.readFileSync(jsonPath, "utf8"));
const primaryCount = stores.filter(s => s.isPrimary488).length;
const diffCount = stores.filter(s => !s.isPrimary488).length;

const html = `<!doctype html>
<html lang="zh-Hant">
  <head>
    <meta charset="UTF-8" />
    <meta name="viewport" content="width=device-width, initial-scale=1.0" />
    <meta name="description" content="花蓮市與吉安鄉共 ${stores.length} 筆店家大規模互動地圖，支援雙色標記（主要488 vs 非488名單）、即時搜尋、群聚檢視與 Google Maps 導航。" />
    <title>花蓮店家地圖 (雙色標記 - 共 ${stores.length} 筆)</title>
    <link rel="preconnect" href="https://fonts.googleapis.com" />
    <link rel="preconnect" href="https://fonts.gstatic.com" crossorigin />
    <link href="https://fonts.googleapis.com/css2?family=Noto+Sans+TC:wght@400;500;600;700;800&display=swap" rel="stylesheet" />
    <link rel="stylesheet" href="https://unpkg.com/leaflet@1.9.4/dist/leaflet.css" />
    <link rel="stylesheet" href="https://unpkg.com/leaflet.markercluster@1.5.3/dist/MarkerCluster.css" />
    <link rel="stylesheet" href="https://unpkg.com/leaflet.markercluster@1.5.3/dist/MarkerCluster.Default.css" />
    <style>
      :root {
        --ink: #17352d;
        --ink-soft: #526760;
        --paper: #f8f5ee;
        --card: #ffffff;
        --accent: #e6613e;
        --accent-dark: #b84324;
        --primary-green: #257a4e;
        --diff-orange: #d9531e;
        --line: #ded7c8;
        --mint: #e8f0ec;
      }
      * { box-sizing: border-box; margin: 0; padding: 0; }
      body {
        font-family: "Noto Sans TC", system-ui, sans-serif;
        background: var(--paper);
        color: var(--ink);
        height: 100vh;
        display: flex;
        flex-direction: column;
        overflow: hidden;
      }
      header {
        background: white;
        border-bottom: 1px solid var(--line);
        padding: 12px 24px;
        display: flex;
        align-items: center;
        justify-content: space-between;
        flex-wrap: wrap;
        gap: 12px;
        z-index: 1000;
        box-shadow: 0 2px 10px rgba(0,0,0,0.04);
      }
      .brand-title {
        font-size: 1.25rem;
        font-weight: 800;
        color: var(--ink);
        display: flex;
        align-items: center;
        gap: 8px;
      }
      .badge-count {
        background: var(--mint);
        color: var(--ink);
        font-size: 0.8rem;
        padding: 3px 8px;
        border-radius: 999px;
        font-weight: 700;
      }
      .header-actions {
        display: flex;
        align-items: center;
        gap: 8px;
        flex-wrap: wrap;
      }
      .btn {
        display: inline-flex;
        align-items: center;
        gap: 6px;
        padding: 8px 14px;
        border-radius: 8px;
        font-size: 0.84rem;
        font-weight: 600;
        text-decoration: none;
        cursor: pointer;
        transition: all 160ms ease;
        border: 1px solid var(--line);
        background: white;
        color: var(--ink);
      }
      .btn:hover { background: #f3efe6; border-color: var(--ink-soft); }
      .btn-primary {
        background: var(--ink);
        color: white;
        border-color: var(--ink);
      }
      .btn-primary:hover { background: #245347; }
      .main-content {
        flex: 1;
        display: flex;
        position: relative;
        overflow: hidden;
      }
      .sidebar {
        width: 440px;
        background: white;
        border-right: 1px solid var(--line);
        display: flex;
        flex-direction: column;
        height: 100%;
        z-index: 500;
      }
      .search-bar {
        padding: 14px 16px;
        border-bottom: 1px solid var(--line);
        background: #faf8f3;
      }
      .search-input {
        width: 100%;
        padding: 9px 12px;
        border: 1px solid var(--line);
        border-radius: 8px;
        font-size: 0.9rem;
        font-family: inherit;
        outline: none;
        transition: border-color 160ms ease;
      }
      .search-input:focus { border-color: var(--accent); }
      .filter-tabs {
        display: flex;
        gap: 6px;
        margin-top: 10px;
      }
      .filter-tab {
        flex: 1;
        text-align: center;
        padding: 6px 4px;
        font-size: 0.76rem;
        font-weight: 700;
        border-radius: 6px;
        cursor: pointer;
        border: 1px solid var(--line);
        background: white;
        color: var(--ink-soft);
        transition: all 140ms ease;
      }
      .filter-tab.active {
        background: var(--ink);
        color: white;
        border-color: var(--ink);
      }
      .filter-tab.tab-diff.active {
        background: var(--diff-orange);
        border-color: var(--diff-orange);
      }
      .filter-tab.tab-primary.active {
        background: var(--primary-green);
        border-color: var(--primary-green);
      }
      .filter-stats {
        margin-top: 8px;
        font-size: 0.78rem;
        color: var(--ink-soft);
        display: flex;
        justify-content: space-between;
      }
      .store-list-wrap {
        flex: 1;
        overflow-y: auto;
      }
      .store-item {
        padding: 12px 16px;
        border-bottom: 1px solid var(--line);
        cursor: pointer;
        transition: background 120ms ease;
      }
      .store-item:hover, .store-item.active {
        background: #f4efe4;
      }
      .store-item.active {
        border-left: 4px solid var(--accent);
      }
      .store-item.is-diff {
        background: #fff8f5;
      }
      .store-item.is-diff:hover, .store-item.is-diff.active {
        background: #fdeee7;
      }
      .item-header {
        display: flex;
        align-items: flex-start;
        justify-content: space-between;
        gap: 8px;
      }
      .item-name {
        font-size: 0.94rem;
        font-weight: 700;
        color: var(--ink);
      }
      .item-badges {
        display: flex;
        align-items: center;
        gap: 4px;
        flex-shrink: 0;
      }
      .badge-tag {
        font-size: 0.68rem;
        font-weight: 800;
        padding: 2px 6px;
        border-radius: 4px;
      }
      .badge-primary {
        background: #e4f3ea;
        color: var(--primary-green);
      }
      .badge-diff {
        background: #fdeee7;
        color: var(--diff-orange);
        border: 1px solid #f9cdb8;
      }
      .item-address {
        font-size: 0.82rem;
        color: var(--ink-soft);
        margin-top: 4px;
      }
      .item-footer {
        margin-top: 6px;
        display: flex;
        align-items: center;
        justify-content: space-between;
      }
      .tag-est {
        font-size: 0.7rem;
        color: #b45309;
        font-weight: 600;
      }
      .nav-link {
        font-size: 0.78rem;
        color: var(--accent-dark);
        text-decoration: none;
        font-weight: 700;
      }
      .nav-link:hover { text-decoration: underline; }
      #map {
        flex: 1;
        height: 100%;
        background: #e5e3df;
      }
      .popup-group {
        display: inline-block;
        font-size: 0.72rem;
        font-weight: 800;
        padding: 2px 6px;
        border-radius: 4px;
        margin-bottom: 4px;
      }
      .popup-group.primary { background: #e4f3ea; color: var(--primary-green); }
      .popup-group.diff { background: #fdeee7; color: var(--diff-orange); }
      .popup-title { font-weight: 700; font-size: 0.96rem; margin-bottom: 4px; color: var(--ink); }
      .popup-addr { font-size: 0.82rem; color: var(--ink-soft); margin-bottom: 8px; }
      .popup-btn {
        display: inline-block;
        padding: 4px 10px;
        background: var(--ink);
        color: white !important;
        border-radius: 6px;
        font-size: 0.75rem;
        font-weight: 600;
        text-decoration: none;
      }

      /* 自訂雙色標記樣式 */
      .custom-pin {
        width: 14px;
        height: 14px;
        border-radius: 50%;
        border: 2px solid white;
        box-shadow: 0 0 4px rgba(0,0,0,0.4);
      }
      .custom-pin.pin-primary { background: var(--primary-green); }
      .custom-pin.pin-diff {
        background: var(--diff-orange);
        border: 2px solid #ffffff;
        box-shadow: 0 0 6px rgba(217,83,30,0.8);
      }

      @media (max-width: 860px) {
        body { height: auto; overflow: auto; }
        .main-content { flex-direction: column-reverse; height: auto; }
        .sidebar { width: 100%; height: 50vh; }
        #map { width: 100%; height: 50vh; }
      }
    </style>
  </head>
  <body>
    <header>
      <div class="brand-title">
        <span>花蓮店家地圖 (雙色標記)</span>
        <span class="badge-count">全收錄 ${stores.length} 筆</span>
      </div>
      <div class="header-actions">
        <a href="./hualien_large_stores.kml" download="hualien_large_stores.kml" class="btn btn-primary" title="可直接匯入 Google 我的地圖（包含雙色與雙資料夾）">
          <span>↓ 匯出 Google 地圖 (KML 雙色版)</span>
        </a>
        <a href="./hualien_large_stores.csv" download="hualien_large_stores.csv" class="btn" title="下載 CSV 試算表（含分組欄位）">
          <span>↓ 匯出 CSV</span>
        </a>
        <a href="./index.html" class="btn" title="切換回精選示範路線規劃">
          <span>← 路線規劃版</span>
        </a>
      </div>
    </header>

    <div class="main-content">
      <aside class="sidebar">
        <div class="search-bar">
          <input type="text" id="searchInput" class="search-input" placeholder="搜尋店名、道路名稱、號碼…" />
          <div class="filter-tabs">
            <button class="filter-tab active" data-filter="all">全部 (${stores.length})</button>
            <button class="filter-tab tab-primary" data-filter="primary">🟢 主要488名單 (${primaryCount})</button>
            <button class="filter-tab tab-diff" data-filter="diff">🟠 非488名單 (${diffCount})</button>
          </div>
          <div class="filter-stats">
            <span id="resultCount">顯示 ${stores.length} / ${stores.length} 筆</span>
            <span>點擊卡片聚焦標記</span>
          </div>
        </div>
        <div id="storeList" class="store-list-wrap"></div>
      </aside>
      <div id="map"></div>
    </div>

    <script src="https://unpkg.com/leaflet@1.9.4/dist/leaflet.js"></script>
    <script src="https://unpkg.com/leaflet.markercluster@1.5.3/dist/leaflet.markercluster.js"></script>
    <script>
      const stores = ${JSON.stringify(stores)};

      const map = L.map('map', { zoomControl: true }).setView([23.985, 121.605], 13);
      L.tileLayer('https://{s}.tile.openstreetmap.org/{z}/{x}/{y}.png', {
        maxZoom: 19,
        attribution: '&copy; <a href="https://www.openstreetmap.org/copyright">OpenStreetMap</a>'
      }).addTo(map);

      const clusterGroup = L.markerClusterGroup({
        maxClusterRadius: 40,
        spiderfyOnMaxZoom: true,
        showCoverageOnHover: false,
        zoomToBoundsOnClick: true
      });

      const primaryIcon = L.divIcon({
        className: '',
        html: '<div class="custom-pin pin-primary"></div>',
        iconSize: [14, 14],
        iconAnchor: [7, 7]
      });

      const diffIcon = L.divIcon({
        className: '',
        html: '<div class="custom-pin pin-diff"></div>',
        iconSize: [16, 16],
        iconAnchor: [8, 8]
      });

      const markerMap = new Map();
      const storeItemMap = new Map();

      function getGoogleMapsNavUrl(store) {
        return 'https://www.google.com/maps/dir/?api=1&destination=' + store.lat + ',' + store.lng;
      }

      stores.forEach((store, index) => {
        const icon = store.isPrimary488 ? primaryIcon : diffIcon;
        const marker = L.marker([store.lat, store.lng], { icon: icon });

        const groupTag = store.isPrimary488 
          ? '<span class="popup-group primary">主要 488 名單</span>' 
          : '<span class="popup-group diff">★ 非 488 名單 (另色標記)</span>';

        const popupContent = document.createElement('div');
        popupContent.innerHTML = 
          groupTag +
          '<div class="popup-title">' + (index + 1) + '. ' + store.name + '</div>' +
          '<div class="popup-addr">' + store.address + '</div>' +
          (store.estimated ? '<div style="color:#b45309;font-size:11px;margin-bottom:6px;">※ 位置為地址估算</div>' : '') +
          '<a class="popup-btn" href="' + getGoogleMapsNavUrl(store) + '" target="_blank" rel="noopener">Google Maps 導航 ↗</a>';
        
        marker.bindPopup(popupContent);
        marker.on('click', () => highlightStore(store.id, false));
        markerMap.set(store.id, marker);
      });

      function rebuildCluster(storesToShow) {
        clusterGroup.clearLayers();
        storesToShow.forEach(s => {
          const m = markerMap.get(s.id);
          if (m) clusterGroup.addLayer(m);
        });
      }

      map.addLayer(clusterGroup);

      const storeListEl = document.getElementById('storeList');
      const searchInput = document.getElementById('searchInput');
      const resultCountEl = document.getElementById('resultCount');
      const filterTabs = document.querySelectorAll('.filter-tab');

      let currentActiveId = null;
      let activeFilter = 'all';

      function highlightStore(storeId, scrollList = true) {
        if (currentActiveId && storeItemMap.has(currentActiveId)) {
          storeItemMap.get(currentActiveId).classList.remove('active');
        }
        currentActiveId = storeId;
        const itemEl = storeItemMap.get(storeId);
        if (itemEl) {
          itemEl.classList.add('active');
          if (scrollList) {
            itemEl.scrollIntoView({ block: 'nearest', behavior: 'smooth' });
          }
        }
        const store = stores.find(s => s.id === storeId);
        const marker = markerMap.get(storeId);
        if (store && marker) {
          clusterGroup.zoomToShowLayer(marker, () => {
            map.flyTo([store.lat, store.lng], 16, { duration: 0.5 });
            marker.openPopup();
          });
        }
      }

      function getFilteredStores() {
        const query = searchInput.value.trim().toLowerCase();
        return stores.filter(s => {
          if (activeFilter === 'primary' && !s.isPrimary488) return false;
          if (activeFilter === 'diff' && s.isPrimary488) return false;
          if (!query) return true;
          return s.name.toLowerCase().includes(query) || s.address.toLowerCase().includes(query);
        });
      }

      function renderList() {
        const filtered = getFilteredStores();
        storeListEl.innerHTML = '';
        storeItemMap.clear();

        filtered.forEach((store) => {
          const originalIndex = stores.findIndex(s => s.id === store.id);
          const el = document.createElement('div');
          el.className = 'store-item' + 
            (store.isPrimary488 ? '' : ' is-diff') + 
            (store.id === currentActiveId ? ' active' : '');

          const badgeHtml = store.isPrimary488
            ? '<span class="badge-tag badge-primary">主要488</span>'
            : '<span class="badge-tag badge-diff">非488</span>';

          el.innerHTML = 
            '<div class="item-header">' +
              '<span class="item-name">' + store.name + '</span>' +
              '<div class="item-badges">' +
                badgeHtml +
                '<span style="font-size:11px;color:#888;">#' + (originalIndex + 1) + '</span>' +
              '</div>' +
            '</div>' +
            '<div class="item-address">' + store.address + '</div>' +
            '<div class="item-footer">' +
              (store.estimated ? '<span class="tag-est">估算位置</span>' : '<span></span>') +
              '<a class="nav-link" href="' + getGoogleMapsNavUrl(store) + '" target="_blank" rel="noopener" onclick="event.stopPropagation()">導航 ↗</a>' +
            '</div>';

          el.addEventListener('click', () => highlightStore(store.id, false));
          storeListEl.appendChild(el);
          storeItemMap.set(store.id, el);
        });

        resultCountEl.textContent = '顯示 ' + filtered.length + ' / ' + stores.length + ' 筆';
        rebuildCluster(filtered);
      }

      filterTabs.forEach(tab => {
        tab.addEventListener('click', () => {
          filterTabs.forEach(t => t.classList.remove('active'));
          tab.classList.add('active');
          activeFilter = tab.dataset.filter;
          renderList();
        });
      });

      searchInput.addEventListener('input', () => {
        renderList();
      });

      renderList();
    </script>
  </body>
</html>
`;

fs.writeFileSync(path.join(__dirname, "large.html"), html, "utf8");
fs.writeFileSync(path.join(__dirname, "dist", "large.html"), html, "utf8");
console.log("雙色版 large.html 網頁已生成！");
