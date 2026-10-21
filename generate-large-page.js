import fs from "node:fs";
import path from "node:path";
import { fileURLToPath } from "node:url";

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);

const jsonPath = path.join(__dirname, "large_stores.json");
const stores = JSON.parse(fs.readFileSync(jsonPath, "utf8"));
const eastCount = stores.filter(s => s.subGroup === "主要名單-東區").length;
const westCount = stores.filter(s => s.subGroup === "主要名單-西區").length;
const diffCount = stores.filter(s => s.subGroup === "非488名單").length;

const html = `<!doctype html>
<html lang="zh-Hant">
  <head>
    <meta charset="UTF-8" />
    <meta name="viewport" content="width=device-width, initial-scale=1.0" />
    <meta name="description" content="花蓮市與吉安鄉共 ${stores.length} 筆店家三色分區互動地圖：主要名單拆為東區（綠色 194 筆）、西區（紅色 194 筆），非 488 名單（亮橙色 112 筆），支援即時搜尋與導航。" />
    <title>花蓮店家地圖 (東區綠色 / 西區紅色 / 非488橙色 - 共 ${stores.length} 筆)</title>
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
        --accent: #257a4e;
        --primary-green: #257a4e;
        --primary-red: #dc2626;
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
      .search-input:focus { border-color: var(--primary-green); }
      .filter-tabs {
        display: grid;
        grid-template-columns: repeat(4, 1fr);
        gap: 4px;
        margin-top: 10px;
      }
      .filter-tab {
        text-align: center;
        padding: 6px 2px;
        font-size: 0.72rem;
        font-weight: 700;
        border-radius: 6px;
        cursor: pointer;
        border: 1px solid var(--line);
        background: white;
        color: var(--ink-soft);
        transition: all 140ms ease;
        white-space: nowrap;
        overflow: hidden;
        text-overflow: ellipsis;
      }
      .filter-tab.active {
        background: var(--ink);
        color: white;
        border-color: var(--ink);
      }
      .filter-tab.tab-east.active {
        background: var(--primary-green);
        border-color: var(--primary-green);
        color: white;
      }
      .filter-tab.tab-west.active {
        background: var(--primary-red);
        border-color: var(--primary-red);
        color: white;
      }
      .filter-tab.tab-diff.active {
        background: var(--diff-orange);
        border-color: var(--diff-orange);
        color: white;
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
        border-left: 4px solid var(--ink);
      }
      .store-item.is-diff {
        background: #fff8f5;
      }
      .store-item.is-diff:hover, .store-item.is-diff.active {
        background: #fdeee7;
      }
      .store-item.is-west:hover, .store-item.is-west.active {
        background: #fef2f2;
      }
      .store-item.is-east:hover, .store-item.is-east.active {
        background: #f0fdf4;
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
      .badge-east {
        background: #e4f3ea;
        color: var(--primary-green);
        border: 1px solid #bbf0d0;
      }
      .badge-west {
        background: #fee2e2;
        color: var(--primary-red);
        border: 1px solid #fecaca;
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
        color: #2563eb;
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
      .popup-group.east { background: #e4f3ea; color: var(--primary-green); }
      .popup-group.west { background: #fee2e2; color: var(--primary-red); }
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

      /* 自訂三色標記樣式 */
      .custom-pin {
        width: 14px;
        height: 14px;
        border-radius: 50%;
        border: 2px solid white;
        box-shadow: 0 0 4px rgba(0,0,0,0.4);
      }
      .custom-pin.pin-east {
        background: var(--primary-green);
        box-shadow: 0 0 5px rgba(37,122,78,0.7);
      }
      .custom-pin.pin-west {
        background: var(--primary-red);
        box-shadow: 0 0 5px rgba(220,38,38,0.7);
      }
      .custom-pin.pin-diff {
        width: 16px;
        height: 16px;
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
        <span>花蓮店家地圖 (三色分區版)</span>
        <span class="badge-count">全收錄 ${stores.length} 筆</span>
      </div>
      <div class="header-actions">
        <a href="./hualien_large_stores.kml" download="hualien_large_stores.kml" class="btn btn-primary" title="可直接匯入 Google 我的地圖（包含三組獨立資料夾與東綠/西紅/非橙圖標）">
          <span>↓ 整合 KML</span>
        </a>
        <a href="./hualien_east_green.kml" download="hualien_east_green.kml" class="btn" style="border-color:var(--primary-green);color:var(--primary-green);" title="🟢 東區主要名單 (194筆) 獨立圖層 KML">
          <span>🟢 東區 (${eastCount})</span>
        </a>
        <a href="./hualien_west_red.kml" download="hualien_west_red.kml" class="btn" style="border-color:var(--primary-red);color:var(--primary-red);" title="🔴 西區主要名單 (194筆) 獨立圖層 KML">
          <span>🔴 西區 (${westCount})</span>
        </a>
        <a href="./hualien_diff_orange.kml" download="hualien_diff_orange.kml" class="btn" style="border-color:var(--diff-orange);color:var(--diff-orange);" title="🟠 非488名單 (112筆) 獨立圖層 KML">
          <span>🟠 非488 (${diffCount})</span>
        </a>
        <a href="./hualien_google_mymaps.csv" download="hualien_google_mymaps.csv" class="btn" title="Google 我的地圖專用 CSV（依分區一鍵分色）">
          <span>↓ MyMaps CSV</span>
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
            <button class="filter-tab active" data-filter="all" title="全部店家">全部 (${stores.length})</button>
            <button class="filter-tab tab-east" data-filter="east" title="界線以東主要名單">🟢 東區 (${eastCount})</button>
            <button class="filter-tab tab-west" data-filter="west" title="界線以西主要名單">🔴 西區 (${westCount})</button>
            <button class="filter-tab tab-diff" data-filter="diff" title="非 488 名單店家">🟠 非488 (${diffCount})</button>
          </div>
          <div class="filter-stats">
            <span id="resultCount">顯示 ${stores.length} / ${stores.length} 筆</span>
            <span>以林森路+德安一街為界</span>
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

      // 繪製林森路 + 德安一街分界線輔助視覺
      const boundaryLine = L.polyline([
        [23.960, 121.581],
        [23.970, 121.5925],
        [23.983, 121.6070],
        [23.995, 121.620]
      ], {
        color: '#4b5563',
        weight: 2,
        dashArray: '6, 6',
        opacity: 0.55
      }).addTo(map);
      boundaryLine.bindTooltip('← 界線以西 (紅) ｜ 林森路+德安一街分界線 ｜ 界線以東 (綠) →', { sticky: true });

      const clusterGroup = L.markerClusterGroup({
        maxClusterRadius: 40,
        spiderfyOnMaxZoom: true,
        showCoverageOnHover: false,
        zoomToBoundsOnClick: true
      });

      const eastIcon = L.divIcon({
        className: '',
        html: '<div class="custom-pin pin-east"></div>',
        iconSize: [14, 14],
        iconAnchor: [7, 7]
      });

      const westIcon = L.divIcon({
        className: '',
        html: '<div class="custom-pin pin-west"></div>',
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
        let icon = eastIcon;
        let groupTag = '<span class="popup-group east">🟢 主要名單 - 界線以東</span>';
        if (store.subGroup === '主要名單-西區') {
          icon = westIcon;
          groupTag = '<span class="popup-group west">🔴 主要名單 - 界線以西</span>';
        } else if (store.subGroup === '非488名單') {
          icon = diffIcon;
          groupTag = '<span class="popup-group diff">🟠 非 488 名單店家</span>';
        }

        const marker = L.marker([store.lat, store.lng], { icon: icon });

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
          if (activeFilter === 'east' && s.subGroup !== '主要名單-東區') return false;
          if (activeFilter === 'west' && s.subGroup !== '主要名單-西區') return false;
          if (activeFilter === 'diff' && s.subGroup !== '非488名單') return false;
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
          
          let groupClass = ' is-east';
          let badgeHtml = '<span class="badge-tag badge-east">🟢 東區</span>';
          if (store.subGroup === '主要名單-西區') {
            groupClass = ' is-west';
            badgeHtml = '<span class="badge-tag badge-west">🔴 西區</span>';
          } else if (store.subGroup === '非488名單') {
            groupClass = ' is-diff';
            badgeHtml = '<span class="badge-tag badge-diff">🟠 非488</span>';
          }

          el.className = 'store-item' + groupClass + (store.id === currentActiveId ? ' active' : '');

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
console.log("三色分區版 large.html 網頁已生成！");
