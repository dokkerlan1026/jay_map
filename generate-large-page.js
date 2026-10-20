import fs from "node:fs";
import path from "node:path";
import { fileURLToPath } from "node:url";

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);

const jsonPath = path.join(__dirname, "large_stores.json");
if (!fs.existsSync(jsonPath)) {
  console.log("等待 large_stores.json 產出...");
  process.exit(1);
}

const stores = JSON.parse(fs.readFileSync(jsonPath, "utf8"));

const html = `<!doctype html>
<html lang="zh-Hant">
  <head>
    <meta charset="UTF-8" />
    <meta name="viewport" content="width=device-width, initial-scale=1.0" />
    <meta name="description" content="花蓮市與吉安鄉共 ${stores.length} 筆店家大規模互動地圖，支援即時搜尋、群聚檢視、Google Maps 導航與一鍵匯出。" />
    <title>花蓮大規模店家地圖 (共 ${stores.length} 筆)</title>
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
        width: 420px;
        background: white;
        border-right: 1px solid var(--line);
        display: flex;
        flex-direction: column;
        height: 100%;
        z-index: 500;
      }
      .search-bar {
        padding: 16px;
        border-bottom: 1px solid var(--line);
        background: #faf8f3;
      }
      .search-input {
        width: 100%;
        padding: 10px 14px;
        border: 1px solid var(--line);
        border-radius: 8px;
        font-size: 0.92rem;
        font-family: inherit;
        outline: none;
        transition: border-color 160ms ease;
      }
      .search-input:focus { border-color: var(--accent); }
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
        padding: 14px 16px;
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
      .item-header {
        display: flex;
        align-items: flex-start;
        justify-content: space-between;
        gap: 8px;
      }
      .item-name {
        font-size: 0.95rem;
        font-weight: 700;
        color: var(--ink);
      }
      .item-index {
        font-size: 0.72rem;
        font-weight: 800;
        color: var(--accent-dark);
      }
      .item-address {
        font-size: 0.82rem;
        color: var(--ink-soft);
        margin-top: 4px;
      }
      .item-footer {
        margin-top: 8px;
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
        <span>花蓮大規模店家地圖</span>
        <span class="badge-count">全收錄 ${stores.length} 筆</span>
      </div>
      <div class="header-actions">
        <a href="./hualien_large_stores.kml" download="hualien_large_stores.kml" class="btn btn-primary" title="可直接匯入 Google 我的地圖">
          <span>↓ 匯出 Google 地圖 (KML)</span>
        </a>
        <a href="./hualien_large_stores.csv" download="hualien_large_stores.csv" class="btn" title="下載 CSV 試算表">
          <span>↓ 匯出 CSV</span>
        </a>
        <a href="./index.html" class="btn" title="切換回 25 家精選示範與最佳路線規劃">
          <span>← 精選 25 店路線規劃</span>
        </a>
      </div>
    </header>

    <div class="main-content">
      <aside class="sidebar">
        <div class="search-bar">
          <input type="text" id="searchInput" class="search-input" placeholder="搜尋店家名稱、道路名稱、號碼…" />
          <div class="filter-stats">
            <span id="resultCount">顯示 ${stores.length} / ${stores.length} 筆</span>
            <span>點擊項目聚焦地圖</span>
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

      const markerMap = new Map();
      const storeItemMap = new Map();

      function getGoogleMapsNavUrl(store) {
        return 'https://www.google.com/maps/dir/?api=1&destination=' + store.lat + ',' + store.lng;
      }

      stores.forEach((store, index) => {
        const marker = L.marker([store.lat, store.lng]);
        const popupContent = document.createElement('div');
        popupContent.innerHTML = 
          '<div class="popup-title">' + (index + 1) + '. ' + store.name + '</div>' +
          '<div class="popup-addr">' + store.address + '</div>' +
          (store.estimated ? '<div style="color:#b45309;font-size:11px;margin-bottom:6px;">※ 位置為地址估算</div>' : '') +
          '<a class="popup-btn" href="' + getGoogleMapsNavUrl(store) + '" target="_blank" rel="noopener">Google Maps 導航 ↗</a>';
        
        marker.bindPopup(popupContent);
        marker.on('click', () => highlightStore(store.id, false));
        clusterGroup.addLayer(marker);
        markerMap.set(store.id, marker);
      });

      map.addLayer(clusterGroup);

      const storeListEl = document.getElementById('storeList');
      const searchInput = document.getElementById('searchInput');
      const resultCountEl = document.getElementById('resultCount');

      let currentActiveId = null;

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

      function renderList(filteredStores) {
        storeListEl.innerHTML = '';
        storeItemMap.clear();

        filteredStores.forEach((store) => {
          const originalIndex = stores.findIndex(s => s.id === store.id);
          const el = document.createElement('div');
          el.className = 'store-item' + (store.id === currentActiveId ? ' active' : '');
          el.innerHTML = 
            '<div class="item-header">' +
              '<span class="item-name">' + store.name + '</span>' +
              '<span class="item-index">#' + (originalIndex + 1) + '</span>' +
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

        resultCountEl.textContent = '顯示 ' + filteredStores.length + ' / ' + stores.length + ' 筆';
      }

      searchInput.addEventListener('input', (e) => {
        const query = e.target.value.trim().toLowerCase();
        if (!query) {
          renderList(stores);
          return;
        }
        const filtered = stores.filter(s => 
          s.name.toLowerCase().includes(query) || 
          s.address.toLowerCase().includes(query)
        );
        renderList(filtered);
      });

      renderList(stores);
    </script>
  </body>
</html>
`;

fs.writeFileSync(path.join(__dirname, "large.html"), html, "utf8");
fs.writeFileSync(path.join(__dirname, "dist", "large.html"), html, "utf8");
console.log("large.html 網頁已生成！");
