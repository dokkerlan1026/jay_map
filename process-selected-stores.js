import fs from "node:fs";
import path from "node:path";
import { fileURLToPath } from "node:url";

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);

const CACHE_FILE = path.join(__dirname, "geocode_cache.json");
let cache = {};
if (fs.existsSync(CACHE_FILE)) {
  try {
    cache = JSON.parse(fs.readFileSync(CACHE_FILE, "utf8"));
  } catch (e) {
    console.error(e);
  }
}

function cleanAddress(rawAddr) {
  let addr = rawAddr.trim();

  // 特殊固定點修正
  if (addr.includes("東大門夜市")) return { clean: "花蓮市東大門夜市", estimated: true, defaultCoord: { lat: 23.974012, lng: 121.61352 } };
  if (addr.includes("國民段466") || addr.includes("國民段483")) return { clean: "花蓮市國聯一路100號", estimated: true, defaultCoord: { lat: 23.9926, lng: 121.6034 } };
  if (addr.includes("富國段富國一小段")) return { clean: "花蓮市富安路", estimated: true, defaultCoord: { lat: 23.9938, lng: 121.5995 } };
  if (addr.includes("福德段91")) return { clean: "花蓮市南濱路", estimated: true, defaultCoord: { lat: 23.9705, lng: 121.6110 } };
  if (addr.includes("綜合市場")) return { clean: "花蓮市復興街61號", estimated: true, defaultCoord: { lat: 23.9785, lng: 121.6075 } };

  let isEstimated = false;

  addr = addr.replace("花蓮縣花蓮市吉安鄉", "花蓮縣吉安鄉");
  addr = addr.replace("吉安鄉吉安鄉", "吉安鄉");
  addr = addr.replace("花蓮市花蓮市", "花蓮市");

  if (/旁|旁邊|對面|空地|攤位|第\d+棟|火車站商業區域|地下室|倉庫|住家|藥局|診所|金庫|內部|冰箱疫苗|大樓教室|教室|頂樓|辦公室/.test(addr)) {
    isEstimated = true;
  }

  addr = addr.replace(/[\(（].*?[\)）]/g, "");
  addr = addr.replace(/(旁|旁邊|對面).*$/, "");
  addr = addr.replace(/[B\d]+樓|B\d+|\d+F|\d+、\d+樓|\d+-\d+樓/g, "");
  addr = addr.replace(/(\d+)[,、\.\s\-~～]\d+.*?號/, "$1號");
  addr = addr.replace(/(\d+)之(\d+)$/, "$1-$2號");

  return { clean: addr.trim(), estimated: isEstimated };
}

async function fetchArcGis(address) {
  const url = `https://geocode.arcgis.com/arcgis/rest/services/World/GeocodeServer/findAddressCandidates?f=json&singleLine=${encodeURIComponent(address)}&location=121.60,23.98&distance=15000&maxLocations=1`;
  const res = await fetch(url);
  if (!res.ok) throw new Error(`ArcGIS error ${res.status}`);
  const data = await res.json();
  const candidate = data.candidates?.[0];
  if (candidate && candidate.location) {
    const lat = candidate.location.y;
    const lng = candidate.location.x;
    if (lat >= 23.90 && lat <= 24.08 && lng >= 121.50 && lng <= 121.68) {
      return { lat, lng, score: candidate.score };
    }
  }
  return null;
}

async function fetchNominatim(address) {
  const url = `https://nominatim.openstreetmap.org/search?format=json&q=${encodeURIComponent(address)}&countrycodes=tw`;
  const res = await fetch(url, { headers: { "User-Agent": "HualienGeocoder488/1.0" } });
  if (!res.ok) return null;
  const data = await res.json();
  if (data && data[0]) {
    const lat = parseFloat(data[0].lat);
    const lng = parseFloat(data[0].lon);
    if (lat >= 23.90 && lat <= 24.08 && lng >= 121.50 && lng <= 121.68) {
      return { lat, lng };
    }
  }
  return null;
}

const rawLines = fs.readFileSync(path.join(__dirname, "selected_stores_raw.txt"), "utf8")
  .split("\n")
  .map(l => l.trim())
  .filter(l => l.length > 0);

const stores = [];

console.log(`開始處理最新清單，共 ${rawLines.length} 行...`);

for (let i = 0; i < rawLines.length; i++) {
  const line = rawLines[i];
  const parts = line.split("\t");
  const name = parts[0]?.trim();
  const address = parts[1]?.trim();
  if (!name || !address) continue;

  const { clean, estimated, defaultCoord } = cleanAddress(address);

  let coord = null;
  let isEstimated = estimated;

  if (defaultCoord) {
    coord = defaultCoord;
    isEstimated = true;
  } else if (cache[clean]) {
    coord = cache[clean];
  } else {
    try {
      coord = await fetchArcGis(clean);
      if (!coord) {
        coord = await fetchArcGis(clean.replace(/^花蓮縣/, ""));
      }
      if (!coord) {
        coord = await fetchNominatim(clean);
      }
      if (coord) {
        cache[clean] = { lat: coord.lat, lng: coord.lng };
        console.log(`[新查得] ${clean} => ${coord.lat}, ${coord.lng}`);
      } else {
        console.warn(`[未找到] ${name} (${clean})，使用備用座標`);
        coord = { lat: 23.9872, lng: 121.6016 };
        isEstimated = true;
      }
    } catch (e) {
      console.error(`查詢失敗 ${clean}:`, e.message);
      coord = { lat: 23.9872, lng: 121.6016 };
      isEstimated = true;
    }
    await new Promise(r => setTimeout(r, 60));
  }

  stores.push({
    id: `store-${stores.length + 1}`,
    name,
    address,
    lat: Number(coord.lat.toFixed(6)),
    lng: Number(coord.lng.toFixed(6)),
    estimated: isEstimated || (coord.score && coord.score < 90),
  });
}

// 儲存快取
fs.writeFileSync(CACHE_FILE, JSON.stringify(cache, null, 2), "utf8");
console.log(`全部處理完成！有效店家筆數: ${stores.length}`);

// 1. 輸出 JSON
fs.writeFileSync(path.join(__dirname, "large_stores.json"), JSON.stringify(stores, null, 2), "utf8");

// 2. 輸出 CSV (帶 UTF-8 BOM)
const csvHeader = "\uFEFF序號,店名,地址,緯度,經度,備註\n";
const csvRows = stores.map((s, idx) => {
  const note = s.estimated ? "位置為地址估算" : "";
  const safeName = `"${s.name.replace(/"/g, '""')}"`;
  const safeAddress = `"${s.address.replace(/"/g, '""')}"`;
  return `${idx + 1},${safeName},${safeAddress},${s.lat},${s.lng},"${note}"`;
}).join("\n");
const csvContent = csvHeader + csvRows;

fs.writeFileSync(path.join(__dirname, "hualien_large_stores.csv"), csvContent, "utf8");
fs.writeFileSync(path.join(__dirname, "dist", "hualien_large_stores.csv"), csvContent, "utf8");

// 3. 輸出 KML (Google Earth / Google My Maps 標準格式)
const kmlPlacemarks = stores.map((s, idx) => {
  const note = s.estimated ? "<br/>※ 位置為地址估算" : "";
  return `    <Placemark>
      <name><![CDATA[${idx + 1}. ${s.name}]]></name>
      <description><![CDATA[地址：${s.address}${note}]]></description>
      <Point>
        <coordinates>${s.lng},${s.lat},0</coordinates>
      </Point>
    </Placemark>`;
}).join("\n");

const kmlContent = `<?xml version="1.0" encoding="UTF-8"?>
<kml xmlns="http://www.opengis.net/kml/2.2">
  <Document>
    <name>花蓮店家地圖 (共 ${stores.length} 筆)</name>
    <description>花蓮市與吉安鄉最新主要店家與地點清單 (共 ${stores.length} 筆)</description>
${kmlPlacemarks}
  </Document>
</kml>
`;

fs.writeFileSync(path.join(__dirname, "hualien_large_stores.kml"), kmlContent, "utf8");
fs.writeFileSync(path.join(__dirname, "dist", "hualien_large_stores.kml"), kmlContent, "utf8");
console.log("KML 與 CSV 匯出檔更新完成！");
