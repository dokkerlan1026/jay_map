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
    console.error("Cache read error:", e);
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

  // 修正重複區域名
  addr = addr.replace("花蓮縣花蓮市吉安鄉", "花蓮縣吉安鄉");
  addr = addr.replace("吉安鄉吉安鄉", "吉安鄉");
  addr = addr.replace("花蓮市花蓮市", "花蓮市");

  // 移除非門牌資訊
  if (/旁|旁邊|對面|空地|攤位|第\d+棟|火車站商業區域|地下室|倉庫|住家|藥局|診所|金庫|內部/.test(addr)) {
    isEstimated = true;
  }

  // 去除括號註記
  addr = addr.replace(/[\(（].*?[\)）]/g, "");

  // 去除旁/旁邊/對面
  addr = addr.replace(/(旁|旁邊|對面).*$/, "");

  // 去除樓層與地下室
  addr = addr.replace(/[B\d]+樓|B\d+|\d+F|\d+、\d+樓/g, "");

  // 去除複合門牌，取第一個，例如 119,119-3號 -> 119號，359.361.363號 -> 359號，48.50號 -> 48號
  addr = addr.replace(/(\d+)[,、\.\-~～]\d+.*?號/, "$1號");

  // 去除之幾結尾沒有號，例如 195之1 -> 195-1號
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
    // 檢查範圍是否在花蓮大生活圈
    if (lat >= 23.90 && lat <= 24.08 && lng >= 121.50 && lng <= 121.68) {
      return { lat, lng, score: candidate.score };
    }
  }
  return null;
}

async function fetchNominatim(address) {
  const url = `https://nominatim.openstreetmap.org/search?format=json&q=${encodeURIComponent(address)}&countrycodes=tw`;
  const res = await fetch(url, { headers: { "User-Agent": "HualienStoreGeocoder/1.0" } });
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

export async function geocodeAll() {
  const rawLines = fs.readFileSync(path.join(__dirname, "raw_large_stores.txt"), "utf8")
    .trim()
    .split("\n");

  const results = [];
  let cacheHits = 0;
  let apiHits = 0;

  for (let i = 0; i < rawLines.length; i++) {
    const line = rawLines[i].trim();
    if (!line) continue;
    const [name, address] = line.split("\t").map(s => s?.trim());
    if (!name || !address) continue;

    const { clean, estimated, defaultCoord } = cleanAddress(address);

    let coord = null;
    let isEstimated = estimated;

    if (defaultCoord) {
      coord = defaultCoord;
      isEstimated = true;
    } else if (cache[clean]) {
      coord = cache[clean];
      cacheHits++;
    } else {
      try {
        coord = await fetchArcGis(clean);
        if (!coord) {
          // 嘗試拿掉「花蓮縣」只留「花蓮市」或「吉安鄉」
          const simplified = clean.replace(/^花蓮縣/, "");
          coord = await fetchArcGis(simplified);
        }
        if (!coord) {
          coord = await fetchNominatim(clean);
        }
        if (coord) {
          cache[clean] = { lat: coord.lat, lng: coord.lng };
          apiHits++;
          // 每 10 筆存一次快取
          if (apiHits % 10 === 0) {
            fs.writeFileSync(CACHE_FILE, JSON.stringify(cache, null, 2), "utf8");
          }
        } else {
          console.warn(`[未找到精準座標] 第 ${i + 1} 筆: ${name} (${address} -> ${clean})，使用備用估算座標`);
          // 備用座標：花蓮市中心
          coord = { lat: 23.987, lng: 121.603 };
          isEstimated = true;
        }
      } catch (err) {
        console.error(`查詢失敗 [${name}]:`, err.message);
        coord = { lat: 23.987, lng: 121.603 };
        isEstimated = true;
      }
      // 適度延遲避免被 rate limit
      await new Promise(r => setTimeout(r, 60));
    }

    results.push({
      id: `store-${i + 1}`,
      name,
      address,
      lat: Number(coord.lat.toFixed(6)),
      lng: Number(coord.lng.toFixed(6)),
      estimated: isEstimated || coord.score < 90,
    });

    if ((i + 1) % 50 === 0 || i === rawLines.length - 1) {
      console.log(`已處理 ${i + 1} / ${rawLines.length} 筆... (快取命中: ${cacheHits}, API 查詢: ${apiHits})`);
    }
  }

  // 儲存完整快取
  fs.writeFileSync(CACHE_FILE, JSON.stringify(cache, null, 2), "utf8");

  return results;
}

if (process.argv[1] === fileURLToPath(import.meta.url)) {
  const items = await geocodeAll();
  console.log(`總計完成 ${items.length} 筆資料處理！`);
  fs.writeFileSync(path.join(__dirname, "large_stores.json"), JSON.stringify(items, null, 2), "utf8");
}
