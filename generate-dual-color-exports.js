import fs from "node:fs";
import path from "node:path";
import { fileURLToPath } from "node:url";

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);

const CACHE_FILE = path.join(__dirname, "geocode_cache.json");
const cache = JSON.parse(fs.readFileSync(CACHE_FILE, "utf8"));

function cleanAddress(rawAddr) {
  let addr = rawAddr.trim();

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

// 讀取基準 488 筆清單集合
const list488 = fs.readFileSync(path.join(__dirname, "selected_stores_raw.txt"), "utf8")
  .split("\n")
  .map(l => l.trim())
  .filter(l => l.length > 0)
  .map(l => {
    const parts = l.split("\t");
    return { name: parts[0]?.trim() || "", address: parts[1]?.trim() || "" };
  })
  .filter(i => i.name && i.address);

const set488Key = new Set(list488.map(i => `${i.name}__${i.address}`));

// 讀取這次要標色的 500 筆清單
const inputLines = fs.readFileSync(path.join(__dirname, "input_for_color.txt"), "utf8")
  .split("\n")
  .map(l => l.trim())
  .filter(l => l.length > 0);

const stores = [];

for (let i = 0; i < inputLines.length; i++) {
  const line = inputLines[i];
  const parts = line.split("\t");
  const name = parts[0]?.trim();
  const address = parts[1]?.trim();
  if (!name || !address) continue;

  const { clean, estimated, defaultCoord } = cleanAddress(address);
  let coord = defaultCoord || cache[clean] || { lat: 23.9872, lng: 121.6016 };

  const isPrimary = set488Key.has(`${name}__${address}`);

  stores.push({
    id: `store-${stores.length + 1}`,
    name,
    address,
    lat: Number(coord.lat.toFixed(6)),
    lng: Number(coord.lng.toFixed(6)),
    estimated: estimated || defaultCoord != null,
    isPrimary488: isPrimary,
    group: isPrimary ? "主要488名單" : "非488名單(另色標記)"
  });
}

console.log(`總計 500 筆店家資料：`);
const primaryCount = stores.filter(s => s.isPrimary488).length;
const diffCount = stores.filter(s => !s.isPrimary488).length;
console.log(`- 主要 488 名單內: ${primaryCount} 筆 (綠色標記)`);
console.log(`- 非 488 名單內: ${diffCount} 筆 (橙色另色標記)`);

// 1. 輸出 JSON
fs.writeFileSync(path.join(__dirname, "large_stores.json"), JSON.stringify(stores, null, 2), "utf8");

// 2. 輸出 CSV (帶 UTF-8 BOM 與分組標記欄位)
const csvHeader = "\uFEFF序號,店名,地址,緯度,經度,分組分類,標記顏色,備註\n";
const csvRows = stores.map((s, idx) => {
  const note = s.estimated ? "位置為地址估算" : "";
  const safeName = `"${s.name.replace(/"/g, '""')}"`;
  const safeAddress = `"${s.address.replace(/"/g, '""')}"`;
  const groupName = s.isPrimary488 ? "主要488名單" : "非488名單(另色)";
  const colorName = s.isPrimary488 ? "綠色" : "橙色";
  return `${idx + 1},${safeName},${safeAddress},${s.lat},${s.lng},"${groupName}","${colorName}","${note}"`;
}).join("\n");

const csvContent = csvHeader + csvRows;
fs.writeFileSync(path.join(__dirname, "hualien_large_stores.csv"), csvContent, "utf8");
fs.writeFileSync(path.join(__dirname, "dist", "hualien_large_stores.csv"), csvContent, "utf8");
console.log("hualien_large_stores.csv 已更新！");

// 3. 輸出 KML (帶 Style 與雙 Folder 分組)
function makePlacemark(s, idx, styleId) {
  const note = s.estimated ? "<br/>※ 位置為地址估算" : "";
  const groupLabel = s.isPrimary488 ? "【主要488名單】" : "【★ 非488名單 - 另色標記】";
  return `      <Placemark>
        <name><![CDATA[${idx + 1}. ${s.name}]]></name>
        <description><![CDATA[${groupLabel}<br/>地址：${s.address}${note}]]></description>
        <styleUrl>#${styleId}</styleUrl>
        <Point>
          <coordinates>${s.lng},${s.lat},0</coordinates>
        </Point>
      </Placemark>`;
}

const primaryPlacemarks = stores
  .map((s, idx) => ({ s, idx }))
  .filter(({ s }) => s.isPrimary488)
  .map(({ s, idx }) => makePlacemark(s, idx, "style-primary"))
  .join("\n");

const diffPlacemarks = stores
  .map((s, idx) => ({ s, idx }))
  .filter(({ s }) => !s.isPrimary488)
  .map(({ s, idx }) => makePlacemark(s, idx, "style-diff"))
  .join("\n");

const kmlContent = `<?xml version="1.0" encoding="UTF-8"?>
<kml xmlns="http://www.opengis.net/kml/2.2">
  <Document>
    <name>花蓮店家地圖 (雙色分組標記 - 共 ${stores.length} 筆)</name>
    <description>綠色：主要 488 名單店家 (${primaryCount} 筆)；橙色：非 488 名單店家 (${diffCount} 筆)</description>
    
    <!-- 樣式1: 主要 488 店家 (綠色圖標) -->
    <Style id="style-primary">
      <IconStyle>
        <color>ff2d8a4e</color>
        <scale>1.1</scale>
        <Icon>
          <href>http://maps.google.com/mapfiles/ms/icons/green-dot.png</href>
        </Icon>
      </IconStyle>
    </Style>

    <!-- 樣式2: 非 488 店家 (橙色圖標) -->
    <Style id="style-diff">
      <IconStyle>
        <color>ff1e60e6</color>
        <scale>1.25</scale>
        <Icon>
          <href>http://maps.google.com/mapfiles/ms/icons/orange-dot.png</href>
        </Icon>
      </IconStyle>
    </Style>

    <Folder>
      <name>🟢 主要 488 店家 (${primaryCount} 筆 - 綠色標記)</name>
${primaryPlacemarks}
    </Folder>

    <Folder>
      <name>🟠 非 488 店家 (${diffCount} 筆 - 另色橙色標記)</name>
${diffPlacemarks}
    </Folder>

  </Document>
</kml>
`;

fs.writeFileSync(path.join(__dirname, "hualien_large_stores.kml"), kmlContent, "utf8");
fs.writeFileSync(path.join(__dirname, "dist", "hualien_large_stores.kml"), kmlContent, "utf8");
console.log("hualien_large_stores.kml 雙色分組版已更新！");
