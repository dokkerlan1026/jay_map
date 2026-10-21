import fs from "node:fs";
import path from "node:path";
import { fileURLToPath } from "node:url";

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);

// 讀取目前的 large_stores.json
const stores = JSON.parse(fs.readFileSync(path.join(__dirname, "large_stores.json"), "utf8"));

// 林森路+德安一街分界線參數:
// 基準點1: 德安一街西南端 (lat: 23.970, lng: 121.5925)
// 基準點2: 林森路東北端 (lat: 23.983, lng: 121.6070)
const p1 = { lat: 23.970, lng: 121.5925 };
const p2 = { lat: 23.983, lng: 121.6070 };
const slope = (p2.lng - p1.lng) / (p2.lat - p1.lat); // 約 1.11538

// 對 388 筆主要店家依據相對於界線的水平偏移量排序，精準二分（各 194 筆）
const primaryStores = stores.filter(s => s.isPrimary488);
const nonPrimaryStores = stores.filter(s => !s.isPrimary488);

const sortedPrimary = primaryStores.map(s => {
  const lineLng = p1.lng + (s.lat - p1.lat) * slope;
  const delta = s.lng - lineLng;
  return { ...s, delta };
}).sort((a, b) => a.delta - b.delta);

// 前 194 筆為界線以西（西區），後 194 筆為界線以東（東區）
const westStores = sortedPrimary.slice(0, 194).map(s => ({
  ...s,
  subGroup: "主要名單-西區",
  region: "界線以西",
  color: "紅色",
  colorHex: "#e53e3e"
}));

const eastStores = sortedPrimary.slice(194).map(s => ({
  ...s,
  subGroup: "主要名單-東區",
  region: "界線以東",
  color: "綠色",
  colorHex: "#257a4e"
}));

const diffStores = nonPrimaryStores.map(s => ({
  ...s,
  subGroup: "非488名單",
  region: "非488店家",
  color: "橙色",
  colorHex: "#d9531e"
}));

// 依原本 id 順序重新組合完整 500 筆
const fullStoresMap = new Map();
[...westStores, ...eastStores, ...diffStores].forEach(s => fullStoresMap.set(s.id, s));
const updatedStores = stores.map(s => fullStoresMap.get(s.id));

console.log("=== 分組結果 ===");
console.log(`- 🟢 主要名單-東區 (界線以東，綠色): ${eastStores.length} 筆`);
console.log(`- 🔴 主要名單-西區 (界線以西，紅色): ${westStores.length} 筆`);
console.log(`- 🟠 非 488 名單店家 (亮橙色): ${diffStores.length} 筆`);
console.log(`- 總計: ${updatedStores.length} 筆`);

// 1. 寫回 large_stores.json
fs.writeFileSync(path.join(__dirname, "large_stores.json"), JSON.stringify(updatedStores, null, 2), "utf8");

// 2. 輸出 CSV (帶 UTF-8 BOM，欄位包含完整分類與顏色)
const csvHeader = "\uFEFF序號,店名,地址,緯度,經度,分組分類,區域,標記顏色,備註\n";
const csvRows = updatedStores.map((s, idx) => {
  const note = s.estimated ? "位置為地址估算" : "";
  const safeName = `"${s.name.replace(/"/g, '""')}"`;
  const safeAddress = `"${s.address.replace(/"/g, '""')}"`;
  return `${idx + 1},${safeName},${safeAddress},${s.lat},${s.lng},"${s.subGroup}","${s.region}","${s.color}","${note}"`;
}).join("\n");

const csvContent = csvHeader + csvRows;
fs.writeFileSync(path.join(__dirname, "hualien_large_stores.csv"), csvContent, "utf8");
fs.writeFileSync(path.join(__dirname, "dist", "hualien_large_stores.csv"), csvContent, "utf8");
console.log("hualien_large_stores.csv 已更新！");

// 3. 輸出 KML (帶 3 種 Style 與 3 個獨立 Folder)
function makePlacemark(s, idx, styleId, badgeText) {
  const note = s.estimated ? "<br/>※ 位置為地址估算" : "";
  return `      <Placemark>
        <name><![CDATA[${idx + 1}. ${s.name}]]></name>
        <description><![CDATA[【${badgeText}】<br/>地址：${s.address}${note}]]></description>
        <styleUrl>#${styleId}</styleUrl>
        <Point>
          <coordinates>${s.lng},${s.lat},0</coordinates>
        </Point>
      </Placemark>`;
}

const eastPlacemarks = updatedStores
  .map((s, idx) => ({ s, idx }))
  .filter(({ s }) => s.subGroup === "主要名單-東區")
  .map(({ s, idx }) => makePlacemark(s, idx, "style-east-green", "🟢 主要名單 - 界線以東"))
  .join("\n");

const westPlacemarks = updatedStores
  .map((s, idx) => ({ s, idx }))
  .filter(({ s }) => s.subGroup === "主要名單-西區")
  .map(({ s, idx }) => makePlacemark(s, idx, "style-west-red", "🔴 主要名單 - 界線以西"))
  .join("\n");

const diffPlacemarks = updatedStores
  .map((s, idx) => ({ s, idx }))
  .filter(({ s }) => s.subGroup === "非488名單")
  .map(({ s, idx }) => makePlacemark(s, idx, "style-diff-orange", "🟠 非 488 名單店家"))
  .join("\n");

const kmlContent = `<?xml version="1.0" encoding="UTF-8"?>
<kml xmlns="http://www.opengis.net/kml/2.2">
  <Document>
    <name>花蓮店家地圖 (東區綠色 / 西區紅色 / 非488橙色 - 共 ${updatedStores.length} 筆)</name>
    <description>界線以東：綠色標記 (${eastStores.length} 筆)；界線以西：紅色標記 (${westStores.length} 筆)；非488名單：橙色標記 (${diffStores.length} 筆)</description>
    
    <!-- 樣式1: 主要東區 (綠色圖標) -->
    <Style id="style-east-green">
      <IconStyle>
        <color>ff2d8a4e</color>
        <scale>1.1</scale>
        <Icon>
          <href>http://maps.google.com/mapfiles/ms/icons/green-dot.png</href>
        </Icon>
      </IconStyle>
    </Style>

    <!-- 樣式2: 主要西區 (紅色圖標) -->
    <Style id="style-west-red">
      <IconStyle>
        <color>ff2020dd</color>
        <scale>1.1</scale>
        <Icon>
          <href>http://maps.google.com/mapfiles/ms/icons/red-dot.png</href>
        </Icon>
      </IconStyle>
    </Style>

    <!-- 樣式3: 非 488 店家 (橙色圖標) -->
    <Style id="style-diff-orange">
      <IconStyle>
        <color>ff1e60e6</color>
        <scale>1.25</scale>
        <Icon>
          <href>http://maps.google.com/mapfiles/ms/icons/orange-dot.png</href>
        </Icon>
      </IconStyle>
    </Style>

    <Folder>
      <name>🟢 主要名單 - 界線以東 (${eastStores.length} 筆 - 綠色標記)</name>
${eastPlacemarks}
    </Folder>

    <Folder>
      <name>🔴 主要名單 - 界線以西 (${westStores.length} 筆 - 紅色標記)</name>
${westPlacemarks}
    </Folder>

    <Folder>
      <name>🟠 非 488 名單店家 (${diffStores.length} 筆 - 橙色標記)</name>
${diffPlacemarks}
    </Folder>

  </Document>
</kml>
`;

fs.writeFileSync(path.join(__dirname, "hualien_large_stores.kml"), kmlContent, "utf8");
fs.writeFileSync(path.join(__dirname, "dist", "hualien_large_stores.kml"), kmlContent, "utf8");
console.log("hualien_large_stores.kml 三色分區版已更新！");
