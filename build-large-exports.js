import fs from "node:fs";
import path from "node:path";
import { fileURLToPath } from "node:url";

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);

const jsonPath = path.join(__dirname, "large_stores.json");
if (!fs.existsSync(jsonPath)) {
  console.error("large_stores.json 尚未產生！");
  process.exit(1);
}

const stores = JSON.parse(fs.readFileSync(jsonPath, "utf8"));
console.log(`正在處理 ${stores.length} 筆店家資料輸出...`);

// 1. 生成 CSV (帶 UTF-8 BOM \uFEFF)
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
console.log("hualien_large_stores.csv 生成成功！");

// 2. 生成 KML (Google Earth / Google My Maps 標準格式)
// 注意：Google My Maps 匯入上限單一圖層為 2000 個標記，500 個標記完全在限制內！
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
    <name>花蓮大規模店家地圖 (共 ${stores.length} 筆)</name>
    <description>花蓮市與吉安鄉完整 500 筆店家與地點清單</description>
${kmlPlacemarks}
  </Document>
</kml>
`;

fs.writeFileSync(path.join(__dirname, "hualien_large_stores.kml"), kmlContent, "utf8");
fs.writeFileSync(path.join(__dirname, "dist", "hualien_large_stores.kml"), kmlContent, "utf8");
console.log("hualien_large_stores.kml 生成成功！");
