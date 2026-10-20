import fs from "node:fs";
import path from "node:path";
import { fileURLToPath } from "node:url";
import { stores } from "./dist/stores.js";

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);

// 1. 生成 CSV (帶 UTF-8 BOM \uFEFF)
const csvHeader = "\uFEFF名稱,地址,緯度,經度,備註\n";
const csvRows = stores.map((s) => {
  const note = s.estimated ? "位置為地址估算" : "";
  // 避免欄位內含逗號
  const safeName = `"${s.name.replace(/"/g, '""')}"`;
  const safeAddress = `"${s.address.replace(/"/g, '""')}"`;
  return `${safeName},${safeAddress},${s.lat},${s.lng},"${note}"`;
}).join("\n");

const csvContent = csvHeader + csvRows;
const csvPath = path.join(__dirname, "hualien_stores.csv");
fs.writeFileSync(csvPath, csvContent, "utf8");
console.log(`CSV 檔案已生成: ${csvPath}`);

// 2. 生成 KML (Google Earth / Google My Maps 標準格式)
// 注意：KML 座標順序為 經度(lng),緯度(lat),高度(0)
const kmlPlacemarks = stores.map((s) => {
  const note = s.estimated ? "<br/>※ 位置為地址估算" : "";
  return `    <Placemark>
      <name><![CDATA[${s.name}]]></name>
      <description><![CDATA[地址：${s.address}${note}]]></description>
      <Point>
        <coordinates>${s.lng},${s.lat},0</coordinates>
      </Point>
    </Placemark>`;
}).join("\n");

const kmlContent = `<?xml version="1.0" encoding="UTF-8"?>
<kml xmlns="http://www.opengis.net/kml/2.2">
  <Document>
    <name>花蓮店家地圖</name>
    <description>花蓮市與吉安鄉共 25 筆店家與地點資料</description>
${kmlPlacemarks}
  </Document>
</kml>
`;

const kmlPath = path.join(__dirname, "hualien_stores.kml");
fs.writeFileSync(kmlPath, kmlContent, "utf8");
console.log(`KML 檔案已生成: ${kmlPath}`);
