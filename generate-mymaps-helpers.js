import fs from "node:fs";
import path from "node:path";
import { fileURLToPath } from "node:url";

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);

// 讀取 large_stores.json
const stores = JSON.parse(fs.readFileSync(path.join(__dirname, "large_stores.json"), "utf8"));

const eastStores = stores.filter(s => s.subGroup === "主要名單-東區");
const westStores = stores.filter(s => s.subGroup === "主要名單-西區");
const diffStores = stores.filter(s => s.subGroup === "非488名單");

console.log(`讀取資料: 東區 ${eastStores.length} 筆, 西區 ${westStores.length} 筆, 非488 ${diffStores.length} 筆`);

// 1. 生成 Google My Maps 專用 CSV (最乾淨的欄位結構)
// 第一欄直接命名為「店名」，包含序號與名稱，確保 Google My Maps 直接當作標題
const csvHeader = "\uFEFF店名,地址,分區,標記顏色,緯度,經度,備註\n";
const csvRows = stores.map((s, idx) => {
  const title = `"${idx + 1}. ${s.name.replace(/"/g, '""')}"`;
  const addr = `"${s.address.replace(/"/g, '""')}"`;
  const group = `"${s.subGroup}"`;
  const color = `"${s.color}"`;
  const note = s.estimated ? '"位置為地址估算"' : '""';
  return `${title},${addr},${group},${color},${s.lat},${s.lng},${note}`;
}).join("\n");

const mymapsCsv = csvHeader + csvRows;
fs.writeFileSync(path.join(__dirname, "hualien_google_mymaps.csv"), mymapsCsv, "utf8");
fs.writeFileSync(path.join(__dirname, "dist", "hualien_google_mymaps.csv"), mymapsCsv, "utf8");
console.log("hualien_google_mymaps.csv 已生成！");

// 2. 生成 3 個獨立的 KML 檔案（供 Google My Maps 分層匯入）
function buildSingleKml(storeList, layerName, colorHexKml, iconUrl, badgePrefix) {
  const placemarks = storeList.map(s => {
    const origIdx = stores.findIndex(x => x.id === s.id) + 1;
    const estNote = s.estimated ? "<br/>※ 位置為地址估算" : "";
    return `    <Placemark>
      <name><![CDATA[${origIdx}. ${s.name}]]></name>
      <description><![CDATA[【${badgePrefix}】<br/>地址：${s.address}${estNote}]]></description>
      <styleUrl>#layerStyle</styleUrl>
      <ExtendedData>
        <Data name="店名"><value><![CDATA[${s.name}]]></value></Data>
        <Data name="地址"><value><![CDATA[${s.address}]]></value></Data>
        <Data name="分區"><value><![CDATA[${badgePrefix}]]></value></Data>
      </ExtendedData>
      <Point>
        <coordinates>${s.lng},${s.lat},0</coordinates>
      </Point>
    </Placemark>`;
  }).join("\n");

  return `<?xml version="1.0" encoding="UTF-8"?>
<kml xmlns="http://www.opengis.net/kml/2.2">
  <Document>
    <name>${layerName}</name>
    <Style id="layerStyle">
      <IconStyle>
        <color>${colorHexKml}</color>
        <scale>1.1</scale>
        <Icon>
          <href>${iconUrl}</href>
        </Icon>
      </IconStyle>
    </Style>
${placemarks}
  </Document>
</kml>
`;
}

// 🟢 東區主要名單 (194 筆 - 綠色)
const eastKml = buildSingleKml(
  eastStores,
  `🟢 主要名單-東區 (${eastStores.length}筆 - 界線以東)`,
  "ff2d8a4e",
  "http://maps.google.com/mapfiles/kml/paddle/grn-blank.png",
  "🟢 主要名單 - 界線以東"
);
fs.writeFileSync(path.join(__dirname, "hualien_east_green.kml"), eastKml, "utf8");
fs.writeFileSync(path.join(__dirname, "dist", "hualien_east_green.kml"), eastKml, "utf8");

// 🔴 西區主要名單 (194 筆 - 紅色)
const westKml = buildSingleKml(
  westStores,
  `🔴 主要名單-西區 (${westStores.length}筆 - 界線以西)`,
  "ff2020dd",
  "http://maps.google.com/mapfiles/kml/paddle/red-blank.png",
  "🔴 主要名單 - 界線以西"
);
fs.writeFileSync(path.join(__dirname, "hualien_west_red.kml"), westKml, "utf8");
fs.writeFileSync(path.join(__dirname, "dist", "hualien_west_red.kml"), westKml, "utf8");

// 🟠 非 488 店家 (112 筆 - 橙色)
const diffKml = buildSingleKml(
  diffStores,
  `🟠 非488名單 (${diffStores.length}筆 - 橙色標記)`,
  "ff1e60e6",
  "http://maps.google.com/mapfiles/kml/paddle/orange-blank.png",
  "🟠 非 488 名單店家"
);
fs.writeFileSync(path.join(__dirname, "hualien_diff_orange.kml"), diffKml, "utf8");
fs.writeFileSync(path.join(__dirname, "dist", "hualien_diff_orange.kml"), diffKml, "utf8");

console.log("3 個獨立分區 KML 檔案已生成！");
