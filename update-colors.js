import fs from "node:fs";
import path from "node:path";
import { fileURLToPath } from "node:url";

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);

// 讀取 large_stores.json
const stores = JSON.parse(fs.readFileSync(path.join(__dirname, "large_stores.json"), "utf8"));

// 將非 488 名單的顏色由橙色改為黃色
const updatedStores = stores.map(s => {
  if (s.subGroup === "非488名單") {
    return {
      ...s,
      color: "黃色",
      colorHex: "#f59e0b"
    };
  }
  return s;
});

fs.writeFileSync(path.join(__dirname, "large_stores.json"), JSON.stringify(updatedStores, null, 2), "utf8");
console.log("large_stores.json 顏色更新完成 (橘色 -> 黃色)！");
