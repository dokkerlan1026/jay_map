import test from "node:test";
import assert from "node:assert/strict";
import { readFile } from "node:fs/promises";

import { HUALIEN_STATION, stores } from "../dist/stores.js";

const expectedStores = [
  ["劉國鈞公館", "花蓮縣花蓮市明心街47號"],
  ["姚定國公館", "花蓮縣花蓮市中山路412巷2-2號"],
  ["天泉銀樓", "花蓮縣花蓮市中華路25號"],
  ["酒棧商行", "花蓮縣花蓮市建國路1段60號一樓"],
  ["佑昌企業社無人停車場", "花蓮縣花蓮市國聯一路18號"],
  ["提拉米蘇精緻蛋糕B棟", "花蓮縣花蓮市中原路738號"],
  ["提拉米蘇精緻蛋糕", "花蓮縣花蓮市中原路738號"],
  ["行政院東部聯合服務中心", "花蓮縣花蓮市中山路371號7-8樓"],
  ["仁和堂藥房", "花蓮縣花蓮市中山路316號"],
  ["彭貴春營業所", "花蓮縣花蓮市林森路243巷1號"],
  ["恆康診所-溫控", "花蓮縣花蓮市中原路711號"],
  ["繼陽建材行", "花蓮縣花蓮市重慶路386號"],
  ["日品模型玩具", "花蓮縣吉安鄉中原路1段131號"],
  ["莊謙信公館", "花蓮縣花蓮市復興街61號"],
  ["全人診所", "花蓮縣花蓮市中福路158號"],
  ["百悅休閒飯店有限公司", "花蓮縣花蓮市新港街80巷5號"],
  ["大芳名產行", "花蓮縣花蓮市中山路301號"],
  ["元盛百貨股份有限公司", "花蓮縣花蓮市國興二街148號"],
  ["三傢一活蝦料理", "花蓮縣花蓮市南濱32號"],
  ["鄭啟聰公館", "花蓮縣吉安鄉自立路1段148巷18號"],
  ["大同綜合花蓮和平門市", "花蓮縣花蓮市和平路415號(415.417號)"],
  ["富甲里實業有限公司", "花蓮縣吉安鄉自強路184號1樓"],
  ["藶峰有限公司", "花蓮縣花蓮市國聯五路126號旁"],
  ["特欣行開發股份有限公司", "花蓮縣吉安鄉建國路1段137之2號"],
  ["全聯花蓮2110", "花蓮縣花蓮市國聯二路110號"],
];

test("page exposes the required interactive regions", async () => {
  const html = await readFile(
    new URL("../dist/index.html", import.meta.url),
    "utf8",
  );

  for (const id of [
    "store-list",
    "map",
    "plan-route",
    "route-status",
    "route-summary",
  ]) {
    assert.match(html, new RegExp(`id=["']${id}["']`));
  }
  assert.match(html, /aria-live=["']polite["']/);
  assert.match(html, /leaflet\.css/);
  assert.match(html, /leaflet\.js/);
  assert.match(html, /src=["']\.\/app\.js["']/);
});

test("Hualien dataset preserves all 25 names, addresses, and coordinates", () => {
  assert.equal(stores.length, 25);
  assert.deepEqual(
    stores.map(({ name, address }) => [name, address]),
    expectedStores,
  );
  assert.equal(new Set(stores.map(({ id }) => id)).size, 25);
  for (const store of stores) {
    assert.ok(store.name && store.address);
    assert.ok(Number.isFinite(store.lat) && Number.isFinite(store.lng));
  }
  assert.equal(
    stores.filter(({ address }) => address === "花蓮縣花蓮市中原路738號")
      .length,
    2,
  );
  assert.equal(HUALIEN_STATION.name, "花蓮車站");
});

test("hidden selection labels remain visually hidden", async () => {
  const css = await readFile(
    new URL("../dist/styles.css", import.meta.url),
    "utf8",
  );

  assert.match(
    css,
    /\.selection-copy\[hidden\]\s*\{[^}]*display:\s*none\s*!important;/s,
  );
});
