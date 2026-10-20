# 花蓮店家資料更新 Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** 以 25 筆花蓮店家完整取代台北示範資料，更新網站文案與地圖視野，並讓最佳路線規劃可在 25 個目的地上快速完成。

**Architecture:** 保留現有靜態網站與 `createRoutePlan` 介面。小型資料仍使用全排列取得精確結果；超過 9 個目的地時改用最近鄰居建立初始路線，再以開放路線 2-opt 改善，讓 UI 不需要知道演算法差異。

**Tech Stack:** HTML5、CSS、原生 JavaScript ES modules、Leaflet、OpenStreetMap、OSRM、Node.js built-in test runner、Sites 靜態部署。

---

### Task 1: 新增可擴充的路線演算法

**Files:**
- Modify: `tests/route-planner.test.js`
- Modify: `dist/route-planner.js`

- [ ] **Step 1: 先寫不存在函式的失敗測試**

在測試匯入 `buildScalableRouteOrder`，加入 26 點矩陣測試：

```js
test("buildScalableRouteOrder visits 25 destinations once without returning", () => {
  const size = 26;
  const matrix = Array.from({ length: size }, (_, from) =>
    Array.from({ length: size }, (_, to) => Math.abs(from - to)),
  );
  const order = buildScalableRouteOrder(matrix);
  assert.equal(order.length, size);
  assert.equal(order[0], 0);
  assert.equal(new Set(order).size, size);
  assert.deepEqual([...order].sort((a, b) => a - b), [...Array(size).keys()]);
  assert.notEqual(order.at(-1), 0);
});
```

- [ ] **Step 2: 執行測試並確認失敗**

Run: `node --test tests/route-planner.test.js`

Expected: FAIL，指出 `buildScalableRouteOrder` 尚未匯出。

- [ ] **Step 3: 實作最近鄰居與開放路線 2-opt**

`buildScalableRouteOrder(matrix)` 固定從索引 0 出發，每步選擇尚未拜訪且距離最短的節點；再檢查反轉 `order[i..j]` 是否降低 `matrix[order[i - 1]][order[i]] + matrix[order[j]][order[j + 1]]`，末段沒有回起點成本。最多執行 100 輪改善。

`findOptimalOrder(matrix)` 在 `matrix.length <= 10` 時保留原全排列，否則呼叫 `buildScalableRouteOrder`。

- [ ] **Step 4: 執行測試並確認通過**

Run: `npm test`

Expected: 8 tests PASS，0 FAIL。

- [ ] **Step 5: 提交演算法變更**

```powershell
git add dist/route-planner.js tests/route-planner.test.js
git commit -m "feat: scale route planning to 25 stores"
```

### Task 2: 以 25 筆花蓮資料取代示範資料

**Files:**
- Modify: `tests/site-shell.test.js`
- Modify: `dist/stores.js`

- [ ] **Step 1: 先寫精確資料集的失敗測試**

測試建立以下 `expectedStores` 陣列並檢查：

```js
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
assert.equal(stores.length, 25);
assert.deepEqual(stores.map(({ name, address }) => [name, address]), expectedStores);
assert.equal(new Set(stores.map(({ id }) => id)).size, 25);
assert.ok(stores.every(({ lat, lng }) => Number.isFinite(lat) && Number.isFinite(lng)));
assert.equal(stores.filter(({ address }) => address === "花蓮縣花蓮市中原路738號").length, 2);
assert.equal(HUALIEN_STATION.name, "花蓮車站");
```

- [ ] **Step 2: 執行測試並確認原 5 筆資料造成失敗**

Run: `node --test tests/site-shell.test.js`

Expected: FAIL，顯示實際長度 5、不符合預期 25。

- [ ] **Step 3: 寫入已查核座標與花蓮車站 fallback**

更新 `dist/stores.js`：使用者原始店名／地址逐字保留；採用 ArcGIS 地址候選座標。`百悅休閒飯店有限公司` 以新港街 80 號附近座標 `23.980811, 121.605509` 並標記 `estimated: true`；`恆康診所-溫控` 使用 StreetAddress 結果並標記估算；`三傢一活蝦料理` 依查核後的「南濱路32號」座標 `23.969325, 121.610974`。其餘使用 PointAddress 結果。花蓮車站 fallback 設為 `23.9929, 121.6013`。

- [ ] **Step 4: 執行測試並確認通過**

Run: `npm test`

Expected: 全部測試 PASS。

- [ ] **Step 5: 提交資料變更**

```powershell
git add dist/stores.js tests/site-shell.test.js
git commit -m "feat: replace samples with Hualien stores"
```

### Task 3: 更新花蓮文案與估算位置提示

**Files:**
- Modify: `tests/site-shell.test.js`
- Modify: `dist/index.html`
- Modify: `dist/app.js`

- [ ] **Step 1: 先寫文案與估算提示的失敗測試**

```js
assert.match(html, /花蓮店家地圖/);
assert.doesNotMatch(html, /台北門市地圖/);
assert.match(appSource, /位置為地址估算/);
```

- [ ] **Step 2: 執行測試並確認舊台北文案造成失敗**

Run: `node --test tests/site-shell.test.js`

Expected: FAIL，指出找不到「花蓮店家地圖」。

- [ ] **Step 3: 更新頁面文案、fallback 匯入與 popup**

`dist/index.html` 的 title、description、主標、說明與 eyebrow 改為花蓮店家語境。`dist/app.js` 將 `TAIPEI_MAIN_STATION` 改為 `HUALIEN_STATION`，初始視野改為 `[23.98, 121.6]`；popup 在 `store.estimated` 為 true 時加入 `<span class="popup-estimate">位置為地址估算</span>`，定位失敗訊息改為「改由花蓮車站出發規劃」。

- [ ] **Step 4: 更新估算提示樣式**

`dist/styles.css` 加入 `.popup-estimate`，使用琥珀色、較小字級與上方間距，讓提示清楚但不壓過地址。

- [ ] **Step 5: 執行完整測試與語法檢查**

Run: `npm test; node --check dist/app.js; node --check dist/route-planner.js; node --check dist/stores.js`

Expected: 全部測試 PASS，語法檢查 exit code 0。

- [ ] **Step 6: 提交介面更新**

```powershell
git add dist/index.html dist/app.js dist/styles.css tests/site-shell.test.js
git commit -m "feat: update map experience for Hualien"
```

### Task 4: 瀏覽器驗證並重新發佈既有 Site

**Files:**
- Preserve: `.openai/hosting.json`
- Modify: source only if browser QA reveals a tested defect

- [ ] **Step 1: 驗證 OSRM 可處理 26 個座標**

使用 Node `fetch` 對現有資料組成的 OSRM table URL 發送一次唯讀請求，檢查 `code === "Ok"` 且距離矩陣為 26×26。

- [ ] **Step 2: 啟動本機預覽並做瀏覽器 QA**

在桌面與 390px 手機 viewport 檢查「花蓮店家地圖」、25 列、25 個 marker、兩筆提拉米蘇、估算 popup、鍵盤選取及花蓮車站 fallback 路線。若瀏覽器因 OSRM 公開服務限制無法完成 26 點路線，確認錯誤訊息與單店導航仍正常。

- [ ] **Step 3: 取得既有 Site 的新 source credential**

以 `.openai/hosting.json` 中的 `project_id` 呼叫 `create_source_repository_write_credential`，不建立新 Site、不更改私人 audience。

- [ ] **Step 4: 推送與封裝相同 source commit**

使用 `site-workflow.mjs` 推送來源；若 Windows 路徑再次使官方 Bash wrapper 失敗，保留已推送 commit，改用官方 `prepare-site-build.cjs` 加系統 `tar` 產生包含 `dist/.openai/hosting.json` 的 archive。

- [ ] **Step 5: 儲存並部署私人版本**

將推送後的完整 `commit_sha` 與未變更 archive 傳給 `save_version_and_deploy_private`，確認回傳 `status: "succeeded"` 與原網址 `https://taipei-store-route-map.hanji001.chatgpt.site`。

- [ ] **Step 6: 最終驗證與清理**

Run: `npm test; node --check dist/app.js; node --check dist/route-planner.js; node --check dist/stores.js; git diff --check; git status --short`

Expected: 全部測試 PASS、語法檢查成功、Git 工作樹乾淨；刪除已確認位於 worktree 內的暫存 stage 與 archive，停止本機伺服器。
