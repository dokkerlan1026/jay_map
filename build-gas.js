import fs from "node:fs";
import path from "node:path";
import { fileURLToPath } from "node:url";

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);

const jsonPath = path.join(__dirname, "large_stores.json");
const stores = JSON.parse(fs.readFileSync(jsonPath, "utf8"));

const gasDir = path.join(__dirname, "gas");
if (!fs.existsSync(gasDir)) {
  fs.mkdirSync(gasDir, { recursive: true });
}

// ==========================================
// 1. 生成 gas/Code.gs (後端 Apps Script)
// ==========================================
const codeGsContent = `/**
 * Google Apps Script - 花蓮店家每日開關排程與地圖 Web App
 * 支援功能：
 * 1. Google 試算表多人共同編輯、打勾控制店家「顯示/關閉」
 * 2. 每日獨立紀錄，隔天可一鍵「載入全開啟」或「載入昨天/指定日期」
 * 3. 試算表與地圖 Web App 雙向同步連動
 */

// ------------------------------------------
// Web App 進入點 (doGet)
// ------------------------------------------
function doGet(e) {
  var template = HtmlService.createTemplateFromFile('Index');
  
  // 取得目標日期（未指定則為今日 YYYY-MM-DD）
  var targetDate = (e && e.parameter && e.parameter.date) ? e.parameter.date : getTodayString();
  template.currentDate = targetDate;
  
  // 載入該日期的店家資料與開關狀態
  var storeData = getStoresForDate(targetDate);
  template.storesJson = JSON.stringify(storeData);
  
  return template.evaluate()
    .setTitle('jay的地圖 - 花蓮店家每日排程地圖 (' + targetDate + ')')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1.0')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// ------------------------------------------
// 試算表自訂選單 (多人共同編輯與快速載入)
// ------------------------------------------
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('🗺️ 店家地圖與每日排程')
    .addItem('🟢 1. 建立/初始化 500 筆店家工作表', 'setupSheetWithData')
    .addSeparator()
    .addItem('☀️ 2. 【隔天操作】載入全開啟 (全部勾選顯示)', 'menuLoadAllActive')
    .addItem('📋 3. 【複製前日】載入昨天的開關狀態', 'menuLoadYesterday')
    .addItem('📅 4. 【指定載入】載入指定日期的開關狀態', 'menuLoadCustomDate')
    .addItem('💾 5. 【手動存檔】儲存今日開關狀態至歷史紀錄', 'menuSaveCurrentToHistory')
    .addSeparator()
    .addItem('🌐 6. 開啟地圖 Web App 網頁', 'openWebAppUrl')
    .addToUi();
}

// ------------------------------------------
// 日期工具函式
// ------------------------------------------
function getTodayString() {
  return Utilities.formatDate(new Date(), "Asia/Taipei", "yyyy-MM-dd");
}

function getYesterdayString(dateStr) {
  var d = dateStr ? new Date(dateStr) : new Date();
  d.setDate(d.getDate() - 1);
  return Utilities.formatDate(d, "Asia/Taipei", "yyyy-MM-dd");
}

// ------------------------------------------
// 取得指定日期的店家資料與開關狀態
// ------------------------------------------
function getStoresForDate(dateStr) {
  var defaultList = getDefaultStores();
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) return defaultList;

  var sheet = ss.getSheetByName("當日店家清單");
  if (!sheet) return defaultList;

  // 檢查試算表頂部的日期是否與查詢日期相符
  var sheetDate = String(sheet.getRange("B1").getValue()).trim();
  
  // 讀取當前工作表中的開關狀態
  var data = sheet.getDataRange().getValues();
  if (data.length <= 3) return defaultList;

  // 若目前工作表不是該日期的資料，嘗試從歷史紀錄抓取該日期的狀態
  var hiddenStoreIds = new Set();
  if (sheetDate === dateStr) {
    for (var i = 3; i < data.length; i++) {
      var row = data[i];
      var isVisible = (row[1] === true || String(row[1]).toUpperCase() === "TRUE");
      var storeId = row[0]; // 序號
      if (!isVisible) {
        hiddenStoreIds.add(Number(storeId));
      }
    }
  } else {
    // 從歷史表讀取
    hiddenStoreIds = getHiddenStoreIdsFromHistory(dateStr);
  }

  // 套用狀態
  return defaultList.map(function(s, idx) {
    var num = idx + 1;
    return Object.assign({}, s, {
      isVisible: !hiddenStoreIds.has(num)
    });
  });
}

// ------------------------------------------
// 歷史紀錄處理函式
// ------------------------------------------
function getHiddenStoreIdsFromHistory(dateStr) {
  var hiddenSet = new Set();
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) return hiddenSet;

  var histSheet = ss.getSheetByName("歷史紀錄存檔");
  if (!histSheet) return hiddenSet;

  var data = histSheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    var rowDate = Utilities.formatDate(new Date(data[i][0]), "Asia/Taipei", "yyyy-MM-dd");
    if (rowDate === dateStr) {
      var hiddenListStr = String(data[i][1] || "");
      if (hiddenListStr) {
        hiddenListStr.split(",").forEach(function(idStr) {
          var n = Number(idStr.trim());
          if (!isNaN(n)) hiddenSet.add(n);
        });
      }
      break;
    }
  }
  return hiddenSet;
}

// 儲存目前試算表的狀態至歷史紀錄
function saveCurrentStatusToHistory(customDate) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) return;

  var sheet = ss.getSheetByName("當日店家清單");
  if (!sheet) return;

  var dateStr = customDate || String(sheet.getRange("B1").getValue()).trim() || getTodayString();
  var data = sheet.getDataRange().getValues();

  var hiddenIds = [];
  var activeCount = 0;
  for (var i = 3; i < data.length; i++) {
    var row = data[i];
    var isVisible = (row[1] === true || String(row[1]).toUpperCase() === "TRUE");
    var storeId = row[0];
    if (!isVisible) {
      hiddenIds.push(storeId);
    } else {
      activeCount++;
    }
  }

  var histSheet = ss.getSheetByName("歷史紀錄存檔");
  if (!histSheet) {
    histSheet = ss.insertSheet("歷史紀錄存檔");
    histSheet.appendRow(["日期", "關閉的店家序號(逗號分隔)", "開啟總數", "關閉總數", "最後存檔時間"]);
    histSheet.getRange(1, 1, 1, 5).setFontWeight("bold").setBackground("#e8f0fe");
  }

  // 尋找是否已有該日期的紀錄，有則更新，無則新增
  var histData = histSheet.getDataRange().getValues();
  var foundRow = -1;
  for (var j = 1; j < histData.length; j++) {
    var d = Utilities.formatDate(new Date(histData[j][0]), "Asia/Taipei", "yyyy-MM-dd");
    if (d === dateStr) {
      foundRow = j + 1;
      break;
    }
  }

  var nowStr = Utilities.formatDate(new Date(), "Asia/Taipei", "yyyy-MM-dd HH:mm:ss");
  var rowValues = [dateStr, hiddenIds.join(","), activeCount, hiddenIds.length, nowStr];

  if (foundRow > 0) {
    histSheet.getRange(foundRow, 1, 1, 5).setValues([rowValues]);
  } else {
    histSheet.appendRow(rowValues);
  }
}

// ------------------------------------------
// 核心操作：隔天載入全開 / 載入昨天 / 載入指定日期
// ------------------------------------------

// 1. 載入全開啟（全部勾選顯示）
function loadAllActive(targetDate) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) throw new Error("請先在 Google 試算表中執行！");

  var sheet = ss.getSheetByName("當日店家清單");
  if (!sheet) throw new Error("請先建立「當日店家清單」！");

  var dateToSet = targetDate || getTodayString();
  sheet.getRange("B1").setValue(dateToSet);

  var lastRow = sheet.getLastRow();
  if (lastRow >= 4) {
    var numRows = lastRow - 3;
    var trueValues = [];
    for (var i = 0; i < numRows; i++) {
      trueValues.push([true]);
    }
    sheet.getRange(4, 2, numRows, 1).setValues(trueValues);
    sheet.getRange("D1").setValue("已顯示: " + numRows + " / " + numRows + " (全開啟)");
  }

  // 自動同步存檔到歷史
  saveCurrentStatusToHistory(dateToSet);
  return { success: true, date: dateToSet, activeCount: numRows };
}

// 2. 載入指定日期的狀態（若為昨天，由外部傳入昨天的日期字串）
function loadStatusFromDate(sourceDateStr, targetDateStr) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) throw new Error("請先在 Google 試算表中執行！");

  var sheet = ss.getSheetByName("當日店家清單");
  if (!sheet) throw new Error("請先建立「當日店家清單」！");

  var hiddenSet = getHiddenStoreIdsFromHistory(sourceDateStr);
  var dateToSet = targetDateStr || getTodayString();
  sheet.getRange("B1").setValue(dateToSet);

  var lastRow = sheet.getLastRow();
  if (lastRow >= 4) {
    var numRows = lastRow - 3;
    var values = [];
    var activeCount = 0;
    for (var i = 1; i <= numRows; i++) {
      var isVisible = !hiddenSet.has(i);
      if (isVisible) activeCount++;
      values.push([isVisible]);
    }
    sheet.getRange(4, 2, numRows, 1).setValues(values);
    sheet.getRange("D1").setValue("已顯示: " + activeCount + " / " + numRows + " (來自 " + sourceDateStr + ")");
  }

  // 自動同步更新今日歷史紀錄
  saveCurrentStatusToHistory(dateToSet);
  return { success: true, targetDate: dateToSet, sourceDate: sourceDateStr, activeCount: activeCount };
}

// ------------------------------------------
// 前端 Web App 呼叫的 API（雙向同步連動）
// ------------------------------------------

// 單點狀態更新（外勤在網頁打勾/取消，即時寫回 Sheet）
function apiUpdateStoreStatus(storeId, isVisible) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    if (!ss) return { success: false, msg: "No active sheet" };

    var sheet = ss.getSheetByName("當日店家清單");
    if (!sheet) return { success: false, msg: "Sheet not found" };

    var rowNumber = Number(storeId) + 3; // 第 4 列開始為序號 1
    sheet.getRange(rowNumber, 2).setValue(isVisible);
    
    // 更新備註時間
    var nowStr = Utilities.formatDate(new Date(), "Asia/Taipei", "HH:mm:ss");
    sheet.getRange(rowNumber, 10).setValue("網頁更新 " + nowStr);

    return { success: true };
  } catch (err) {
    return { success: false, error: String(err) };
  }
}

// 批次狀態更新
function apiBatchUpdateStatus(updates) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    if (!ss) return { success: false };

    var sheet = ss.getSheetByName("當日店家清單");
    if (!sheet) return { success: false };

    updates.forEach(function(u) {
      var row = Number(u.id) + 3;
      sheet.getRange(row, 2).setValue(u.isVisible);
    });

    saveCurrentStatusToHistory();
    return { success: true };
  } catch (err) {
    return { success: false, error: String(err) };
  }
}

// ------------------------------------------
// 試算表選單按鈕執行函式
// ------------------------------------------
function menuLoadAllActive() {
  var ui = SpreadsheetApp.getUi();
  var res = ui.alert("載入全開啟", "確定要將今天的店家狀態全部設為【開啟顯示】嗎？", ui.ButtonSet.YES_NO);
  if (res === ui.Button.YES) {
    var out = loadAllActive();
    ui.alert("成功！", "已成功將今天（" + out.date + "）所有店家設為開啟顯示！", ui.ButtonSet.OK);
  }
}

function menuLoadYesterday() {
  var ui = SpreadsheetApp.getUi();
  var today = getTodayString();
  var yesterday = getYesterdayString(today);

  var res = ui.alert("載入昨天的設定", "確定要載入昨天（" + yesterday + "）的開關狀態作為今天（" + today + "）的紀錄嗎？", ui.ButtonSet.YES_NO);
  if (res === ui.Button.YES) {
    var out = loadStatusFromDate(yesterday, today);
    ui.alert("載入完成！", "已載入昨日（" + yesterday + "）狀態至今日，目前開啟店家數：" + out.activeCount + " 家。", ui.ButtonSet.OK);
  }
}

function menuLoadCustomDate() {
  var ui = SpreadsheetApp.getUi();
  var prompt = ui.prompt("載入指定日期", "請輸入要載入的歷史日期 (格式：YYYY-MM-DD，例如 2026-10-20)：", ui.ButtonSet.OK_CANCEL);
  if (prompt.getSelectedButton() === ui.Button.OK) {
    var srcDate = prompt.getResponseText().trim();
    if (!/^\\d{4}-\\d{2}-\\d{2}$/.test(srcDate)) {
      ui.alert("錯誤", "日期格式不正確，請使用 YYYY-MM-DD 格式！", ui.ButtonSet.OK);
      return;
    }
    var today = getTodayString();
    var out = loadStatusFromDate(srcDate, today);
    ui.alert("載入完成！", "已成功載入 " + srcDate + " 的設定，目前開啟店家數：" + out.activeCount + " 家。", ui.ButtonSet.OK);
  }
}

function menuSaveCurrentToHistory() {
  var ui = SpreadsheetApp.getUi();
  var today = getTodayString();
  saveCurrentStatusToHistory(today);
  ui.alert("存檔成功！", "今日（" + today + "）的開關紀錄已成功備份至【歷史紀錄存檔】工作表！", ui.ButtonSet.OK);
}

function openWebAppUrl() {
  var url = ScriptApp.getService().getUrl();
  var ui = SpreadsheetApp.getUi();
  if (!url) {
    ui.alert("請先完成部署", "請點擊右上角【部署】>【新增部署作業】發佈為網頁應用程式後即可使用！", ui.ButtonSet.OK);
    return;
  }
  var html = '<div style="font-family:sans-serif;padding:12px;">' +
             '<p>點擊下方按鈕即可開啟地圖網頁（支援手機與電腦）：</p>' +
             '<p><a href="' + url + '" target="_blank" style="display:inline-block;padding:10px 18px;background:#1a73e8;color:white;text-decoration:none;border-radius:6px;font-weight:bold;">開啟店家地圖 Web App ↗</a></p>' +
             '<p style="color:#666;font-size:12px;">網址：<br/>' + url + '</p>' +
             '</div>';
  ui.showModalDialog(HtmlService.createHtmlOutput(html).setWidth(420).setHeight(220), "開啟地圖網頁");
}

// ------------------------------------------
// 一鍵建立 / 初始化 500 筆店家工作表
// ------------------------------------------
function setupSheetWithData() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) {
    throw new Error("請先將此 Apps Script 專案綁定至 Google 試算表（擴充功能 > Apps Script）");
  }

  var sheet = ss.getSheetByName("當日店家清單");
  if (!sheet) {
    sheet = ss.insertSheet("當日店家清單", 0);
  } else {
    sheet.clear();
  }

  var today = getTodayString();
  var stores = getDefaultStores();

  // 頂部控制與日期列
  sheet.getRange("A1").setValue("📅 排程日期：").setFontWeight("bold");
  sheet.getRange("B1").setValue(today).setFontWeight("bold").setFontColor("#1a73e8").setFontSize(12);
  sheet.getRange("C1").setValue("📊 狀態統計：").setFontWeight("bold");
  sheet.getRange("D1").setValue("已顯示: " + stores.length + " / " + stores.length).setFontWeight("bold");
  sheet.getRange("E1").setValue("💡 提示：勾選/取消 B 欄可立即切換地圖顯示；透過上方選單可載入昨天或全開！").setFontColor("#5f6368");

  // 表頭
  var headers = [["序號", "顯示開關", "店名", "地址", "分組分類", "區域", "標記顏色", "緯度", "經度", "備註"]];
  sheet.getRange(3, 1, 1, headers[0].length).setValues(headers);
  sheet.getRange(3, 1, 1, headers[0].length).setFontWeight("bold").setBackground("#c5221f").setFontColor("#ffffff");

  // 資料列（B 欄為原生 Checkbox）
  var rows = stores.map(function(s, idx) {
    return [
      idx + 1,
      true, // 預設全部顯示 (打勾)
      s.name,
      s.address,
      s.subGroup,
      s.region,
      s.color,
      s.lat,
      s.lng,
      s.estimated ? "位置為地址估算" : ""
    ];
  });

  sheet.getRange(4, 1, rows.length, rows[0].length).setValues(rows);

  // 在 B 欄建立漂亮的核取方塊 (Checkbox)
  var chkRange = sheet.getRange(4, 2, rows.length, 1);
  chkRange.insertCheckboxes();

  // 凍結前 3 列
  sheet.setFrozenRows(3);
  sheet.setFrozenColumns(3);

  // 欄寬美化
  sheet.setColumnWidth(1, 50);  // 序號
  sheet.setColumnWidth(2, 80);  // 顯示開關
  sheet.setColumnWidth(3, 200); // 店名
  sheet.setColumnWidth(4, 260); // 地址
  sheet.setColumnWidth(5, 140); // 分組分類

  // 同步初始化歷史表
  saveCurrentStatusToHistory(today);

  var ui = SpreadsheetApp.getUi();
  ui.alert("初始化完成！", "已成功建立【當日店家清單】工作表！\\n\\n1. B 欄已設定為核取方塊，打勾即可開啟/關閉顯示\\n2. 點擊上方選單【🗺️ 店家地圖與每日排程】可進行隔天載入或全開\\n3. 任何人共同編輯此試算表，地圖皆可即時連動！", ui.ButtonSet.OK);
}

// ------------------------------------------
// 內建 500 筆資料
// ------------------------------------------
function getDefaultStores() {
  return STORE_DATA_PLACEHOLDER;
}
`;

const fullCodeGs = codeGsContent.replace(
  "STORE_DATA_PLACEHOLDER",
  JSON.stringify(stores, null, 2)
);
fs.writeFileSync(path.join(gasDir, "Code.gs"), fullCodeGs, "utf8");
console.log("gas/Code.gs 已成功生成！");

// ==========================================
// 2. 生成 gas/Index.html (前端 Web App)
// 包含：頂部日期切換、隔天全開載入、複製昨天、多人共同編輯即時同步
// ==========================================
const eastStores = stores.filter(s => s.subGroup === "主要名單-東區");
const westStores = stores.filter(s => s.subGroup === "主要名單-西區");
const diffStores = stores.filter(s => s.subGroup === "非488名單");

const indexHtml = `<!doctype html>
<html lang="zh-Hant">
  <head>
    <meta charset="UTF-8" />
    <meta name="viewport" content="width=device-width, initial-scale=1.0" />
    <title>jay的地圖 - 花蓮店家每日排程地圖</title>
    <link rel="preconnect" href="https://fonts.googleapis.com" />
    <link rel="preconnect" href="https://fonts.gstatic.com" crossorigin />
    <link href="https://fonts.googleapis.com/css2?family=Noto+Sans+TC:wght@400;500;600;700;800&family=Roboto:wght@400;500;700&display=swap" rel="stylesheet" />
    <link rel="stylesheet" href="https://unpkg.com/leaflet@1.9.4/dist/leaflet.css" />
    <style>
      :root {
        --google-red: #c5221f;
        --google-hover: #f1f3f4;
        --color-east: #2e7d32;       /* 綠色 */
        --color-west: #d32f2f;       /* 紅色 */
        --color-diff: #e65100;       /* 橙色 */
        --border-color: #dadce0;
        --text-primary: #202124;
        --text-secondary: #5f6368;
        --sidebar-width: 380px;
      }
      * { box-sizing: border-box; margin: 0; padding: 0; }
      body {
        font-family: "Noto Sans TC", "Roboto", system-ui, sans-serif;
        background: #f8f9fa;
        color: var(--text-primary);
        height: 100vh;
        display: flex;
        flex-direction: column;
        overflow: hidden;
      }

      /* 頂部 Google 我的地圖經典紅色工具列 */
      header {
        background: var(--google-red);
        color: white;
        height: 56px;
        padding: 0 16px;
        display: flex;
        align-items: center;
        justify-content: space-between;
        gap: 12px;
        box-shadow: 0 2px 4px rgba(0,0,0,0.2);
        z-index: 1000;
        flex-shrink: 0;
      }
      .header-left {
        display: flex;
        align-items: center;
        gap: 12px;
      }
      .menu-btn {
        background: none;
        border: none;
        color: white;
        width: 36px;
        height: 36px;
        border-radius: 50%;
        display: flex;
        align-items: center;
        justify-content: center;
        cursor: pointer;
        transition: background 150ms;
      }
      .menu-btn:hover { background: rgba(255,255,255,0.18); }
      .brand-box {
        display: flex;
        flex-direction: column;
      }
      .brand-title {
        font-size: 1.15rem;
        font-weight: 700;
        letter-spacing: 0.3px;
        display: flex;
        align-items: center;
        gap: 8px;
      }
      .date-badge {
        font-size: 0.76rem;
        background: #ffffff;
        color: var(--google-red);
        padding: 2px 8px;
        border-radius: 12px;
        font-weight: 800;
      }
      .header-right {
        display: flex;
        align-items: center;
        gap: 8px;
      }
      .header-btn {
        background: rgba(255,255,255,0.18);
        color: white;
        border: 1px solid rgba(255,255,255,0.3);
        padding: 6px 10px;
        border-radius: 4px;
        font-size: 0.78rem;
        font-weight: 600;
        text-decoration: none;
        display: inline-flex;
        align-items: center;
        gap: 4px;
        cursor: pointer;
        transition: all 140ms ease;
      }
      .header-btn:hover {
        background: white;
        color: var(--google-red);
      }

      /* 每日排程快捷工具列 (Google Sheets 連動專用) */
      .schedule-bar {
        background: #ffffff;
        border-bottom: 1px solid var(--border-color);
        padding: 8px 16px;
        display: flex;
        align-items: center;
        justify-content: space-between;
        gap: 12px;
        flex-wrap: wrap;
        font-size: 0.82rem;
        box-shadow: 0 1px 2px rgba(0,0,0,0.04);
      }
      .schedule-left {
        display: flex;
        align-items: center;
        gap: 8px;
        flex-wrap: wrap;
      }
      .date-input {
        padding: 4px 8px;
        border: 1px solid var(--border-color);
        border-radius: 4px;
        font-size: 0.82rem;
        font-family: inherit;
        outline: none;
      }
      .sched-btn {
        padding: 5px 10px;
        border: 1px solid var(--border-color);
        background: #f8f9fa;
        border-radius: 4px;
        font-size: 0.78rem;
        font-weight: 600;
        cursor: pointer;
        color: var(--text-primary);
        display: inline-flex;
        align-items: center;
        gap: 4px;
        transition: all 120ms;
      }
      .sched-btn:hover {
        background: #e8f0fe;
        color: #1a73e8;
        border-color: #d2e3fc;
      }
      .sched-btn.action-primary {
        background: #e6f4ea;
        color: #137333;
        border-color: #ceead6;
      }
      .sched-btn.action-primary:hover {
        background: #ceead6;
      }

      /* 主版面：左側欄 + 右側地圖 */
      .layout-body {
        flex: 1;
        display: flex;
        position: relative;
        overflow: hidden;
      }

      /* 左側抽屜 / 面板 */
      .sidebar {
        width: var(--sidebar-width);
        background: white;
        border-right: 1px solid var(--border-color);
        display: flex;
        flex-direction: column;
        height: 100%;
        z-index: 500;
        box-shadow: 2px 0 6px rgba(0,0,0,0.06);
        transition: transform 240ms ease;
      }
      .sidebar.collapsed {
        transform: translateX(calc(-1 * var(--sidebar-width)));
      }

      /* 統計與過濾操作區 */
      .panel-intro {
        padding: 10px 16px;
        border-bottom: 1px solid var(--border-color);
        background: #fafafa;
      }
      .intro-meta {
        font-size: 0.78rem;
        color: var(--text-secondary);
        display: flex;
        align-items: center;
        justify-content: space-between;
        margin-bottom: 8px;
      }
      .quick-actions {
        display: flex;
        gap: 4px;
        flex-wrap: wrap;
      }
      .btn-chip {
        padding: 3px 8px;
        border: 1px solid var(--border-color);
        background: white;
        border-radius: 12px;
        font-size: 0.72rem;
        font-weight: 600;
        color: var(--text-secondary);
        cursor: pointer;
        transition: all 120ms;
      }
      .btn-chip:hover {
        background: var(--google-hover);
        color: var(--text-primary);
        border-color: #bdc1c6;
      }
      .btn-chip.primary {
        background: #e8f0fe;
        color: #1a73e8;
        border-color: #d2e3fc;
      }

      /* 搜尋列 */
      .search-box {
        padding: 8px 16px;
        border-bottom: 1px solid var(--border-color);
        background: #ffffff;
      }
      .search-input-wrap {
        position: relative;
        display: flex;
        align-items: center;
      }
      .search-input {
        width: 100%;
        padding: 7px 30px 7px 10px;
        border: 1px solid var(--border-color);
        border-radius: 6px;
        font-size: 0.86rem;
        font-family: inherit;
        outline: none;
      }
      .search-input:focus { border-color: #1a73e8; }
      .search-clear {
        position: absolute;
        right: 8px;
        background: none;
        border: none;
        font-size: 16px;
        color: #9aa0a6;
        cursor: pointer;
        display: none;
      }

      /* 圖層清單（樹狀結構） */
      .layer-tree {
        flex: 1;
        overflow-y: auto;
      }

      /* 單一圖層群組 */
      .layer-group {
        border-bottom: 1px solid var(--border-color);
      }
      .group-header {
        display: flex;
        align-items: center;
        padding: 10px 12px 10px 16px;
        background: #ffffff;
        cursor: pointer;
        user-select: none;
        transition: background 120ms;
        gap: 10px;
      }
      .group-header:hover {
        background: #f1f3f4;
      }
      .group-checkbox {
        width: 18px;
        height: 18px;
        cursor: pointer;
        accent-color: var(--google-red);
      }
      .group-icon {
        width: 12px;
        height: 12px;
        border-radius: 50%;
        flex-shrink: 0;
      }
      .group-icon.east { background: var(--color-east); }
      .group-icon.west { background: var(--color-west); }
      .group-icon.diff { background: var(--color-diff); }

      .group-title-wrap {
        flex: 1;
        display: flex;
        flex-direction: column;
      }
      .group-title {
        font-size: 0.9rem;
        font-weight: 700;
        color: var(--text-primary);
        display: flex;
        align-items: center;
        gap: 6px;
      }
      .group-stats {
        font-size: 0.72rem;
        color: var(--text-secondary);
      }
      .arrow-btn {
        width: 24px;
        height: 24px;
        display: flex;
        align-items: center;
        justify-content: center;
        font-size: 0.8rem;
        color: var(--text-secondary);
        transition: transform 180ms ease;
      }
      .layer-group.expanded .arrow-btn {
        transform: rotate(180deg);
      }

      .item-list {
        display: none;
        background: #ffffff;
      }
      .layer-group.expanded .item-list {
        display: block;
      }

      /* 單一店家項目（支援獨立勾選） */
      .store-row {
        display: flex;
        align-items: center;
        padding: 8px 12px 8px 24px;
        border-top: 1px solid #f1f3f4;
        cursor: pointer;
        transition: background 100ms;
        gap: 10px;
      }
      .store-row:hover {
        background: #f8f9fa;
      }
      .store-row.active {
        background: #e8f0fe;
      }
      .store-row.disabled {
        opacity: 0.42;
      }
      .item-checkbox {
        width: 16px;
        height: 16px;
        cursor: pointer;
        accent-color: var(--google-red);
        flex-shrink: 0;
      }
      .pin-svg {
        width: 14px;
        height: 18px;
        flex-shrink: 0;
      }
      .store-info {
        flex: 1;
        min-width: 0;
      }
      .store-name-text {
        font-size: 0.86rem;
        font-weight: 500;
        color: var(--text-primary);
        white-space: nowrap;
        overflow: hidden;
        text-overflow: ellipsis;
      }
      .store-addr-text {
        font-size: 0.72rem;
        color: var(--text-secondary);
        white-space: nowrap;
        overflow: hidden;
        text-overflow: ellipsis;
      }

      /* 側邊欄收合切換浮動按鈕 */
      .toggle-sidebar-btn {
        position: absolute;
        top: 14px;
        left: calc(var(--sidebar-width) + 12px);
        z-index: 999;
        background: white;
        border: 1px solid var(--border-color);
        box-shadow: 0 1px 4px rgba(0,0,0,0.18);
        border-radius: 4px;
        padding: 6px 10px;
        font-size: 0.8rem;
        font-weight: 600;
        cursor: pointer;
        display: flex;
        align-items: center;
        gap: 6px;
        transition: all 200ms ease;
      }
      .sidebar.collapsed ~ .toggle-sidebar-btn {
        left: 14px;
      }

      /* 右側 Leaflet 地圖容器 */
      #map {
        flex: 1;
        height: 100%;
        background: #f2efe9;
      }

      /* 自訂 Google 經典水滴圖釘 Marker */
      .gmarker-wrap {
        background: transparent;
        border: none;
      }
      .gmarker-svg {
        filter: drop-shadow(0 2px 3px rgba(0,0,0,0.35));
        transition: transform 120ms ease;
      }
      .gmarker-svg:hover {
        transform: scale(1.2) translateY(-2px);
      }

      /* Popup 美化 */
      .leaflet-popup-content-wrapper {
        border-radius: 8px;
        box-shadow: 0 2px 10px rgba(0,0,0,0.22);
        padding: 0;
        overflow: hidden;
      }
      .leaflet-popup-content {
        margin: 0;
        padding: 14px 16px;
        min-width: 240px;
        max-width: 300px;
        font-family: inherit;
      }
      .pop-header {
        display: flex;
        align-items: center;
        justify-content: space-between;
        margin-bottom: 6px;
      }
      .pop-badge {
        font-size: 0.7rem;
        font-weight: 700;
        padding: 2px 8px;
        border-radius: 12px;
      }
      .pop-badge.east { background: #e8f5e9; color: var(--color-east); }
      .pop-badge.west { background: #ffebee; color: var(--color-west); }
      .pop-badge.diff { background: #fff3e0; color: var(--color-diff); }

      .pop-title {
        font-size: 1rem;
        font-weight: 700;
        color: var(--text-primary);
        margin-bottom: 4px;
      }
      .pop-addr {
        font-size: 0.8rem;
        color: var(--text-secondary);
        margin-bottom: 12px;
        line-height: 1.4;
      }
      .pop-actions {
        display: flex;
        gap: 8px;
      }
      .pop-nav-btn {
        flex: 1;
        text-align: center;
        padding: 6px 10px;
        background: #1a73e8;
        color: white !important;
        text-decoration: none;
        border-radius: 4px;
        font-size: 0.8rem;
        font-weight: 600;
      }
      .pop-nav-btn:hover { background: #1557b0; }
      .pop-toggle-btn {
        padding: 6px 10px;
        background: #f1f3f4;
        border: 1px solid var(--border-color);
        border-radius: 4px;
        font-size: 0.78rem;
        font-weight: 600;
        cursor: pointer;
        color: var(--text-secondary);
      }

      /* 儲存中 / 同步中 Toast 提示 */
      .sync-toast {
        position: fixed;
        bottom: 20px;
        right: 20px;
        background: #323232;
        color: white;
        padding: 10px 16px;
        border-radius: 6px;
        font-size: 0.82rem;
        z-index: 2000;
        display: none;
        box-shadow: 0 3px 8px rgba(0,0,0,0.3);
      }

      @media (max-width: 768px) {
        .layout-body { flex-direction: column-reverse; }
        .sidebar { width: 100%; height: 50vh; }
        #map { width: 100%; height: 50vh; }
        .toggle-sidebar-btn { display: none; }
        .schedule-bar { font-size: 0.75rem; }
      }
    </style>
  </head>
  <body>
    <!-- 頂部 Google 我的地圖工具列 -->
    <header>
      <div class="header-left">
        <button class="menu-btn" id="menuToggleBtn" title="展開/收合側邊欄">
          <svg width="22" height="22" viewBox="0 0 24 24" fill="currentColor">
            <path d="M3 18h18v-2H3v2zm0-5h18v-2H3v2zm0-7v2h18V6H3z"/>
          </svg>
        </button>
        <div class="brand-box">
          <div class="brand-title">
            <span>jay的地圖</span>
            <span class="date-badge" id="headerDateLabel"><?= currentDate ?></span>
          </div>
        </div>
      </div>
      <div class="header-right">
        <button class="header-btn" id="btnSyncFromSheet" title="讀取 Google 試算表最新勾選狀態">
          <span>🔄 同步試算表</span>
        </button>
      </div>
    </header>

    <!-- 每日排程與範本快速載入列 -->
    <div class="schedule-bar">
      <div class="schedule-left">
        <span>📅 日期：</span>
        <input type="date" id="dateSelector" class="date-input" value="<?= currentDate ?>" />
        <button class="sched-btn action-primary" id="btnLoadAllOpen" title="隔天快速載入全部打勾開啟">
          <span>☀️ 載入全開啟</span>
        </button>
        <button class="sched-btn" id="btnLoadYesterday" title="複製昨天的開關狀態到今天">
          <span>📋 載入昨天的設定</span>
        </button>
        <button class="sched-btn" id="btnLoadCustom" title="載入選取日期的歷史設定">
          <span>📅 載入所選日期</span>
        </button>
      </div>
      <div style="color:var(--text-secondary);font-size:0.75rem;">
        <span>💡 共同編輯：在試算表打勾，地圖自動同步！</span>
      </div>
    </div>

    <div class="layout-body">
      <aside class="sidebar" id="sidebar">
        <!-- 頂部資訊與快速操作按鈕 -->
        <div class="panel-intro">
          <div class="intro-meta">
            <span id="activeCountLabel">目前顯示：計算中...</span>
            <span>以林森路+德安一街為界</span>
          </div>
          <div class="quick-actions">
            <button class="btn-chip primary" id="btnShowAll">全部顯示</button>
            <button class="btn-chip" id="btnHideAll">全部隱藏</button>
            <button class="btn-chip" id="btnOnlyEast">僅看東區 (綠)</button>
            <button class="btn-chip" id="btnOnlyWest">僅看西區 (紅)</button>
            <button class="btn-chip" id="btnOnlyDiff">僅看非488 (橙)</button>
          </div>
        </div>

        <!-- 搜尋輸入框 -->
        <div class="search-box">
          <div class="search-input-wrap">
            <input type="text" id="searchInput" class="search-input" placeholder="搜尋店名、路名、地址…" />
            <button id="searchClearBtn" class="search-clear">×</button>
          </div>
        </div>

        <!-- 圖層清單 -->
        <div class="layer-tree">
          <!-- 1. 東區主要名單 (綠色) -->
          <div class="layer-group expanded" id="group-east">
            <div class="group-header">
              <input type="checkbox" class="group-checkbox" id="chk-group-east" checked title="全選/取消東區" />
              <div class="group-icon east"></div>
              <div class="group-title-wrap" onclick="toggleGroupExpand('group-east')">
                <span class="group-title">🟢 主要名單-東區 (界線以東)</span>
                <span class="group-stats" id="stats-east">計算中...</span>
              </div>
              <div class="arrow-btn" onclick="toggleGroupExpand('group-east')">▼</div>
            </div>
            <div class="item-list" id="list-east"></div>
          </div>

          <!-- 2. 西區主要名單 (紅色) -->
          <div class="layer-group expanded" id="group-west">
            <div class="group-header">
              <input type="checkbox" class="group-checkbox" id="chk-group-west" checked title="全選/取消西區" />
              <div class="group-icon west"></div>
              <div class="group-title-wrap" onclick="toggleGroupExpand('group-west')">
                <span class="group-title">🔴 主要名單-西區 (界線以西)</span>
                <span class="group-stats" id="stats-west">計算中...</span>
              </div>
              <div class="arrow-btn" onclick="toggleGroupExpand('group-west')">▼</div>
            </div>
            <div class="item-list" id="list-west"></div>
          </div>

          <!-- 3. 非 488 店家名單 (橙色) -->
          <div class="layer-group expanded" id="group-diff">
            <div class="group-header">
              <input type="checkbox" class="group-checkbox" id="chk-group-diff" checked title="全選/取消非488名單" />
              <div class="group-icon diff"></div>
              <div class="group-title-wrap" onclick="toggleGroupExpand('group-diff')">
                <span class="group-title">🟠 非488名單 (亮橙色標記)</span>
                <span class="group-stats" id="stats-diff">計算中...</span>
              </div>
              <div class="arrow-btn" onclick="toggleGroupExpand('group-diff')">▼</div>
            </div>
            <div class="item-list" id="list-diff"></div>
          </div>
        </div>
      </aside>

      <!-- 側邊欄浮動收合切換鈕 -->
      <button class="toggle-sidebar-btn" id="floatingToggleBtn">
        <span id="toggleIcon">◀</span>
        <span id="toggleText">收合清單</span>
      </button>

      <div id="map"></div>
    </div>

    <!-- 同步中提示 -->
    <div id="syncToast" class="sync-toast">正在與 Google 試算表同步...</div>

    <script src="https://unpkg.com/leaflet@1.9.4/dist/leaflet.js"></script>
    <script>
      let stores = [];
      try {
        stores = JSON.parse(<?= storesJson ?>);
      } catch (e) {
        stores = ${JSON.stringify(stores)};
      }

      // 初始化 Leaflet 地圖
      const map = L.map('map', { zoomControl: true }).setView([23.985, 121.605], 13);
      L.tileLayer('https://{s}.tile.openstreetmap.org/{z}/{x}/{y}.png', {
        maxZoom: 19,
        attribution: '&copy; OpenStreetMap & Google Sheets Sync'
      }).addTo(map);

      // 林森路 + 德安一街分界輔助虛線
      const boundaryLine = L.polyline([
        [23.960, 121.581],
        [23.970, 121.5925],
        [23.983, 121.6070],
        [23.995, 121.620]
      ], {
        color: '#6b7280',
        weight: 2,
        dashArray: '6, 6',
        opacity: 0.6
      }).addTo(map);
      boundaryLine.bindTooltip('← 界線以西 (西區紅) ｜ 林森路+德安一街分界 ｜ 界線以東 (東區綠) →', { sticky: true });

      function createPinIcon(colorHex) {
        const svg = \`
          <svg class="gmarker-svg" width="22" height="30" viewBox="0 0 24 34" fill="none" xmlns="http://www.w3.org/2000/svg">
            <path d="M12 0C5.372 0 0 5.373 0 12C0 19.5 12 34 12 34C12 34 24 19.5 24 12C24 5.373 18.628 0 12 0Z" fill="\${colorHex}" stroke="#ffffff" stroke-width="1.5"/>
            <circle cx="12" cy="11" r="4.5" fill="#ffffff"/>
          </svg>
        \`;
        return L.divIcon({
          className: 'gmarker-wrap',
          html: svg,
          iconSize: [22, 30],
          iconAnchor: [11, 30],
          popupAnchor: [0, -28]
        });
      }

      const iconEast = createPinIcon('#2e7d32');
      const iconWest = createPinIcon('#d32f2f');
      const iconDiff = createPinIcon('#e65100');

      function getSmallPinSvg(colorHex) {
        return \`<svg class="pin-svg" viewBox="0 0 24 34" fill="none"><path d="M12 0C5.372 0 0 5.373 0 12C0 19.5 12 34 12 34C12 34 24 19.5 24 12C24 5.373 18.628 0 12 0Z" fill="\${colorHex}" stroke="#ffffff" stroke-width="1.5"/><circle cx="12" cy="11" r="4" fill="#ffffff"/></svg>\`;
      }

      const storeStates = new Map();

      function getGoogleMapsNavUrl(s) {
        return 'https://www.google.com/maps/dir/?api=1&destination=' + s.lat + ',' + s.lng;
      }

      function showToast(text, duration = 2000) {
        const t = document.getElementById('syncToast');
        t.textContent = text;
        t.style.display = 'block';
        setTimeout(() => { t.style.display = 'none'; }, duration);
      }

      // 建立地圖 Marker
      stores.forEach((store, index) => {
        let icon = iconEast;
        let badgeClass = 'east';
        let badgeText = '🟢 主要名單 - 界線以東';
        if (store.subGroup === '主要名單-西區') {
          icon = iconWest;
          badgeClass = 'west';
          badgeText = '🔴 主要名單 - 界線以西';
        } else if (store.subGroup === '非488名單') {
          icon = iconDiff;
          badgeClass = 'diff';
          badgeText = '🟠 非 488 名單店家';
        }

        const marker = L.marker([store.lat, store.lng], { icon: icon });

        const popupDiv = document.createElement('div');
        popupDiv.innerHTML = \`
          <div class="pop-header">
            <span class="pop-badge \${badgeClass}">\${badgeText}</span>
            <span style="font-size:0.75rem;color:#888;">#\${index + 1}</span>
          </div>
          <div class="pop-title">\${store.name}</div>
          <div class="pop-addr">\${store.address}\${store.estimated ? '<br/><span style="color:#b45309;font-size:11px;">※ 位置為地址估算</span>' : ''}</div>
          <div class="pop-actions">
            <a class="pop-nav-btn" href="\${getGoogleMapsNavUrl(store)}" target="_blank" rel="noopener">Google 導航 ↗</a>
            <button class="pop-toggle-btn" onclick="toggleSingleStoreVisible(\${store.id}, false)">隱藏此點</button>
          </div>
        \`;
        marker.bindPopup(popupDiv);
        marker.on('click', () => highlightStoreRow(store.id));

        // 判斷預設是否顯示 (若資料有 isVisible 則遵循，否則預設 true)
        const isInitiallyVisible = (store.isVisible !== undefined) ? store.isVisible : true;
        if (isInitiallyVisible) {
          marker.addTo(map);
        }

        storeStates.set(store.id, {
          store,
          marker,
          isVisible: isInitiallyVisible,
          rowElement: null,
          chkElement: null
        });
      });

      // 渲染左側店家列表清單（帶獨立 Checkbox）
      const listEast = document.getElementById('list-east');
      const listWest = document.getElementById('list-west');
      const listDiff = document.getElementById('list-diff');

      stores.forEach((store, index) => {
        const state = storeStates.get(store.id);
        const row = document.createElement('div');
        row.className = 'store-row' + (state.isVisible ? '' : ' disabled');
        row.id = 'store-row-' + store.id;

        const pinColor = store.subGroup === '主要名單-東區' ? '#2e7d32' : (store.subGroup === '主要名單-西區' ? '#d32f2f' : '#e65100');

        row.innerHTML = \`
          <input type="checkbox" class="item-checkbox" id="chk-store-\${store.id}" \${state.isVisible ? 'checked' : ''} />
          \${getSmallPinSvg(pinColor)}
          <div class="store-info">
            <div class="store-name-text" title="\${store.name}">\${index + 1}. \${store.name}</div>
            <div class="store-addr-text" title="\${store.address}">\${store.address}</div>
          </div>
        \`;

        const chk = row.querySelector('.item-checkbox');
        state.rowElement = row;
        state.chkElement = chk;

        chk.addEventListener('click', (e) => {
          e.stopPropagation();
          setStoreVisibility(store.id, chk.checked, true);
        });

        row.addEventListener('click', () => {
          focusStoreOnMap(store.id);
        });

        if (store.subGroup === '主要名單-東區') {
          listEast.appendChild(row);
        } else if (store.subGroup === '主要名單-西區') {
          listWest.appendChild(row);
        } else {
          listDiff.appendChild(row);
        }
      });

      // 設定單一店家的顯示 / 隱藏狀態（syncToSheet: 是否透過 GAS 寫回 Google Sheet）
      function setStoreVisibility(storeId, isVisible, syncToSheet = false) {
        const state = storeStates.get(storeId);
        if (!state) return;
        state.isVisible = isVisible;
        state.chkElement.checked = isVisible;
        if (isVisible) {
          state.marker.addTo(map);
          state.rowElement.classList.remove('disabled');
        } else {
          map.removeLayer(state.marker);
          state.rowElement.classList.add('disabled');
        }
        updateGroupCheckboxes();
        updateStats();

        // 寫回 Google Sheet
        if (syncToSheet && typeof google !== 'undefined' && google.script && google.script.run) {
          showToast('正在同步至 Google 試算表...', 1200);
          google.script.run.apiUpdateStoreStatus(storeId, isVisible);
        }
      }

      window.toggleSingleStoreVisible = function(storeId, forceState) {
        const state = storeStates.get(storeId);
        if (state) {
          const next = forceState !== undefined ? forceState : !state.isVisible;
          setStoreVisibility(storeId, next, true);
        }
      };

      function focusStoreOnMap(storeId) {
        const state = storeStates.get(storeId);
        if (!state) return;
        if (!state.isVisible) {
          setStoreVisibility(storeId, true, true);
        }
        map.flyTo([state.store.lat, state.store.lng], 16, { duration: 0.5 });
        state.marker.openPopup();
        highlightStoreRow(storeId);
      }

      let currentActiveRow = null;
      function highlightStoreRow(storeId) {
        if (currentActiveRow) {
          currentActiveRow.classList.remove('active');
        }
        const state = storeStates.get(storeId);
        if (state && state.rowElement) {
          state.rowElement.classList.add('active');
          state.rowElement.scrollIntoView({ block: 'nearest', behavior: 'smooth' });
          currentActiveRow = state.rowElement;
        }
      }

      function setGroupVisibility(groupName, isVisible) {
        const updates = [];
        stores.filter(s => s.subGroup === groupName).forEach(s => {
          const state = storeStates.get(s.id);
          state.isVisible = isVisible;
          state.chkElement.checked = isVisible;
          if (isVisible) {
            state.marker.addTo(map);
            state.rowElement.classList.remove('disabled');
          } else {
            map.removeLayer(state.marker);
            state.rowElement.classList.add('disabled');
          }
          updates.push({ id: s.id, isVisible: isVisible });
        });
        updateGroupCheckboxes();
        updateStats();

        if (typeof google !== 'undefined' && google.script && google.script.run) {
          showToast('正在批次儲存至試算表...', 1500);
          google.script.run.apiBatchUpdateStatus(updates);
        }
      }

      const chkGroupEast = document.getElementById('chk-group-east');
      const chkGroupWest = document.getElementById('chk-group-west');
      const chkGroupDiff = document.getElementById('chk-group-diff');

      chkGroupEast.addEventListener('change', () => setGroupVisibility('主要名單-東區', chkGroupEast.checked));
      chkGroupWest.addEventListener('change', () => setGroupVisibility('主要名單-西區', chkGroupWest.checked));
      chkGroupDiff.addEventListener('change', () => setGroupVisibility('非488名單', chkGroupDiff.checked));

      function updateGroupCheckboxes() {
        function checkStatus(groupName, groupChk) {
          const groupStores = stores.filter(s => s.subGroup === groupName);
          const visibleCount = groupStores.filter(s => storeStates.get(s.id).isVisible).length;
          if (visibleCount === groupStores.length) {
            groupChk.checked = true;
            groupChk.indeterminate = false;
          } else if (visibleCount === 0) {
            groupChk.checked = false;
            groupChk.indeterminate = false;
          } else {
            groupChk.checked = false;
            groupChk.indeterminate = true;
          }
          return { visibleCount, total: groupStores.length };
        }

        const eastStats = checkStatus('主要名單-東區', chkGroupEast);
        const westStats = checkStatus('主要名單-西區', chkGroupWest);
        const diffStats = checkStatus('非488名單', chkGroupDiff);

        document.getElementById('stats-east').textContent = \`顯示 \${eastStats.visibleCount} / \${eastStats.total} 筆\`;
        document.getElementById('stats-west').textContent = \`顯示 \${westStats.visibleCount} / \${westStats.total} 筆\`;
        document.getElementById('stats-diff').textContent = \`顯示 \${diffStats.visibleCount} / \${diffStats.total} 筆\`;
      }

      function updateStats() {
        let visibleCount = 0;
        storeStates.forEach(st => {
          if (st.isVisible) visibleCount++;
        });
        document.getElementById('activeCountLabel').textContent = \`目前顯示：\${visibleCount} / \${stores.length} 處店家\`;
      }

      // 快捷篩選
      document.getElementById('btnShowAll').addEventListener('click', () => {
        stores.forEach(s => setStoreVisibility(s.id, true));
      });
      document.getElementById('btnHideAll').addEventListener('click', () => {
        stores.forEach(s => setStoreVisibility(s.id, false));
      });
      document.getElementById('btnOnlyEast').addEventListener('click', () => {
        stores.forEach(s => setStoreVisibility(s.id, s.subGroup === '主要名單-東區'));
      });
      document.getElementById('btnOnlyWest').addEventListener('click', () => {
        stores.forEach(s => setStoreVisibility(s.id, s.subGroup === '主要名單-西區'));
      });
      document.getElementById('btnOnlyDiff').addEventListener('click', () => {
        stores.forEach(s => setStoreVisibility(s.id, s.subGroup === '非488名單'));
      });

      // ----------------------------------------------------
      // 隔天載入全開啟 / 載入昨天 / 載入指定日期按鈕事件
      // ----------------------------------------------------
      document.getElementById('btnLoadAllOpen').addEventListener('click', () => {
        if (!confirm('確定要將今天的所有店家設為【全部開啟顯示】嗎？')) return;
        showToast('正在載入全開啟並同步至試算表...', 2000);
        stores.forEach(s => setStoreVisibility(s.id, true));
        if (typeof google !== 'undefined' && google.script && google.script.run) {
          google.script.run.withSuccessHandler(() => {
            showToast('已成功載入全開啟！', 1800);
          }).loadAllActive();
        }
      });

      document.getElementById('btnLoadYesterday').addEventListener('click', () => {
        if (!confirm('確定要載入【昨天的開關設定】到今天嗎？')) return;
        showToast('正在讀取昨日設定...', 2500);
        if (typeof google !== 'undefined' && google.script && google.script.run) {
          const today = document.getElementById('dateSelector').value;
          const yesterdayDate = new Date(today);
          yesterdayDate.setDate(yesterdayDate.getDate() - 1);
          const yStr = yesterdayDate.toISOString().slice(0, 10);

          google.script.run.withSuccessHandler((res) => {
            alert('已成功載入昨日（' + res.sourceDate + '）的開關設定！目前開啟：' + res.activeCount + ' 家');
            location.reload();
          }).loadStatusFromDate(yStr, today);
        } else {
          alert('請在 Google Apps Script 發佈環境下執行！');
        }
      });

      document.getElementById('btnLoadCustom').addEventListener('click', () => {
        const chosenDate = document.getElementById('dateSelector').value;
        if (!chosenDate) {
          alert('請先選擇日期！');
          return;
        }
        if (confirm('確定要載入指定日期【' + chosenDate + '】的開關設定嗎？')) {
          showToast('正在載入指定日期設定...', 2500);
          if (typeof google !== 'undefined' && google.script && google.script.run) {
            google.script.run.withSuccessHandler((res) => {
              alert('已成功載入 ' + chosenDate + ' 的開關設定！目前開啟：' + res.activeCount + ' 家');
              location.reload();
            }).loadStatusFromDate(chosenDate, chosenDate);
          }
        }
      });

      document.getElementById('btnSyncFromSheet').addEventListener('click', () => {
        showToast('重新載入最新狀態...', 1200);
        location.reload();
      });

      window.toggleGroupExpand = function(groupId) {
        const el = document.getElementById(groupId);
        if (el) el.classList.toggle('expanded');
      };

      const sidebar = document.getElementById('sidebar');
      const floatingToggleBtn = document.getElementById('floatingToggleBtn');
      const menuToggleBtn = document.getElementById('menuToggleBtn');
      const toggleIcon = document.getElementById('toggleIcon');
      const toggleText = document.getElementById('toggleText');

      function toggleSidebar() {
        const isCollapsed = sidebar.classList.toggle('collapsed');
        if (isCollapsed) {
          toggleIcon.textContent = '▶';
          toggleText.textContent = '展開清單';
        } else {
          toggleIcon.textContent = '◀';
          toggleText.textContent = '收合清單';
        }
        setTimeout(() => map.invalidateSize(), 260);
      }
      floatingToggleBtn.addEventListener('click', toggleSidebar);
      menuToggleBtn.addEventListener('click', toggleSidebar);

      const searchInput = document.getElementById('searchInput');
      const searchClearBtn = document.getElementById('searchClearBtn');

      searchInput.addEventListener('input', () => {
        const q = searchInput.value.trim().toLowerCase();
        searchClearBtn.style.display = q ? 'block' : 'none';

        stores.forEach(store => {
          const state = storeStates.get(store.id);
          const matches = !q || store.name.toLowerCase().includes(q) || store.address.toLowerCase().includes(q);
          state.rowElement.style.display = matches ? 'flex' : 'none';
        });
      });

      searchClearBtn.addEventListener('click', () => {
        searchInput.value = '';
        searchClearBtn.style.display = 'none';
        stores.forEach(store => {
          storeStates.get(store.id).rowElement.style.display = 'flex';
        });
      });

      updateGroupCheckboxes();
      updateStats();
    </script>
  </body>
</html>
`;

fs.writeFileSync(path.join(gasDir, "Index.html"), indexHtml, "utf8");
console.log("gas/Index.html 已成功生成！");
