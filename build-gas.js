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
 * 升級功能：
 * 1. 【有動作就自動更新】：只要在網頁或試算表進行任何開關、打卡或篩選，即自動同步試算表並記錄歷史。
 * 2. 【記錄最後 10 次更新】：歷史表完整保存每一次變更，並提供 API 即時提取最後 10 次歷史版本。
 * 3. 【一鍵回到昨天最後紀錄】：精確鎖定昨天最後一次存檔版本，快速回溯。
 * 4. 【版本自由回溯】：支援在網頁上一鍵回復至最近 10 次內的任何一個歷史版本。
 */

function doGet(e) {
  var template = HtmlService.createTemplateFromFile('Index');
  
  var targetDate = (e && e.parameter && e.parameter.date) ? e.parameter.date : getTodayString();
  template.currentDate = targetDate;
  
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  template.sheetUrl = ss ? ss.getUrl() : "";
  
  var storeData = getStoresForDate(targetDate);
  template.storesJson = JSON.stringify(storeData);
  
  return template.evaluate()
    .setTitle('jay的地圖 - 花蓮店家每日排程地圖 (' + targetDate + ')')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1.0')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('🗺️ 店家地圖與每日排程')
    .addItem('🟢 1. 建立/初始化 500 筆店家工作表', 'setupSheetWithData')
    .addSeparator()
    .addItem('☀️ 2. 【隔天操作】載入全開啟 (全部勾選顯示)', 'menuLoadAllActive')
    .addItem('📋 3. 【一鍵回溯】回到昨天的最後紀錄', 'menuLoadYesterdayLast')
    .addItem('📜 4. 【檢視歷史】檢視最近 10 次更新紀錄', 'menuShowRecentHistory')
    .addItem('📅 5. 【指定載入】載入指定日期的開關狀態', 'menuLoadCustomDate')
    .addSeparator()
    .addItem('🌐 6. 開啟地圖 Web App 網頁', 'openWebAppUrl')
    .addToUi();
}

// ------------------------------------------
// 日期工具函式 (台北時區 GMT+8，完全避免時區偏差)
// ------------------------------------------
function getTodayString() {
  return Utilities.formatDate(new Date(), "Asia/Taipei", "yyyy-MM-dd");
}

function getYesterdayString(dateStr) {
  var baseDateStr = dateStr || getTodayString();
  var parts = baseDateStr.split('-');
  var y = parseInt(parts[0], 10);
  var m = parseInt(parts[1], 10) - 1;
  var d = parseInt(parts[2], 10);
  var dt = new Date(y, m, d - 1);
  return Utilities.formatDate(dt, "Asia/Taipei", "yyyy-MM-dd");
}

function normalizeDateString(val) {
  if (!val) return "";
  if (val instanceof Date) {
    return Utilities.formatDate(val, "Asia/Taipei", "yyyy-MM-dd");
  }
  var str = String(val).trim();
  var m = str.match(/(\\d{4})[-/](\\d{1,2})[-/](\\d{1,2})/);
  if (m) {
    var y = m[1];
    var mm = m[2].length === 1 ? ("0" + m[2]) : m[2];
    var dd = m[3].length === 1 ? ("0" + m[3]) : m[3];
    return y + "-" + mm + "-" + dd;
  }
  try {
    var d = new Date(str);
    if (!isNaN(d.getTime())) {
      return Utilities.formatDate(d, "Asia/Taipei", "yyyy-MM-dd");
    }
  } catch(e) {}
  return str;
}

function parseStoreNum(rawId) {
  if (typeof rawId === 'number') return rawId;
  var digits = String(rawId).replace(/\\D/g, '');
  return digits ? parseInt(digits, 10) : 0;
}

// ------------------------------------------
// 取得或建立歷史紀錄存檔工作表
// ------------------------------------------
function getOrCreateHistorySheet(ss) {
  var histSheet = ss.getSheetByName("歷史紀錄存檔");
  if (!histSheet) {
    histSheet = ss.insertSheet("歷史紀錄存檔");
    var headers = [["版本時間", "所屬日期", "操作動作說明", "開啟總數", "關閉總數", "關閉店家序號清單", "已打卡店家序號清單"]];
    histSheet.getRange(1, 1, 1, headers[0].length).setValues(headers);
    histSheet.getRange(1, 1, 1, headers[0].length).setFontWeight("bold").setBackground("#e8f0fe");
    histSheet.setFrozenRows(1);
    histSheet.getRange("A:B").setNumberFormat("@");
    histSheet.getRange("F:G").setNumberFormat("@");
    histSheet.setColumnWidth(1, 160);
    histSheet.setColumnWidth(2, 100);
    histSheet.setColumnWidth(3, 200);
    histSheet.setColumnWidth(4, 80);
    histSheet.setColumnWidth(5, 80);
    histSheet.setColumnWidth(6, 250);
    histSheet.setColumnWidth(7, 180);
  }
  return histSheet;
}

// ------------------------------------------
// 核心：查詢指定日期的店家資料
// ------------------------------------------
function getStoresForDate(dateStr) {
  var defaultList = getDefaultStores();
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) return defaultList;

  var sheet = ss.getSheetByName("當日店家清單");
  if (!sheet) return defaultList;

  var targetDate = normalizeDateString(dateStr || getTodayString());
  var sheetDate = normalizeDateString(sheet.getRange("B1").getValue());
  var data = sheet.getDataRange().getValues();
  if (data.length <= 3) return defaultList;

  var hiddenStoreIds = new Set();
  var checkedInStoreIds = new Set();

  if (sheetDate === targetDate) {
    for (var i = 3; i < data.length; i++) {
      var row = data[i];
      var isVisible = (row[1] === true || String(row[1]).toUpperCase() === "TRUE");
      var storeNum = parseStoreNum(row[0]);
      if (!isVisible) {
        hiddenStoreIds.add(storeNum);
      }
      var note = String(row[9] || "");
      if (note.indexOf("已打卡") !== -1) {
        checkedInStoreIds.add(storeNum);
      }
    }
  } else {
    // 跨日期查詢：從歷史紀錄中最新的一筆載入
    var histRes = getLatestHistoryForDate(targetDate);
    hiddenStoreIds = histRes.hiddenSet;
    checkedInStoreIds = histRes.checkedInSet;
  }

  return defaultList.map(function(s, idx) {
    var num = idx + 1;
    return Object.assign({}, s, {
      isVisible: !hiddenStoreIds.has(num),
      isCheckedIn: checkedInStoreIds.has(num)
    });
  });
}

// 倒序尋找指定日期的最後一次存檔
function getLatestHistoryForDate(dateStr) {
  var hiddenSet = new Set();
  var checkedInSet = new Set();
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) return { hiddenSet: hiddenSet, checkedInSet: checkedInSet, found: false };

  var histSheet = ss.getSheetByName("歷史紀錄存檔");
  if (!histSheet) return { hiddenSet: hiddenSet, checkedInSet: checkedInSet, found: false };

  var targetDate = normalizeDateString(dateStr);
  var data = histSheet.getDataRange().getValues();
  var isNewFormat = (data.length > 0 && data[0].length >= 7 && data[0][2] === "操作動作說明");

  for (var i = data.length - 1; i >= 1; i--) {
    var row = data[i];
    var rowDate = isNewFormat ? normalizeDateString(row[1]) : normalizeDateString(row[0]);
    if (rowDate === targetDate) {
      var hiddenListStr = isNewFormat ? String(row[5] || "") : String(row[1] || "");
      var checkedListStr = isNewFormat ? String(row[6] || "") : "";

      if (hiddenListStr) {
        hiddenListStr.split(",").forEach(function(idStr) {
          var n = parseInt(idStr.trim(), 10);
          if (!isNaN(n)) hiddenSet.add(n);
        });
      }
      if (checkedListStr) {
        checkedListStr.split(",").forEach(function(idStr) {
          var n = parseInt(idStr.trim(), 10);
          if (!isNaN(n)) checkedInSet.add(n);
        });
      }
      return {
        hiddenSet: hiddenSet,
        checkedInSet: checkedInSet,
        found: true,
        rowIndex: i + 1,
        timestamp: String(row[0]),
        action: isNewFormat ? String(row[2]) : "歷史存檔"
      };
    }
  }
  return { hiddenSet: hiddenSet, checkedInSet: checkedInSet, found: false };
}

// ------------------------------------------
// 寫入一筆新的更新紀錄至歷史表
// ------------------------------------------
function appendHistoryRecord(dateStr, actionNote, hiddenIds, checkedInIds, activeCount) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) return null;

  var histSheet = getOrCreateHistorySheet(ss);
  var nowStr = Utilities.formatDate(new Date(), "Asia/Taipei", "yyyy-MM-dd HH:mm:ss");
  var dStr = normalizeDateString(dateStr || getTodayString());
  var closedCount = hiddenIds ? hiddenIds.length : 0;
  var openCount = (typeof activeCount === 'number') ? activeCount : (500 - closedCount);
  var hiddenStr = (hiddenIds && hiddenIds.length) ? hiddenIds.join(",") : "";
  var checkedStr = (checkedInIds && checkedInIds.length) ? checkedInIds.join(",") : "";

  var rowValues = [nowStr, dStr, actionNote || "狀態更新", openCount, closedCount, hiddenStr, checkedStr];
  histSheet.appendRow(rowValues);

  return {
    timestamp: nowStr,
    date: dStr,
    action: actionNote || "狀態更新",
    activeCount: openCount,
    closedCount: closedCount
  };
}

// ------------------------------------------
// 取得最近 10 次更新紀錄 (供前端下拉選單即時呈現)
// ------------------------------------------
function apiGetRecentHistory(limit) {
  try {
    var maxCount = limit || 10;
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    if (!ss) return { success: false, msg: "未綁定試算表", list: [] };

    var histSheet = ss.getSheetByName("歷史紀錄存檔");
    if (!histSheet) return { success: true, list: [] };

    var data = histSheet.getDataRange().getValues();
    if (data.length <= 1) return { success: true, list: [] };

    var isNewFormat = (data[0].length >= 7 && data[0][2] === "操作動作說明");
    var todayStr = getTodayString();
    var yesterdayStr = getYesterdayString(todayStr);

    var list = [];
    var yesterdayLastMarked = false;

    // 從最後一行往上收集
    for (var i = data.length - 1; i >= 1; i--) {
      var row = data[i];
      var timeVal = String(row[0] || "");
      var dateVal = isNewFormat ? normalizeDateString(row[1]) : normalizeDateString(row[0]);
      var actionVal = isNewFormat ? String(row[2] || "狀態更新") : "手動存檔";
      var openVal = isNewFormat ? parseInt(row[3], 10) : parseInt(row[2], 10);
      var closeVal = isNewFormat ? parseInt(row[4], 10) : parseInt(row[3], 10);
      var hiddenStr = isNewFormat ? String(row[5] || "") : String(row[1] || "");
      var checkedStr = isNewFormat ? String(row[6] || "") : "";

      var isYesterday = (dateVal === yesterdayStr);
      var isYesterdayLast = false;
      if (isYesterday && !yesterdayLastMarked) {
        isYesterdayLast = true;
        yesterdayLastMarked = true;
      }

      list.push({
        rowIndex: i + 1,
        timestamp: timeVal,
        date: dateVal,
        action: actionVal,
        activeCount: isNaN(openVal) ? 0 : openVal,
        closedCount: isNaN(closeVal) ? 0 : closeVal,
        hiddenIds: hiddenStr,
        checkedInIds: checkedStr,
        isLatest: (list.length === 0),
        isYesterdayLast: isYesterdayLast
      });

      if (list.length >= maxCount) break;
    }

    return {
      success: true,
      today: todayStr,
      yesterday: yesterdayStr,
      list: list
    };
  } catch (err) {
    return { success: false, error: String(err), list: [] };
  }
}

// ------------------------------------------
// 核心 API：【有動作就更新】自動存檔完整狀態
// ------------------------------------------
function apiAutoSaveFullState(dateStr, hiddenIds, checkedInIds, actionNote) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    if (!ss) return { success: false, msg: "未綁定試算表" };

    var sheet = ss.getSheetByName("當日店家清單");
    if (!sheet) return { success: false, msg: "找不到【當日店家清單】工作表" };

    var targetDate = normalizeDateString(dateStr || getTodayString());
    sheet.getRange("B1").setValue(targetDate);

    var hiddenSet = new Set((hiddenIds || []).map(function(id) { return parseStoreNum(id); }));
    var checkedSet = new Set((checkedInIds || []).map(function(id) { return parseStoreNum(id); }));

    var lastRow = sheet.getLastRow();
    var numRows = 0;
    var activeCount = 0;

    if (lastRow >= 4) {
      numRows = lastRow - 3;
      var currentVisibility = sheet.getRange(4, 2, numRows, 1).getValues();
      var currentNotes = sheet.getRange(4, 10, numRows, 1).getValues();

      var visChanges = false;
      var noteChanges = false;
      var newVisibility = [];
      var newNotes = [];

      for (var i = 1; i <= numRows; i++) {
        var isVis = !hiddenSet.has(i);
        if (isVis) activeCount++;
        newVisibility.push([isVis]);
        if (currentVisibility[i - 1][0] !== isVis) {
          visChanges = true;
        }

        var isChk = checkedSet.has(i);
        var curNote = String(currentNotes[i - 1][0] || "");
        var newNote = curNote;
        if (isChk && curNote.indexOf("已打卡") === -1) {
          var nowTime = Utilities.formatDate(new Date(), "Asia/Taipei", "HH:mm");
          newNote = "已打卡 " + nowTime;
          noteChanges = true;
        } else if (!isChk && curNote.indexOf("已打卡") !== -1) {
          newNote = curNote.replace(/已打卡\\s*\\d{2}:\\d{2}/, "").trim();
          noteChanges = true;
        }
        newNotes.push([newNote]);
      }

      if (visChanges) {
        sheet.getRange(4, 2, numRows, 1).setValues(newVisibility);
      }
      if (noteChanges) {
        sheet.getRange(4, 10, numRows, 1).setValues(newNotes);
      }

      sheet.getRange("D1").setValue("已顯示: " + activeCount + " / " + numRows);
    }

    // 寫入歷史表 (記錄本次更新)
    var note = actionNote || ("自動存檔 (開啟 " + activeCount + " / 關閉 " + (hiddenSet.size) + ")");
    var histRec = appendHistoryRecord(targetDate, note, Array.from(hiddenSet), Array.from(checkedSet), activeCount);

    return {
      success: true,
      timestamp: histRec ? histRec.timestamp : Utilities.formatDate(new Date(), "Asia/Taipei", "HH:mm:ss"),
      activeCount: activeCount,
      closedCount: hiddenSet.size
    };
  } catch (err) {
    return { success: false, error: String(err) };
  }
}

// ------------------------------------------
// 核心 API：一鍵回到昨天的最後紀錄
// ------------------------------------------
function apiLoadYesterdayLastRecord(currentDateStr) {
  try {
    var today = normalizeDateString(currentDateStr || getTodayString());
    var yesterday = getYesterdayString(today);

    var hist = getLatestHistoryForDate(yesterday);
    if (!hist.found) {
      return {
        success: false,
        msg: "在【歷史紀錄存檔】中找不到昨日（" + yesterday + "）的存檔紀錄！",
        yesterday: yesterday
      };
    }

    // 將昨日狀態套用至今日
    var hiddenArray = Array.from(hist.hiddenSet);
    var checkedArray = Array.from(hist.checkedInSet);
    var actionNote = "回復至昨日最後紀錄 (" + yesterday + " " + hist.timestamp + ")";

    var saveRes = apiAutoSaveFullState(today, hiddenArray, checkedArray, actionNote);
    var latestStores = getStoresForDate(today);

    return {
      success: true,
      yesterdayDate: yesterday,
      yesterdayTimestamp: hist.timestamp,
      activeCount: saveRes.activeCount,
      closedCount: saveRes.closedCount,
      stores: latestStores
    };
  } catch (err) {
    return { success: false, error: String(err) };
  }
}

// ------------------------------------------
// 核心 API：回復至指定歷史版本 (由最近 10 次清單觸發)
// ------------------------------------------
function apiRevertToHistoryVersion(rowIndexOrTimestamp, todayStr) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    if (!ss) return { success: false, msg: "未綁定試算表" };

    var histSheet = ss.getSheetByName("歷史紀錄存檔");
    if (!histSheet) return { success: false, msg: "找不到歷史紀錄表" };

    var data = histSheet.getDataRange().getValues();
    if (data.length <= 1) return { success: false, msg: "歷史紀錄為空" };

    var isNewFormat = (data[0].length >= 7 && data[0][2] === "操作動作說明");
    var targetRow = null;

    if (typeof rowIndexOrTimestamp === 'number' && rowIndexOrTimestamp >= 2 && rowIndexOrTimestamp <= data.length) {
      targetRow = data[rowIndexOrTimestamp - 1];
    } else {
      var queryStr = String(rowIndexOrTimestamp).trim();
      for (var i = data.length - 1; i >= 1; i--) {
        if (String(data[i][0]).indexOf(queryStr) !== -1) {
          targetRow = data[i];
          break;
        }
      }
    }

    if (!targetRow) {
      return { success: false, msg: "找不到指定的歷史版本" };
    }

    var hiddenStr = isNewFormat ? String(targetRow[5] || "") : String(targetRow[1] || "");
    var checkedStr = isNewFormat ? String(targetRow[6] || "") : "";
    var targetTime = String(targetRow[0]);

    var hiddenIds = [];
    if (hiddenStr) {
      hiddenStr.split(",").forEach(function(s) {
        var n = parseInt(s.trim(), 10);
        if (!isNaN(n)) hiddenIds.push(n);
      });
    }

    var checkedIds = [];
    if (checkedStr) {
      checkedStr.split(",").forEach(function(s) {
        var n = parseInt(s.trim(), 10);
        if (!isNaN(n)) checkedIds.push(n);
      });
    }

    var today = normalizeDateString(todayStr || getTodayString());
    var actionNote = "回復至歷史版本 [" + targetTime + "]";
    var saveRes = apiAutoSaveFullState(today, hiddenIds, checkedIds, actionNote);
    var latestStores = getStoresForDate(today);

    return {
      success: true,
      revertedTime: targetTime,
      activeCount: saveRes.activeCount,
      closedCount: saveRes.closedCount,
      stores: latestStores
    };
  } catch (err) {
    return { success: false, error: String(err) };
  }
}

// ------------------------------------------
// 載入全開啟 API
// ------------------------------------------
function loadAllActive(targetDate) {
  var today = normalizeDateString(targetDate || getTodayString());
  var saveRes = apiAutoSaveFullState(today, [], [], "☀️ 載入全開啟 (500家全部開啟)");
  var latestStores = getStoresForDate(today);
  return {
    success: true,
    date: today,
    activeCount: 500,
    closedCount: 0,
    stores: latestStores
  };
}

// 載入指定日期的狀態 API
function loadStatusFromDate(sourceDateStr, targetDateStr) {
  var srcDate = normalizeDateString(sourceDateStr);
  var targetDate = normalizeDateString(targetDateStr || getTodayString());

  var hist = getLatestHistoryForDate(srcDate);
  var hiddenArray = Array.from(hist.hiddenSet);
  var checkedArray = Array.from(hist.checkedInSet);
  var actionNote = "載入歷史日期 [" + srcDate + "] 之開關設定";

  var saveRes = apiAutoSaveFullState(targetDate, hiddenArray, checkedArray, actionNote);
  var latestStores = getStoresForDate(targetDate);

  return {
    success: true,
    targetDate: targetDate,
    sourceDate: srcDate,
    activeCount: saveRes.activeCount,
    closedCount: saveRes.closedCount,
    stores: latestStores
  };
}

function apiLoadYesterdayStatus(currentDateStr) {
  return apiLoadYesterdayLastRecord(currentDateStr);
}

function apiGetStoresForDate(dateStr) {
  try {
    var targetDate = normalizeDateString(dateStr || getTodayString());
    var list = getStoresForDate(targetDate);
    return { success: true, stores: list };
  } catch (err) {
    return { success: false, error: String(err) };
  }
}

// 單筆快速 API (相容舊呼叫)
function apiToggleCheckin(rawStoreId, isCheckedIn) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    if (!ss) return { success: false };
    var sheet = ss.getSheetByName("當日店家清單");
    if (!sheet) return { success: false };
    var num = parseStoreNum(rawStoreId);
    if (!num) return { success: false };
    var rowNumber = num + 3;
    var nowTime = Utilities.formatDate(new Date(), "Asia/Taipei", "HH:mm");
    var noteVal = isCheckedIn ? ("已打卡 " + nowTime) : "";
    sheet.getRange(rowNumber, 10).setValue(noteVal);
    return { success: true };
  } catch (e) {
    return { success: false };
  }
}

function apiUpdateStoreStatus(rawStoreId, isVisible) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    if (!ss) return { success: false };
    var sheet = ss.getSheetByName("當日店家清單");
    if (!sheet) return { success: false };
    var num = parseStoreNum(rawStoreId);
    if (!num) return { success: false };
    var rowNumber = num + 3;
    sheet.getRange(rowNumber, 2).setValue(isVisible);
    return { success: true };
  } catch (e) {
    return { success: false };
  }
}

function apiBatchUpdateStatus(updates) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    if (!ss) return { success: false };
    var sheet = ss.getSheetByName("當日店家清單");
    if (!sheet) return { success: false };
    updates.forEach(function(u) {
      var num = parseStoreNum(u.id);
      if (num) {
        sheet.getRange(num + 3, 2).setValue(u.isVisible);
      }
    });
    return { success: true };
  } catch (e) {
    return { success: false };
  }
}

// ------------------------------------------
// 試算表選單事件
// ------------------------------------------
function menuLoadAllActive() {
  var ui = SpreadsheetApp.getUi();
  var res = ui.alert("載入全開啟", "確定要將今天的店家狀態全部設為【開啟顯示】嗎？", ui.ButtonSet.YES_NO);
  if (res === ui.Button.YES) {
    var out = loadAllActive();
    ui.alert("成功！", "已成功將今天（" + out.date + "）所有 500 筆店家設為開啟顯示！", ui.ButtonSet.OK);
  }
}

function menuLoadYesterdayLast() {
  var ui = SpreadsheetApp.getUi();
  var today = getTodayString();
  var yesterday = getYesterdayString(today);

  var res = ui.alert("一鍵回到昨天最後紀錄", "確定要載入昨天（" + yesterday + "）的最後一次存檔紀錄到今天嗎？", ui.ButtonSet.YES_NO);
  if (res === ui.Button.YES) {
    var out = apiLoadYesterdayLastRecord(today);
    if (out.success) {
      ui.alert("載入完成！", "已成功載入昨日（" + out.yesterdayDate + "）的最後存檔！\\n- 存檔時間：" + out.yesterdayTimestamp + "\\n- 開啟店家數：" + out.activeCount + " 家\\n- 關閉店家數：" + out.closedCount + " 家", ui.ButtonSet.OK);
    } else {
      ui.alert("提示", out.msg || "找不到昨日紀錄！", ui.ButtonSet.OK);
    }
  }
}

function menuShowRecentHistory() {
  var res = apiGetRecentHistory(10);
  var ui = SpreadsheetApp.getUi();
  if (!res.success || !res.list || res.list.length === 0) {
    ui.alert("歷史紀錄", "目前尚無歷史更新紀錄！", ui.ButtonSet.OK);
    return;
  }
  var msg = "【最近更新紀錄 (最後 " + res.list.length + " 次)】\\n\\n";
  res.list.forEach(function(item, idx) {
    var tag = item.isLatest ? " [最新]" : (item.isYesterdayLast ? " ⭐[昨天最後]" : "");
    msg += (idx + 1) + ". " + item.timestamp + tag + "\\n";
    msg += "   - 動作：" + item.action + "\\n";
    msg += "   - 狀態：開啟 " + item.activeCount + " 家 / 關閉 " + item.closedCount + " 家\\n\\n";
  });
  msg += "提示：您也可以在網頁 Web App 上方點擊【📜 最近 10 次紀錄】一鍵復原！";
  ui.alert("最近更新紀錄", msg, ui.ButtonSet.OK);
}

function menuLoadCustomDate() {
  var ui = SpreadsheetApp.getUi();
  var prompt = ui.prompt("載入指定日期", "請輸入要載入的歷史日期 (格式：YYYY-MM-DD，例如 2026-10-21)：", ui.ButtonSet.OK_CANCEL);
  if (prompt.getSelectedButton() === ui.Button.OK) {
    var srcDate = normalizeDateString(prompt.getResponseText().trim());
    if (!/^\\d{4}-\\d{2}-\\d{2}$/.test(srcDate)) {
      ui.alert("錯誤", "日期格式不正確，請使用 YYYY-MM-DD 格式！", ui.ButtonSet.OK);
      return;
    }
    var today = getTodayString();
    var out = loadStatusFromDate(srcDate, today);
    ui.alert("載入完成！", "已成功載入 " + srcDate + " 的設定！\\n- 開啟店家數：" + out.activeCount + " 家\\n- 關閉店家數：" + out.closedCount + " 家", ui.ButtonSet.OK);
  }
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

  sheet.getRange("A1").setValue("📅 排程日期：").setFontWeight("bold");
  sheet.getRange("B1").setValue(today).setFontWeight("bold").setFontColor("#1a73e8").setFontSize(12);
  sheet.getRange("C1").setValue("📊 狀態統計：").setFontWeight("bold");
  sheet.getRange("D1").setValue("已顯示: " + stores.length + " / " + stores.length).setFontWeight("bold");
  sheet.getRange("E1").setValue("💡 自動更新機制：網頁端有任何勾選、篩選或打卡動作皆會自動存檔並記錄最近 10 次歷程！").setFontColor("#5f6368");

  var headers = [["序號", "顯示開關", "店名", "地址", "分組分類", "區域", "標記顏色", "緯度", "經度", "備註/打卡"]];
  sheet.getRange(3, 1, 1, headers[0].length).setValues(headers);
  sheet.getRange(3, 1, 1, headers[0].length).setFontWeight("bold").setBackground("#c5221f").setFontColor("#ffffff");

  var rows = stores.map(function(s, idx) {
    return [
      idx + 1,
      true,
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

  var chkRange = sheet.getRange(4, 2, rows.length, 1);
  chkRange.insertCheckboxes();

  sheet.setFrozenRows(3);
  sheet.setFrozenColumns(3);

  sheet.setColumnWidth(1, 50);
  sheet.setColumnWidth(2, 80);
  sheet.setColumnWidth(3, 200);
  sheet.setColumnWidth(4, 260);
  sheet.setColumnWidth(5, 140);
  sheet.setColumnWidth(10, 130);

  // 初始化歷史表第一筆
  getOrCreateHistorySheet(ss);
  appendHistoryRecord(today, "初始化店家清單 (全開 500 家)", [], [], 500);

  var ui = SpreadsheetApp.getUi();
  ui.alert("初始化完成！", "已成功建立【當日店家清單】與【歷史紀錄存檔】！\\n\\n1. B 欄支援打勾即時切換顯示\\n2. 網頁端任何動作皆會自動存檔至試算表\\n3. 自動保留最後 10 次更新紀錄，方便隨時回到昨天或前幾次更新！", ui.ButtonSet.OK);
}

function getDefaultStores() {
  return STORE_DATA_PLACEHOLDER;
}
`;

const fullCodeGs = codeGsContent.replace(
  "STORE_DATA_PLACEHOLDER",
  JSON.stringify(stores, null, 2)
);
fs.writeFileSync(path.join(gasDir, "Code.gs"), fullCodeGs, "utf8");
console.log("gas/Code.gs 已成功生成 (含自動存檔與最後10次更新紀錄)！");

// ==========================================
// 2. 生成 gas/Index.html
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
        --color-diff: #f59e0b;       /* 黃色/金黃色 */
        --border-color: #dadce0;
        --text-primary: #202124;
        --text-secondary: #5f6368;
        --sidebar-width: 400px;
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

      header {
        background: var(--google-red);
        color: white;
        height: 56px;
        padding: 0 16px;
        display: flex;
        align-items: center;
        justify-content: space-between;
        box-shadow: 0 1px 3px rgba(0,0,0,0.2);
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
        cursor: pointer;
        padding: 6px;
        border-radius: 50%;
        display: flex;
        align-items: center;
        justify-content: center;
      }
      .menu-btn:hover { background: rgba(255,255,255,0.15); }

      .brand-box {
        display: flex;
        flex-direction: column;
      }
      .brand-title {
        font-size: 1.15rem;
        font-weight: 700;
        display: flex;
        align-items: center;
        gap: 8px;
      }
      .date-badge {
        font-size: 0.75rem;
        background: rgba(255,255,255,0.25);
        padding: 2px 8px;
        border-radius: 12px;
        font-weight: 600;
      }

      .header-right {
        display: flex;
        align-items: center;
        gap: 8px;
      }
      .header-btn {
        background: white;
        color: var(--google-red);
        border: none;
        padding: 6px 12px;
        border-radius: 6px;
        font-size: 0.82rem;
        font-weight: 700;
        cursor: pointer;
        display: inline-flex;
        align-items: center;
        gap: 6px;
        text-decoration: none;
        transition: all 150ms ease;
      }
      .header-btn:hover {
        background: #fdf2f2;
        transform: translateY(-1px);
      }
      .header-btn-sheet {
        background: #188038;
        color: white;
      }
      .header-btn-sheet:hover {
        background: #137333;
      }

      /* 每日排程與歷史控制列 */
      .schedule-bar {
        background: #ffffff;
        border-bottom: 1px solid var(--border-color);
        padding: 8px 16px;
        display: flex;
        align-items: center;
        justify-content: space-between;
        flex-wrap: wrap;
        gap: 8px;
        font-size: 0.84rem;
        z-index: 999;
        flex-shrink: 0;
      }
      .schedule-left {
        display: flex;
        align-items: center;
        gap: 8px;
        flex-wrap: wrap;
      }
      .schedule-right {
        display: flex;
        align-items: center;
        gap: 10px;
      }
      .date-input {
        border: 1px solid var(--border-color);
        border-radius: 4px;
        padding: 4px 8px;
        font-family: inherit;
        font-size: 0.82rem;
        color: var(--text-primary);
        outline: none;
      }
      .date-input:focus { border-color: #1a73e8; }

      .sched-btn {
        background: #f1f3f4;
        border: 1px solid var(--border-color);
        border-radius: 6px;
        padding: 5px 10px;
        font-size: 0.8rem;
        font-weight: 600;
        color: var(--text-primary);
        cursor: pointer;
        display: inline-flex;
        align-items: center;
        gap: 4px;
        transition: background 120ms ease;
      }
      .sched-btn:hover { background: #e8eaed; }
      .sched-btn.action-primary {
        background: #e8f0fe;
        color: #1a73e8;
        border-color: #c2e7ff;
      }
      .sched-btn.action-primary:hover { background: #d2e3fc; }
      .sched-btn.action-history {
        background: #fef3c7;
        color: #92400e;
        border-color: #fde68a;
      }
      .sched-btn.action-history:hover { background: #fde68a; }

      /* 即時自動儲存狀態標籤 */
      .autosave-pill {
        display: inline-flex;
        align-items: center;
        gap: 5px;
        padding: 4px 10px;
        border-radius: 12px;
        font-size: 0.75rem;
        font-weight: 700;
        background: #f1f3f4;
        color: var(--text-secondary);
        transition: all 200ms ease;
      }
      .autosave-pill.saved {
        background: #e6f4ea;
        color: #137333;
      }
      .autosave-pill.saving {
        background: #fef7e0;
        color: #b06000;
      }
      .autosave-dot {
        width: 7px;
        height: 7px;
        border-radius: 50%;
        background: currentColor;
      }

      /* 側邊欄與地圖主容器 */
      .layout-body {
        display: flex;
        flex: 1;
        overflow: hidden;
        position: relative;
      }

      .sidebar {
        width: var(--sidebar-width);
        background: #ffffff;
        border-right: 1px solid var(--border-color);
        display: flex;
        flex-direction: column;
        height: 100%;
        flex-shrink: 0;
        transition: width 240ms cubic-bezier(0.4, 0, 0.2, 1);
        z-index: 500;
        overflow: hidden;
      }
      .sidebar.collapsed {
        width: 0;
        border-right: none;
      }

      .toggle-sidebar-btn {
        position: absolute;
        top: 12px;
        left: 412px;
        z-index: 600;
        background: white;
        border: 1px solid var(--border-color);
        border-radius: 6px;
        padding: 6px 10px;
        font-size: 0.78rem;
        font-weight: 600;
        color: var(--text-secondary);
        cursor: pointer;
        box-shadow: 0 1px 4px rgba(0,0,0,0.15);
        display: flex;
        align-items: center;
        gap: 4px;
        transition: left 240ms cubic-bezier(0.4, 0, 0.2, 1);
      }
      .sidebar.collapsed + .toggle-sidebar-btn {
        left: 12px;
      }

      #map {
        flex: 1;
        height: 100%;
        background: #eaeaea;
      }

      /* 側邊欄控制面板 */
      .panel-intro {
        padding: 10px 16px;
        border-bottom: 1px solid var(--border-color);
        background: #f8f9fa;
      }
      .intro-meta {
        display: flex;
        justify-content: space-between;
        align-items: center;
        font-size: 0.78rem;
        color: var(--text-secondary);
        font-weight: 600;
        margin-bottom: 8px;
      }
      .quick-actions {
        display: flex;
        flex-wrap: wrap;
        gap: 5px;
      }
      .btn-chip {
        padding: 4px 8px;
        border: 1px solid var(--border-color);
        background: white;
        border-radius: 12px;
        font-size: 0.72rem;
        font-weight: 600;
        color: var(--text-secondary);
        cursor: pointer;
        transition: all 120ms;
        display: inline-flex;
        align-items: center;
        gap: 4px;
      }
      .btn-chip:hover {
        background: var(--google-hover);
        color: var(--text-primary);
        border-color: #bdc1c6;
      }
      .btn-chip.highlight-redgreen {
        background: #e8f5e9;
        color: #1b5e20;
        border-color: #a5d6a7;
        font-weight: 700;
      }
      .btn-chip.highlight-redgreen:hover {
        background: #c8e6c9;
      }

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

      .layer-tree {
        flex: 1;
        overflow-y: auto;
      }

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

      .store-row {
        display: flex;
        align-items: center;
        padding: 8px 12px 8px 16px;
        border-top: 1px solid #f1f3f4;
        cursor: pointer;
        transition: background 100ms;
        gap: 8px;
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

      .checkin-btn {
        padding: 4px 8px;
        font-size: 0.72rem;
        font-weight: 600;
        border-radius: 4px;
        border: 1px solid #d1d5db;
        background: #ffffff;
        color: #4b5563;
        cursor: pointer;
        flex-shrink: 0;
        transition: all 140ms ease;
        display: inline-flex;
        align-items: center;
        gap: 3px;
      }
      .checkin-btn:hover {
        background: #f3f4f6;
        border-color: #9ca3af;
      }
      .checkin-btn.checked-in {
        background: #ecfdf5;
        border-color: #10b981;
        color: #059669;
        font-weight: 700;
      }

      /* 彈窗樣式 */
      .pop-header {
        display: flex;
        justify-content: space-between;
        align-items: center;
        margin-bottom: 6px;
      }
      .pop-badge {
        font-size: 11px;
        font-weight: 700;
        padding: 2px 6px;
        border-radius: 4px;
        color: white;
      }
      .pop-badge.east { background: var(--color-east); }
      .pop-badge.west { background: var(--color-west); }
      .pop-badge.diff { background: var(--color-diff); }

      .pop-title {
        font-size: 14px;
        font-weight: 700;
        color: #111827;
        margin-bottom: 4px;
      }
      .pop-addr {
        font-size: 12px;
        color: #6b7280;
        margin-bottom: 8px;
        line-height: 1.4;
      }
      .pop-actions {
        display: flex;
        gap: 6px;
        margin-top: 6px;
      }
      .pop-btn {
        flex: 1;
        padding: 6px 10px;
        border-radius: 4px;
        font-size: 11px;
        font-weight: 600;
        text-align: center;
        cursor: pointer;
        text-decoration: none;
        display: inline-flex;
        align-items: center;
        justify-content: center;
        border: none;
      }
      .pop-btn-nav {
        background: #1a73e8;
        color: white;
      }
      .pop-btn-nav:hover { background: #1557b0; }
      .pop-btn-checkin {
        background: #f3f4f6;
        color: #374151;
        border: 1px solid #d1d5db;
      }
      .pop-btn-checkin.checked-in {
        background: #ecfdf5;
        color: #059669;
        border-color: #10b981;
        font-weight: 700;
      }

      /* 最近 10 次歷史浮動面板 */
      .history-modal-overlay {
        position: fixed;
        top: 0; left: 0; right: 0; bottom: 0;
        background: rgba(0,0,0,0.3);
        z-index: 2000;
        display: none;
        align-items: center;
        justify-content: center;
      }
      .history-modal {
        background: white;
        border-radius: 8px;
        width: 90%;
        max-width: 580px;
        max-height: 80vh;
        display: flex;
        flex-direction: column;
        box-shadow: 0 10px 25px rgba(0,0,0,0.2);
        overflow: hidden;
      }
      .history-modal-header {
        padding: 14px 18px;
        background: #f8f9fa;
        border-bottom: 1px solid var(--border-color);
        display: flex;
        justify-content: space-between;
        align-items: center;
      }
      .history-modal-title {
        font-size: 1rem;
        font-weight: 700;
        color: var(--text-primary);
        display: flex;
        align-items: center;
        gap: 6px;
      }
      .history-modal-close {
        background: none;
        border: none;
        font-size: 20px;
        color: #9aa0a6;
        cursor: pointer;
        padding: 0 4px;
      }
      .history-modal-close:hover { color: #202124; }
      .history-list {
        padding: 12px;
        overflow-y: auto;
        flex: 1;
        display: flex;
        flex-direction: column;
        gap: 8px;
      }
      .history-card {
        border: 1px solid var(--border-color);
        border-radius: 6px;
        padding: 10px 14px;
        background: #ffffff;
        display: flex;
        align-items: center;
        justify-content: space-between;
        gap: 12px;
        transition: all 120ms;
      }
      .history-card:hover {
        border-color: #1a73e8;
        background: #f8fbff;
      }
      .history-card.yesterday-last {
        border-color: #f59e0b;
        background: #fffbeb;
      }
      .history-meta {
        flex: 1;
      }
      .history-time-row {
        display: flex;
        align-items: center;
        gap: 8px;
        margin-bottom: 4px;
      }
      .history-time {
        font-weight: 700;
        font-size: 0.85rem;
        color: var(--text-primary);
      }
      .history-badge {
        font-size: 0.68rem;
        padding: 1px 6px;
        border-radius: 4px;
        font-weight: 700;
      }
      .history-badge.badge-latest {
        background: #e6f4ea;
        color: #137333;
      }
      .history-badge.badge-yesterday {
        background: #fef3c7;
        color: #b45309;
        border: 1px solid #fde68a;
      }
      .history-desc {
        font-size: 0.78rem;
        color: var(--text-secondary);
      }
      .history-counts {
        font-size: 0.75rem;
        font-weight: 600;
        color: #374151;
        margin-top: 2px;
      }
      .history-revert-btn {
        padding: 6px 12px;
        background: #1a73e8;
        color: white;
        border: none;
        border-radius: 4px;
        font-size: 0.78rem;
        font-weight: 600;
        cursor: pointer;
        flex-shrink: 0;
        transition: background 120ms;
      }
      .history-revert-btn:hover { background: #1557b0; }

      .toast {
        position: fixed;
        bottom: 24px;
        left: 50%;
        transform: translateX(-50%);
        background: #202124;
        color: white;
        padding: 9px 18px;
        border-radius: 20px;
        font-size: 0.84rem;
        font-weight: 500;
        box-shadow: 0 4px 12px rgba(0,0,0,0.25);
        z-index: 3000;
        display: none;
        animation: fadeIn 200ms ease;
      }
      @keyframes fadeIn {
        from { opacity: 0; transform: translate(-50%, 10px); }
        to { opacity: 1; transform: translate(-50%, 0); }
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
        <a href="<?= sheetUrl ?>" id="btnOpenSheetLink" target="_blank" rel="noopener" class="header-btn header-btn-sheet" title="在新分頁開啟 Google 試算表共同編輯">
          <span>📊 開啟試算表 ↗</span>
        </a>
        <button class="header-btn" id="btnSyncFromSheet" title="讀取試算表最新狀態 (背景同步)">
          <span id="syncBtnText">🔄 重新整理</span>
        </button>
      </div>
    </header>

    <div class="schedule-bar">
      <div class="schedule-left">
        <span>📅 日期：</span>
        <input type="date" id="dateSelector" class="date-input" value="<?= currentDate ?>" />
        
        <button class="sched-btn action-primary" id="btnLoadAllOpen" title="隔天快速載入全部打勾開啟">
          <span>☀️ 載入全開啟</span>
        </button>
        
        <button class="sched-btn action-history" id="btnLoadYesterday" title="一鍵回到昨天最後一次存檔紀錄">
          <span>↩️ 回到昨天最後紀錄</span>
        </button>

        <button class="sched-btn" id="btnOpenHistoryModal" title="檢視最後 10 次更新紀錄並自由回到任一版本">
          <span>📜 最近 10 次紀錄 ▼</span>
        </button>

        <button class="sched-btn" id="btnLoadCustom" title="載入選取日期的歷史設定">
          <span>📅 載入所選日期</span>
        </button>
      </div>

      <div class="schedule-right">
        <div class="autosave-pill saved" id="autoSaveIndicator" title="系統具備有動作就自動存檔功能">
          <span class="autosave-dot"></span>
          <span id="autoSaveText">已自動儲存</span>
        </div>
      </div>
    </div>

    <div class="layout-body">
      <aside class="sidebar" id="sidebar">
        <div class="panel-intro">
          <div class="intro-meta">
            <span id="activeCountLabel">目前顯示：計算中...</span>
            <span id="checkinStatsLabel" style="color:#16a34a;font-weight:700;">已打卡：0 家</span>
          </div>
          <div class="quick-actions">
            <button class="btn-chip" id="btnShowAll">全部顯示</button>
            <button class="btn-chip" id="btnHideAll">全部隱藏</button>
            <button class="btn-chip highlight-redgreen" id="btnShowRedGreen" title="僅顯示東區(綠)與西區(紅)，隱藏非488名單">
              <span>🟢🔴 紅綠同時顯示</span>
            </button>
            <button class="btn-chip" id="btnOnlyEast">僅東區 (綠)</button>
            <button class="btn-chip" id="btnOnlyWest">僅西區 (紅)</button>
            <button class="btn-chip" id="btnOnlyDiff" style="color:#b45309;">僅非488 (黃)</button>
          </div>
        </div>

        <div class="search-box">
          <div class="search-input-wrap">
            <input type="text" id="searchInput" class="search-input" placeholder="搜尋店名、路名、地址…" />
            <button id="searchClearBtn" class="search-clear">×</button>
          </div>
        </div>

        <div class="layer-tree">
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

          <div class="layer-group expanded" id="group-diff">
            <div class="group-header">
              <input type="checkbox" class="group-checkbox" id="chk-group-diff" checked title="全選/取消非488名單" />
              <div class="group-icon diff"></div>
              <div class="group-title-wrap" onclick="toggleGroupExpand('group-diff')">
                <span class="group-title">🟡 非488名單 (黃色標記)</span>
                <span class="group-stats" id="stats-diff">計算中...</span>
              </div>
              <div class="arrow-btn" onclick="toggleGroupExpand('group-diff')">▼</div>
            </div>
            <div class="item-list" id="list-diff"></div>
          </div>
        </div>
      </aside>

      <button class="toggle-sidebar-btn" id="floatingToggleBtn">
        <span id="toggleIcon">◀</span>
        <span id="toggleText">收合清單</span>
      </button>

      <div id="map"></div>
    </div>

    <!-- 最近 10 次更新歷史彈窗 -->
    <div class="history-modal-overlay" id="historyModalOverlay">
      <div class="history-modal">
        <div class="history-modal-header">
          <div class="history-modal-title">
            <span>📜 最近 10 次更新紀錄</span>
          </div>
          <button class="history-modal-close" id="historyModalCloseBtn">×</button>
        </div>
        <div class="history-list" id="historyListContainer">
          <div style="text-align:center;padding:24px;color:#888;">載入中...</div>
        </div>
      </div>
    </div>

    <div id="toastBox" class="toast">已完成打卡！</div>

    <script src="https://unpkg.com/leaflet@1.9.4/dist/leaflet.js"></script>
    <script>
      let stores = [];
      try {
        stores = JSON.parse(<?= storesJson ?>);
      } catch (e) {
        stores = ${JSON.stringify(stores)};
      }

      const currentDate = document.getElementById('dateSelector').value || '<?= currentDate ?>';
      const STORAGE_VIS_KEY = 'hualien_vis_' + currentDate;
      const STORAGE_CHK_KEY = 'hualien_chk_' + currentDate;

      let localVisMap = {};
      let localChkSet = new Set();
      try {
        const v = localStorage.getItem(STORAGE_VIS_KEY);
        if (v) localVisMap = JSON.parse(v);
        const c = localStorage.getItem(STORAGE_CHK_KEY);
        if (c) localChkSet = new Set(JSON.parse(c));
      } catch (err) {}

      function persistLocalState() {
        try {
          const visObj = {};
          storeStates.forEach((st, id) => {
            visObj[id] = st.isVisible;
          });
          localStorage.setItem(STORAGE_VIS_KEY, JSON.stringify(visObj));
          localStorage.setItem(STORAGE_CHK_KEY, JSON.stringify(Array.from(checkinSet)));
        } catch (e) {}
      }

      // ==========================================
      // 🌟 有動作就自動存檔 (Debounce 800ms)
      // ==========================================
      let autoSaveTimer = null;
      const autoSaveIndicator = document.getElementById('autoSaveIndicator');
      const autoSaveText = document.getElementById('autoSaveText');

      function updateAutoSaveStatus(status, text) {
        if (!autoSaveIndicator || !autoSaveText) return;
        autoSaveIndicator.className = 'autosave-pill ' + status;
        autoSaveText.textContent = text;
      }

      function triggerAutoSave(actionNote = '狀態變更') {
        persistLocalState();
        updateAutoSaveStatus('saving', '儲存中...');

        if (autoSaveTimer) {
          clearTimeout(autoSaveTimer);
        }

        autoSaveTimer = setTimeout(() => {
          doAutoSaveToServer(actionNote);
        }, 800);
      }

      function doAutoSaveToServer(actionNote) {
        if (typeof google === 'undefined' || !google.script || !google.script.run) {
          updateAutoSaveStatus('saved', '本機已儲存');
          return;
        }

        const hiddenIds = [];
        storeStates.forEach((st, id) => {
          if (!st.isVisible) hiddenIds.push(id);
        });
        const checkedInIds = Array.from(checkinSet);
        const curDate = document.getElementById('dateSelector').value || currentDate;

        google.script.run
          .withSuccessHandler((res) => {
            if (res && res.success) {
              const timeStr = res.timestamp ? res.timestamp.split(' ').pop() : '';
              updateAutoSaveStatus('saved', '已自動儲存 ' + timeStr);
            } else {
              updateAutoSaveStatus('saved', '已存於快取');
            }
          })
          .withFailureHandler((err) => {
            updateAutoSaveStatus('saved', '已暫存 (網路稍後重試)');
          })
          .apiAutoSaveFullState(curDate, hiddenIds, checkedInIds, actionNote);
      }

      // ==========================================
      // 地圖初始化
      // ==========================================
      const map = L.map('map', { zoomControl: true }).setView([23.985, 121.605], 13);
      L.tileLayer('https://{s}.tile.openstreetmap.org/{z}/{x}/{y}.png', {
        maxZoom: 19,
        attribution: '&copy; OpenStreetMap & Google Sheets Sync'
      }).addTo(map);

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
      const iconDiff = createPinIcon('#f59e0b');

      function getSmallPinSvg(colorHex) {
        return \`<svg class="pin-svg" viewBox="0 0 24 34" fill="none"><path d="M12 0C5.372 0 0 5.373 0 12C0 19.5 12 34 12 34C12 34 24 19.5 24 12C24 5.373 18.628 0 12 0Z" fill="\${colorHex}" stroke="#ffffff" stroke-width="1.5"/><circle cx="12" cy="11" r="4" fill="#ffffff"/></svg>\`;
      }

      function showToast(text, duration = 2500) {
        const toast = document.getElementById('toastBox');
        toast.textContent = text;
        toast.style.display = 'block';
        setTimeout(() => { toast.style.display = 'none'; }, duration);
      }

      const checkinSet = new Set();
      stores.forEach(s => {
        if (localChkSet.has(s.id) || s.isCheckedIn) {
          checkinSet.add(s.id);
        }
      });

      const storeStates = new Map();

      function getGoogleMapsNavUrl(s) {
        return 'https://www.google.com/maps/dir/?api=1&destination=' + s.lat + ',' + s.lng;
      }

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
          badgeText = '🟡 非 488 名單店家';
        }

        const marker = L.marker([store.lat, store.lng], { icon: icon });

        function updatePopupContent() {
          const isCheckedIn = checkinSet.has(store.id);
          const popupDiv = document.createElement('div');
          popupDiv.innerHTML = \`
            <div class="pop-header">
              <span class="pop-badge \${badgeClass}">\${badgeText}</span>
              <span style="font-size:0.75rem;color:#888;">#\${index + 1}</span>
            </div>
            <div class="pop-title">\${store.name}</div>
            <div class="pop-addr">\${store.address}\${store.estimated ? '<br/><span style="color:#b45309;font-size:11px;">※ 位置為地址估算</span>' : ''}</div>
            <div class="pop-actions">
              <a class="pop-btn pop-btn-nav" href="\${getGoogleMapsNavUrl(store)}" target="_blank" rel="noopener">Google 導航 ↗</a>
              <button class="pop-btn pop-btn-checkin \${isCheckedIn ? 'checked-in' : ''}" id="pop-checkin-\${store.id}">
                \${isCheckedIn ? '✅ 已打卡' : '📍 打卡'}
              </button>
            </div>
          \`;

          const btn = popupDiv.querySelector('#pop-checkin-' + store.id);
          btn.addEventListener('click', (e) => {
            e.stopPropagation();
            toggleCheckin(store.id);
            updatePopupContent();
            marker.setPopupContent(popupDiv);
          });

          return popupDiv;
        }

        marker.bindPopup(updatePopupContent());
        marker.on('click', () => {
          marker.setPopupContent(updatePopupContent());
          highlightStoreRow(store.id);
        });

        let isInitiallyVisible = true;
        if (localVisMap[store.id] !== undefined) {
          isInitiallyVisible = localVisMap[store.id];
        } else if (store.isVisible !== undefined) {
          isInitiallyVisible = store.isVisible;
        }

        if (isInitiallyVisible) {
          marker.addTo(map);
        }

        storeStates.set(store.id, {
          store,
          marker,
          isVisible: isInitiallyVisible,
          rowElement: null,
          chkElement: null,
          checkinBtn: null
        });
      });

      // 側邊欄清單生成
      const listEast = document.getElementById('list-east');
      const listWest = document.getElementById('list-west');
      const listDiff = document.getElementById('list-diff');

      stores.forEach((store, index) => {
        const state = storeStates.get(store.id);
        const row = document.createElement('div');
        row.className = 'store-row' + (state.isVisible ? '' : ' disabled');
        row.id = 'store-row-' + store.id;

        const pinColor = store.subGroup === '主要名單-東區' ? '#2e7d32' : (store.subGroup === '主要名單-西區' ? '#d32f2f' : '#f59e0b');
        const isCheckedIn = checkinSet.has(store.id);

        row.innerHTML = \`
          <input type="checkbox" class="item-checkbox" id="chk-store-\${store.id}" \${state.isVisible ? 'checked' : ''} />
          \${getSmallPinSvg(pinColor)}
          <div class="store-info">
            <div class="store-name-text" title="\${store.name}">\${index + 1}. \${store.name}</div>
            <div class="store-addr-text" title="\${store.address}">\${store.address}</div>
          </div>
          <button class="checkin-btn \${isCheckedIn ? 'checked-in' : ''}" id="btn-chk-\${store.id}" title="點擊打卡">
            \${isCheckedIn ? '✅ 已打卡' : '📍 打卡'}
          </button>
        \`;

        const chk = row.querySelector('.item-checkbox');
        const btnCheckin = row.querySelector('#btn-chk-' + store.id);
        state.rowElement = row;
        state.chkElement = chk;
        state.checkinBtn = btnCheckin;

        chk.addEventListener('click', (e) => {
          e.stopPropagation();
          setStoreVisibility(store.id, chk.checked);
          triggerAutoSave('變更店家【' + store.name + '】顯示狀態');
        });

        btnCheckin.addEventListener('click', (e) => {
          e.stopPropagation();
          toggleCheckin(store.id);
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

      function toggleCheckin(storeId) {
        const state = storeStates.get(storeId);
        if (!state) return;

        let isNowCheckedIn = false;
        if (checkinSet.has(storeId)) {
          checkinSet.delete(storeId);
          state.checkinBtn.classList.remove('checked-in');
          state.checkinBtn.innerHTML = '📍 打卡';
          showToast('已取消【' + state.store.name + '】打卡');
          isNowCheckedIn = false;
        } else {
          checkinSet.add(storeId);
          state.checkinBtn.classList.add('checked-in');
          state.checkinBtn.innerHTML = '✅ 已打卡';
          showToast('✅ 已完成【' + state.store.name + '】打卡！');
          isNowCheckedIn = true;
        }

        updateCheckinStats();
        triggerAutoSave((isNowCheckedIn ? '完成打卡：' : '取消打卡：') + state.store.name);
      }

      function updateCheckinStats() {
        const label = document.getElementById('checkinStatsLabel');
        if (label) {
          label.textContent = '已打卡：' + checkinSet.size + ' 家';
        }
      }

      function setStoreVisibility(storeId, isVisible) {
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
      }

      function focusStoreOnMap(storeId) {
        const state = storeStates.get(storeId);
        if (!state) return;
        if (!state.isVisible) {
          setStoreVisibility(storeId, true);
          triggerAutoSave('點擊開啟【' + state.store.name + '】');
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
        stores.filter(s => s.subGroup === groupName).forEach(s => {
          setStoreVisibility(s.id, isVisible);
        });
        updateGroupCheckboxes();
        updateStats();
        triggerAutoSave((isVisible ? '開啟' : '關閉') + '群組：' + groupName);
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

      // 快捷按鈕
      document.getElementById('btnShowAll').addEventListener('click', () => {
        stores.forEach(s => setStoreVisibility(s.id, true));
        triggerAutoSave('全部店家開啟顯示 (500筆)');
      });
      document.getElementById('btnHideAll').addEventListener('click', () => {
        stores.forEach(s => setStoreVisibility(s.id, false));
        triggerAutoSave('全部店家隱藏關閉 (0筆)');
      });

      document.getElementById('btnShowRedGreen').addEventListener('click', () => {
        stores.forEach(s => {
          const isRedOrGreen = (s.subGroup === '主要名單-東區' || s.subGroup === '主要名單-西區');
          setStoreVisibility(s.id, isRedOrGreen);
        });
        showToast('已切換為：🟢🔴 紅綠主要名單同時顯示 (388筆)');
        triggerAutoSave('🟢🔴 紅綠主要名單同時顯示 (388筆)');
      });

      document.getElementById('btnOnlyEast').addEventListener('click', () => {
        stores.forEach(s => setStoreVisibility(s.id, s.subGroup === '主要名單-東區'));
        triggerAutoSave('僅顯示東區 (綠色206筆)');
      });
      document.getElementById('btnOnlyWest').addEventListener('click', () => {
        stores.forEach(s => setStoreVisibility(s.id, s.subGroup === '主要名單-西區'));
        triggerAutoSave('僅顯示西區 (紅色182筆)');
      });
      document.getElementById('btnOnlyDiff').addEventListener('click', () => {
        stores.forEach(s => setStoreVisibility(s.id, s.subGroup === '非488名單'));
        triggerAutoSave('僅顯示非488名單 (黃色112筆)');
      });

      // ----------------------------------------------------
      // 🌟 載入全開啟
      // ----------------------------------------------------
      document.getElementById('btnLoadAllOpen').addEventListener('click', () => {
        if (!confirm('確定要將今天的所有店家設為【全部開啟顯示】嗎？')) return;
        showToast('正在載入全開啟並同步更新...', 2000);
        stores.forEach(s => setStoreVisibility(s.id, true));
        triggerAutoSave('☀️ 載入全開啟 (500筆全部顯示)');
      });

      // ----------------------------------------------------
      // 🌟 核心升級：一鍵回到昨天最後紀錄 (鎖定昨日最新存檔)
      // ----------------------------------------------------
      document.getElementById('btnLoadYesterday').addEventListener('click', () => {
        if (!confirm('確定要一鍵載入【昨天的最後一次存檔紀錄】嗎？')) return;
        showToast('正在讀取昨日最後紀錄...', 3000);

        if (typeof google !== 'undefined' && google.script && google.script.run) {
          const currentInputDate = document.getElementById('dateSelector').value;
          google.script.run
            .withSuccessHandler((res) => {
              if (res && res.success && res.stores) {
                localVisMap = {};
                res.stores.forEach(s => {
                  setStoreVisibility(s.id, s.isVisible);
                });
                persistLocalState();
                updateAutoSaveStatus('saved', '已載入昨日最後紀錄');
                showToast(\`✅ 成功回到昨日最後存檔！開啟：\${res.activeCount} 家，關閉：\${res.closedCount} 家\`, 3500);
              } else {
                showToast(res.msg || '找不到昨日歷史存檔紀錄', 3000);
              }
            })
            .withFailureHandler((err) => {
              showToast('讀取失敗: ' + err, 2500);
            })
            .apiLoadYesterdayLastRecord(currentInputDate);
        } else {
          alert('請在 Google Apps Script 發佈環境下執行！');
        }
      });

      // ----------------------------------------------------
      // 🌟 核心升級：最近 10 次更新歷史彈窗與版本自由回溯
      // ----------------------------------------------------
      const historyModalOverlay = document.getElementById('historyModalOverlay');
      const historyListContainer = document.getElementById('historyListContainer');
      const btnOpenHistoryModal = document.getElementById('btnOpenHistoryModal');
      const historyModalCloseBtn = document.getElementById('historyModalCloseBtn');

      function openHistoryModal() {
        historyModalOverlay.style.display = 'flex';
        historyListContainer.innerHTML = '<div style="text-align:center;padding:24px;color:#888;">⏳ 正在讀取最近 10 次更新版本...</div>';

        if (typeof google !== 'undefined' && google.script && google.script.run) {
          google.script.run
            .withSuccessHandler((res) => {
              if (!res || !res.success || !res.list || res.list.length === 0) {
                historyListContainer.innerHTML = '<div style="text-align:center;padding:24px;color:#888;">尚無歷史更新紀錄</div>';
                return;
              }
              renderHistoryList(res.list);
            })
            .withFailureHandler((err) => {
              historyListContainer.innerHTML = '<div style="text-align:center;padding:24px;color:#c5221f;">讀取失敗: ' + err + '</div>';
            })
            .apiGetRecentHistory(10);
        } else {
          historyListContainer.innerHTML = \`
            <div style="text-align:center;padding:20px;color:#666;">
              <p>【本機離線展示模式】</p>
              <p style="font-size:12px;margin-top:6px;">發布至 Google Apps Script 後將列出最後 10 次更新歷程！</p>
            </div>
          \`;
        }
      }

      function renderHistoryList(list) {
        historyListContainer.innerHTML = '';
        list.forEach((item, idx) => {
          const card = document.createElement('div');
          card.className = 'history-card' + (item.isYesterdayLast ? ' yesterday-last' : '');

          let badgeHtml = '';
          if (item.isLatest) {
            badgeHtml = '<span class="history-badge badge-latest">最新版本</span>';
          } else if (item.isYesterdayLast) {
            badgeHtml = '<span class="history-badge badge-yesterday">⭐ 昨天的最後紀錄</span>';
          }

          card.innerHTML = \`
            <div class="history-meta">
              <div class="history-time-row">
                <span class="history-time">\${item.timestamp}</span>
                \${badgeHtml}
              </div>
              <div class="history-desc">動作：\${item.action || '狀態更新'}</div>
              <div class="history-counts">📊 開啟：\${item.activeCount} 家 ｜ 關閉：\${item.closedCount} 家</div>
            </div>
            <button class="history-revert-btn" id="btn-revert-\${item.rowIndex}">
              ↩️ 回到此版本
            </button>
          \`;

          const btnRevert = card.querySelector('#btn-revert-' + item.rowIndex);
          btnRevert.addEventListener('click', () => {
            revertToVersion(item);
          });

          historyListContainer.appendChild(card);
        });
      }

      function revertToVersion(item) {
        if (!confirm('確定要將店家狀態回復至【' + item.timestamp + '】嗎？')) return;
        showToast('正在回復至該版本...', 2500);

        if (typeof google !== 'undefined' && google.script && google.script.run) {
          const curDate = document.getElementById('dateSelector').value;
          google.script.run
            .withSuccessHandler((res) => {
              if (res && res.success && res.stores) {
                localVisMap = {};
                res.stores.forEach(s => {
                  setStoreVisibility(s.id, s.isVisible);
                });
                persistLocalState();
                historyModalOverlay.style.display = 'none';
                updateAutoSaveStatus('saved', '已回復至 ' + item.timestamp);
                showToast(\`✅ 成功回復版本！開啟：\${res.activeCount} 家，關閉：\${res.closedCount} 家\`, 3500);
              } else {
                showToast(res.msg || '版本回復失敗', 2500);
              }
            })
            .withFailureHandler((err) => {
              showToast('回復失敗: ' + err, 2500);
            })
            .apiRevertToHistoryVersion(item.rowIndex, curDate);
        }
      }

      btnOpenHistoryModal.addEventListener('click', openHistoryModal);
      historyModalCloseBtn.addEventListener('click', () => {
        historyModalOverlay.style.display = 'none';
      });
      historyModalOverlay.addEventListener('click', (e) => {
        if (e.target === historyModalOverlay) historyModalOverlay.style.display = 'none';
      });

      // 載入指定日期的設定
      document.getElementById('btnLoadCustom').addEventListener('click', () => {
        const chosenDate = document.getElementById('dateSelector').value;
        if (!chosenDate) {
          alert('請先選擇日期！');
          return;
        }
        if (confirm('確定要載入指定日期【' + chosenDate + '】的開關設定嗎？')) {
          showToast('正在載入指定日期設定...', 3000);
          if (typeof google !== 'undefined' && google.script && google.script.run) {
            google.script.run
              .withSuccessHandler((res) => {
                if (res && res.success && res.stores) {
                  localVisMap = {};
                  res.stores.forEach(s => {
                    setStoreVisibility(s.id, s.isVisible);
                  });
                  persistLocalState();
                  showToast(\`✅ 成功載入 \${res.sourceDate} 設定！開啟：\${res.activeCount} 家，關閉：\${res.closedCount} 家\`, 3500);
                } else {
                  showToast('無該日期的歷史紀錄', 2500);
                }
              })
              .withFailureHandler((err) => {
                showToast('讀取失敗: ' + err, 2500);
              })
              .loadStatusFromDate(chosenDate, chosenDate);
          }
        }
      });

      // 重新整理
      function refreshDataAsync(dateToFetch) {
        const syncBtnText = document.getElementById('syncBtnText');
        if (syncBtnText) syncBtnText.textContent = '⏳ 同步中...';
        showToast('正在同步最新狀態...', 2000);

        if (typeof google !== 'undefined' && google.script && google.script.run) {
          const targetDate = dateToFetch || document.getElementById('dateSelector').value;
          google.script.run
            .withSuccessHandler((res) => {
              if (syncBtnText) syncBtnText.textContent = '🔄 重新整理';
              if (res && res.success && res.stores) {
                res.stores.forEach(s => {
                  setStoreVisibility(s.id, s.isVisible);
                  const state = storeStates.get(s.id);
                  if (state) {
                    if (s.isCheckedIn) {
                      checkinSet.add(s.id);
                      state.checkinBtn.classList.add('checked-in');
                      state.checkinBtn.innerHTML = '✅ 已打卡';
                    } else {
                      checkinSet.delete(s.id);
                      state.checkinBtn.classList.remove('checked-in');
                      state.checkinBtn.innerHTML = '📍 打卡';
                    }
                  }
                });
                persistLocalState();
                updateCheckinStats();
                showToast('✅ 最新狀態同步完成！', 2000);
              }
            })
            .withFailureHandler((err) => {
              if (syncBtnText) syncBtnText.textContent = '🔄 重新整理';
              showToast('同步失敗: ' + err, 2500);
            })
            .apiGetStoresForDate(targetDate);
        } else {
          if (syncBtnText) syncBtnText.textContent = '🔄 重新整理';
          showToast('目前為本機展示模式', 1500);
        }
      }

      document.getElementById('btnSyncFromSheet').addEventListener('click', () => {
        refreshDataAsync();
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
      updateCheckinStats();
    </script>
  </body>
</html>
`;

fs.writeFileSync(path.join(gasDir, "Index.html"), indexHtml, "utf8");
console.log("gas/Index.html 已成功生成 (含動作自動更新與最後10次版本歷程)！");
