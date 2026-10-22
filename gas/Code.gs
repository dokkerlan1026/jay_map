/**
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
  var m = str.match(/(\d{4})[-/](\d{1,2})[-/](\d{1,2})/);
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
  var digits = String(rawId).replace(/\D/g, '');
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
          newNote = curNote.replace(/已打卡\s*\d{2}:\d{2}/, "").trim();
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
      ui.alert("載入完成！", "已成功載入昨日（" + out.yesterdayDate + "）的最後存檔！\n- 存檔時間：" + out.yesterdayTimestamp + "\n- 開啟店家數：" + out.activeCount + " 家\n- 關閉店家數：" + out.closedCount + " 家", ui.ButtonSet.OK);
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
  var msg = "【最近更新紀錄 (最後 " + res.list.length + " 次)】\n\n";
  res.list.forEach(function(item, idx) {
    var tag = item.isLatest ? " [最新]" : (item.isYesterdayLast ? " ⭐[昨天最後]" : "");
    msg += (idx + 1) + ". " + item.timestamp + tag + "\n";
    msg += "   - 動作：" + item.action + "\n";
    msg += "   - 狀態：開啟 " + item.activeCount + " 家 / 關閉 " + item.closedCount + " 家\n\n";
  });
  msg += "提示：您也可以在網頁 Web App 上方點擊【📜 最近 10 次紀錄】一鍵復原！";
  ui.alert("最近更新紀錄", msg, ui.ButtonSet.OK);
}

function menuLoadCustomDate() {
  var ui = SpreadsheetApp.getUi();
  var prompt = ui.prompt("載入指定日期", "請輸入要載入的歷史日期 (格式：YYYY-MM-DD，例如 2026-10-21)：", ui.ButtonSet.OK_CANCEL);
  if (prompt.getSelectedButton() === ui.Button.OK) {
    var srcDate = normalizeDateString(prompt.getResponseText().trim());
    if (!/^\d{4}-\d{2}-\d{2}$/.test(srcDate)) {
      ui.alert("錯誤", "日期格式不正確，請使用 YYYY-MM-DD 格式！", ui.ButtonSet.OK);
      return;
    }
    var today = getTodayString();
    var out = loadStatusFromDate(srcDate, today);
    ui.alert("載入完成！", "已成功載入 " + srcDate + " 的設定！\n- 開啟店家數：" + out.activeCount + " 家\n- 關閉店家數：" + out.closedCount + " 家", ui.ButtonSet.OK);
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
  ui.alert("初始化完成！", "已成功建立【當日店家清單】與【歷史紀錄存檔】！\n\n1. B 欄支援打勾即時切換顯示\n2. 網頁端任何動作皆會自動存檔至試算表\n3. 自動保留最後 10 次更新紀錄，方便隨時回到昨天或前幾次更新！", ui.ButtonSet.OK);
}

function getDefaultStores() {
  return [
  {
    "id": "store-1",
    "name": "蘇姵綸營業所",
    "address": "花蓮縣吉安鄉中山路50號(G49.G50攤位)",
    "lat": 23.97419,
    "lng": 121.611724,
    "estimated": true,
    "isPrimary488": true,
    "group": "主要488名單",
    "delta": 0.014550538461534757,
    "subGroup": "主要名單-東區",
    "region": "界線以東",
    "color": "綠色",
    "colorHex": "#257a4e"
  },
  {
    "id": "store-2",
    "name": "梅珍香食品原料有限公司",
    "address": "花蓮縣吉安鄉中原路1段128號",
    "lat": 23.979014,
    "lng": 121.58844,
    "estimated": false,
    "isPrimary488": true,
    "group": "主要488名單",
    "delta": -0.014114076923064545,
    "subGroup": "主要名單-西區",
    "region": "界線以西",
    "color": "紅色",
    "colorHex": "#e53e3e"
  },
  {
    "id": "store-3",
    "name": "日品模型玩具",
    "address": "花蓮縣吉安鄉中原路1段131號",
    "lat": 23.978708,
    "lng": 121.588141,
    "estimated": false,
    "isPrimary488": true,
    "group": "主要488名單",
    "delta": -0.014071769230781683,
    "subGroup": "主要名單-西區",
    "region": "界線以西",
    "color": "紅色",
    "colorHex": "#e53e3e"
  },
  {
    "id": "store-4",
    "name": "隆昌寵物用品有限公司",
    "address": "花蓮縣吉安鄉中原路1段86號旁邊",
    "lat": 23.983037,
    "lng": 121.580642,
    "estimated": true,
    "isPrimary488": true,
    "group": "主要488名單",
    "delta": -0.0263992692307653,
    "subGroup": "主要名單-西區",
    "region": "界線以西",
    "color": "紅色",
    "colorHex": "#e53e3e"
  },
  {
    "id": "store-5",
    "name": "廣隆進企業有限公司",
    "address": "花蓮縣吉安鄉中原路82號",
    "lat": 23.964778,
    "lng": 121.60324,
    "estimated": false,
    "isPrimary488": false,
    "group": "非488名單(另色標記)",
    "subGroup": "非488名單",
    "region": "非488店家",
    "color": "黃色",
    "colorHex": "#f59e0b"
  },
  {
    "id": "store-6",
    "name": "相正企業有限公司",
    "address": "花蓮縣吉安鄉民治路136號",
    "lat": 23.979787,
    "lng": 121.591367,
    "estimated": false,
    "isPrimary488": true,
    "group": "主要488名單",
    "delta": -0.012049269230772097,
    "subGroup": "主要名單-西區",
    "region": "界線以西",
    "color": "紅色",
    "colorHex": "#e53e3e"
  },
  {
    "id": "store-7",
    "name": "莊立良公館",
    "address": "花蓮縣吉安鄉民治路2巷1弄2號",
    "lat": 23.979066,
    "lng": 121.593721,
    "estimated": false,
    "isPrimary488": true,
    "group": "主要488名單",
    "delta": -0.00889107692307789,
    "subGroup": "主要名單-西區",
    "region": "界線以西",
    "color": "紅色",
    "colorHex": "#e53e3e"
  },
  {
    "id": "store-8",
    "name": "黃世傑公館",
    "address": "花蓮縣吉安鄉民治路2巷1弄70號",
    "lat": 23.979613,
    "lng": 121.592414,
    "estimated": false,
    "isPrimary488": true,
    "group": "主要488名單",
    "delta": -0.010808192307692366,
    "subGroup": "主要名單-西區",
    "region": "界線以西",
    "color": "紅色",
    "colorHex": "#e53e3e"
  },
  {
    "id": "store-9",
    "name": "睿昕實業有限公司",
    "address": "花蓮縣吉安鄉民治路6號",
    "lat": 23.978535,
    "lng": 121.593287,
    "estimated": false,
    "isPrimary488": true,
    "group": "主要488名單",
    "delta": -0.00873280769231144,
    "subGroup": "主要名單-西區",
    "region": "界線以西",
    "color": "紅色",
    "colorHex": "#e53e3e"
  },
  {
    "id": "store-10",
    "name": "鴻龍行",
    "address": "花蓮縣吉安鄉吉安鄉民治路90號",
    "lat": 23.979023,
    "lng": 121.592539,
    "estimated": false,
    "isPrimary488": true,
    "group": "主要488名單",
    "delta": -0.010025115384621586,
    "subGroup": "主要名單-西區",
    "region": "界線以西",
    "color": "紅色",
    "colorHex": "#e53e3e"
  },
  {
    "id": "store-11",
    "name": "鄭啟聰公館",
    "address": "花蓮縣吉安鄉自立路1段148巷18號",
    "lat": 23.979909,
    "lng": 121.592707,
    "estimated": false,
    "isPrimary488": true,
    "group": "主要488名單",
    "delta": -0.010845346153843138,
    "subGroup": "主要名單-西區",
    "region": "界線以西",
    "color": "紅色",
    "colorHex": "#e53e3e"
  },
  {
    "id": "store-12",
    "name": "趙添福公館",
    "address": "花蓮縣吉安鄉自立路1段148巷22號",
    "lat": 23.979679,
    "lng": 121.592873,
    "estimated": false,
    "isPrimary488": true,
    "group": "主要488名單",
    "delta": -0.010422807692307856,
    "subGroup": "主要名單-西區",
    "region": "界線以西",
    "color": "紅色",
    "colorHex": "#e53e3e"
  },
  {
    "id": "store-13",
    "name": "合隆汽車音響有限公司",
    "address": "花蓮縣吉安鄉自立路2段39巷21號",
    "lat": 23.982155,
    "lng": 121.592474,
    "estimated": false,
    "isPrimary488": true,
    "group": "主要488名單",
    "delta": -0.013583499999995752,
    "subGroup": "主要名單-西區",
    "region": "界線以西",
    "color": "紅色",
    "colorHex": "#e53e3e"
  },
  {
    "id": "store-14",
    "name": "統冠超商管理中心庫房",
    "address": "花蓮縣吉安鄉自立路2段39巷72-2號",
    "lat": 23.983085,
    "lng": 121.591435,
    "estimated": false,
    "isPrimary488": true,
    "group": "主要488名單",
    "delta": -0.015659807692301797,
    "subGroup": "主要名單-西區",
    "region": "界線以西",
    "color": "紅色",
    "colorHex": "#e53e3e"
  },
  {
    "id": "store-15",
    "name": "統冠超商管理中心",
    "address": "花蓮縣吉安鄉自立路2段39巷72-2號",
    "lat": 23.983085,
    "lng": 121.591435,
    "estimated": false,
    "isPrimary488": true,
    "group": "主要488名單",
    "delta": -0.015659807692301797,
    "subGroup": "主要名單-西區",
    "region": "界線以西",
    "color": "紅色",
    "colorHex": "#e53e3e"
  },
  {
    "id": "store-16",
    "name": "相正企業有限公司倉庫",
    "address": "花蓮縣吉安鄉自立路2段39巷72-3號",
    "lat": 23.983085,
    "lng": 121.591435,
    "estimated": false,
    "isPrimary488": true,
    "group": "主要488名單",
    "delta": -0.015659807692301797,
    "subGroup": "主要名單-西區",
    "region": "界線以西",
    "color": "紅色",
    "colorHex": "#e53e3e"
  },
  {
    "id": "store-17",
    "name": "自立菸酒商行",
    "address": "花蓮縣吉安鄉自立路2段42號",
    "lat": 23.982054,
    "lng": 121.593464,
    "estimated": false,
    "isPrimary488": true,
    "group": "主要488名單",
    "delta": -0.012480846153849257,
    "subGroup": "主要名單-西區",
    "region": "界線以西",
    "color": "紅色",
    "colorHex": "#e53e3e"
  },
  {
    "id": "store-18",
    "name": "富甲里實業有限公司",
    "address": "花蓮縣吉安鄉自強路184號1樓",
    "lat": 23.97996,
    "lng": 121.587488,
    "estimated": false,
    "isPrimary488": true,
    "group": "主要488名單",
    "delta": -0.016121230769229555,
    "subGroup": "主要名單-西區",
    "region": "界線以西",
    "color": "紅色",
    "colorHex": "#e53e3e"
  },
  {
    "id": "store-19",
    "name": "劉銓二手車行",
    "address": "花蓮縣吉安鄉自強路418-1號旁",
    "lat": 23.983558,
    "lng": 121.58961,
    "estimated": true,
    "isPrimary488": false,
    "group": "非488名單(另色標記)",
    "subGroup": "非488名單",
    "region": "非488店家",
    "color": "黃色",
    "colorHex": "#f59e0b"
  },
  {
    "id": "store-20",
    "name": "陸德環保有限公司",
    "address": "花蓮縣吉安鄉自強路471號",
    "lat": 23.985029,
    "lng": 121.590493,
    "estimated": false,
    "isPrimary488": false,
    "group": "非488名單(另色標記)",
    "subGroup": "非488名單",
    "region": "非488店家",
    "color": "黃色",
    "colorHex": "#f59e0b"
  },
  {
    "id": "store-21",
    "name": "陸德環保有限公司",
    "address": "花蓮縣吉安鄉自強路471號",
    "lat": 23.985029,
    "lng": 121.590493,
    "estimated": false,
    "isPrimary488": false,
    "group": "非488名單(另色標記)",
    "subGroup": "非488名單",
    "region": "非488店家",
    "color": "黃色",
    "colorHex": "#f59e0b"
  },
  {
    "id": "store-22",
    "name": "伊蕾名店花蓮吉安自強店",
    "address": "花蓮縣吉安鄉自強路504號",
    "lat": 23.986038,
    "lng": 121.59159,
    "estimated": false,
    "isPrimary488": true,
    "group": "主要488名單",
    "delta": -0.018798538461538783,
    "subGroup": "主要名單-西區",
    "region": "界線以西",
    "color": "紅色",
    "colorHex": "#e53e3e"
  },
  {
    "id": "store-23",
    "name": "台新銀行ATM花蓮統冠超市自強",
    "address": "花蓮縣吉安鄉自強路532號",
    "lat": 23.98685,
    "lng": 121.592203,
    "estimated": false,
    "isPrimary488": false,
    "group": "非488名單(另色標記)",
    "subGroup": "非488名單",
    "region": "非488店家",
    "color": "黃色",
    "colorHex": "#f59e0b"
  },
  {
    "id": "store-24",
    "name": "統冠超商－自強店",
    "address": "花蓮縣吉安鄉自強路532號",
    "lat": 23.98685,
    "lng": 121.592203,
    "estimated": false,
    "isPrimary488": false,
    "group": "非488名單(另色標記)",
    "subGroup": "非488名單",
    "region": "非488店家",
    "color": "黃色",
    "colorHex": "#f59e0b"
  },
  {
    "id": "store-25",
    "name": "林妍容營業所",
    "address": "花蓮縣吉安鄉自強路552號",
    "lat": 23.987036,
    "lng": 121.592358,
    "estimated": false,
    "isPrimary488": true,
    "group": "主要488名單",
    "delta": -0.019143692307679316,
    "subGroup": "主要名單-西區",
    "region": "界線以西",
    "color": "紅色",
    "colorHex": "#e53e3e"
  },
  {
    "id": "store-26",
    "name": "花蓮二信建國分社",
    "address": "花蓮縣吉安鄉自強路558號",
    "lat": 23.987097,
    "lng": 121.59245,
    "estimated": false,
    "isPrimary488": true,
    "group": "主要488名單",
    "delta": -0.019119730769233456,
    "subGroup": "主要名單-西區",
    "region": "界線以西",
    "color": "紅色",
    "colorHex": "#e53e3e"
  },
  {
    "id": "store-27",
    "name": "孫振剛-公館",
    "address": "花蓮縣吉安鄉自強路578號",
    "lat": 23.987489,
    "lng": 121.592806,
    "estimated": false,
    "isPrimary488": true,
    "group": "主要488名單",
    "delta": -0.019200961538459183,
    "subGroup": "主要名單-西區",
    "region": "界線以西",
    "color": "紅色",
    "colorHex": "#e53e3e"
  },
  {
    "id": "store-28",
    "name": "吳秀娟營業所",
    "address": "花蓮縣吉安鄉自強路590號1樓",
    "lat": 23.987694,
    "lng": 121.592985,
    "estimated": false,
    "isPrimary488": false,
    "group": "非488名單(另色標記)",
    "subGroup": "非488名單",
    "region": "非488店家",
    "color": "黃色",
    "colorHex": "#f59e0b"
  },
  {
    "id": "store-29",
    "name": "上和不動產",
    "address": "花蓮縣吉安鄉和平路1段116號",
    "lat": 23.982601,
    "lng": 121.590329,
    "estimated": false,
    "isPrimary488": false,
    "group": "非488名單(另色標記)",
    "subGroup": "非488名單",
    "region": "非488店家",
    "color": "黃色",
    "colorHex": "#f59e0b"
  },
  {
    "id": "store-30",
    "name": "昇輝汽車有限公司",
    "address": "花蓮縣吉安鄉和平路1段147號",
    "lat": 23.982674,
    "lng": 121.58956,
    "estimated": false,
    "isPrimary488": true,
    "group": "主要488名單",
    "delta": -0.01707638461537897,
    "subGroup": "主要名單-西區",
    "region": "界線以西",
    "color": "紅色",
    "colorHex": "#e53e3e"
  },
  {
    "id": "store-31",
    "name": "凱娜汽車美研社通報",
    "address": "花蓮縣吉安鄉和平路1段149號1樓",
    "lat": 23.982771,
    "lng": 121.589428,
    "estimated": false,
    "isPrimary488": false,
    "group": "非488名單(另色標記)",
    "subGroup": "非488名單",
    "region": "非488店家",
    "color": "黃色",
    "colorHex": "#f59e0b"
  },
  {
    "id": "store-32",
    "name": "凱娜汽車美研社",
    "address": "花蓮縣吉安鄉和平路1段149號1樓",
    "lat": 23.982771,
    "lng": 121.589428,
    "estimated": false,
    "isPrimary488": false,
    "group": "非488名單(另色標記)",
    "subGroup": "非488名單",
    "region": "非488店家",
    "color": "黃色",
    "colorHex": "#f59e0b"
  },
  {
    "id": "store-33",
    "name": "中油巨豐加油站",
    "address": "花蓮縣吉安鄉和平路1段45號",
    "lat": 23.981371,
    "lng": 121.591977,
    "estimated": false,
    "isPrimary488": true,
    "group": "主要488名單",
    "delta": -0.013206038461532899,
    "subGroup": "主要名單-西區",
    "region": "界線以西",
    "color": "紅色",
    "colorHex": "#e53e3e"
  },
  {
    "id": "store-34",
    "name": "言晟股份有限公司",
    "address": "花蓮縣吉安鄉建國路1段117巷10號(花蓮水15-20)",
    "lat": 23.98484,
    "lng": 121.592124,
    "estimated": false,
    "isPrimary488": true,
    "group": "主要488名單",
    "delta": -0.016928307692310796,
    "subGroup": "主要名單-西區",
    "region": "界線以西",
    "color": "紅色",
    "colorHex": "#e53e3e"
  },
  {
    "id": "store-35",
    "name": "國泰世華銀行ATM全聯花蓮建國店",
    "address": "花蓮縣吉安鄉建國路1段119,119-3號",
    "lat": 23.985433,
    "lng": 121.592366,
    "estimated": false,
    "isPrimary488": true,
    "group": "主要488名單",
    "delta": -0.017347730769230907,
    "subGroup": "主要名單-西區",
    "region": "界線以西",
    "color": "紅色",
    "colorHex": "#e53e3e"
  },
  {
    "id": "store-36",
    "name": "燦坤實業花蓮建國店",
    "address": "花蓮縣吉安鄉建國路1段119-2號",
    "lat": 23.985433,
    "lng": 121.592366,
    "estimated": false,
    "isPrimary488": true,
    "group": "主要488名單",
    "delta": -0.017347730769230907,
    "subGroup": "主要名單-西區",
    "region": "界線以西",
    "color": "紅色",
    "colorHex": "#e53e3e"
  },
  {
    "id": "store-37",
    "name": "全聯花蓮建國2114",
    "address": "花蓮縣吉安鄉建國路1段119號",
    "lat": 23.985433,
    "lng": 121.592366,
    "estimated": false,
    "isPrimary488": true,
    "group": "主要488名單",
    "delta": -0.017347730769230907,
    "subGroup": "主要名單-西區",
    "region": "界線以西",
    "color": "紅色",
    "colorHex": "#e53e3e"
  },
  {
    "id": "store-38",
    "name": "大樹藥局花蓮吉安加盟店",
    "address": "花蓮縣吉安鄉建國路1段137之1號",
    "lat": 23.985742,
    "lng": 121.591431,
    "estimated": false,
    "isPrimary488": false,
    "group": "非488名單(另色標記)",
    "subGroup": "非488名單",
    "region": "非488店家",
    "color": "黃色",
    "colorHex": "#f59e0b"
  },
  {
    "id": "store-39",
    "name": "特欣行開發股份有限公司",
    "address": "花蓮縣吉安鄉建國路1段137之2號",
    "lat": 23.985742,
    "lng": 121.591431,
    "estimated": false,
    "isPrimary488": true,
    "group": "主要488名單",
    "delta": -0.018627384615385267,
    "subGroup": "主要名單-西區",
    "region": "界線以西",
    "color": "紅色",
    "colorHex": "#e53e3e"
  },
  {
    "id": "store-40",
    "name": "富利餐飲必勝客吉安店",
    "address": "花蓮縣吉安鄉建國路1段138號",
    "lat": 23.985969,
    "lng": 121.591803,
    "estimated": false,
    "isPrimary488": false,
    "group": "非488名單(另色標記)",
    "subGroup": "非488名單",
    "region": "非488店家",
    "color": "黃色",
    "colorHex": "#f59e0b"
  },
  {
    "id": "store-41",
    "name": "小天下電腦",
    "address": "花蓮縣吉安鄉建國路1段144號",
    "lat": 23.986185,
    "lng": 121.5911,
    "estimated": false,
    "isPrimary488": false,
    "group": "非488名單(另色標記)",
    "subGroup": "非488名單",
    "region": "非488店家",
    "color": "黃色",
    "colorHex": "#f59e0b"
  },
  {
    "id": "store-42",
    "name": "銓來旺彩券行",
    "address": "花蓮縣吉安鄉建國路1段149號",
    "lat": 23.985947,
    "lng": 121.590843,
    "estimated": false,
    "isPrimary488": false,
    "group": "非488名單(另色標記)",
    "subGroup": "非488名單",
    "region": "非488店家",
    "color": "黃色",
    "colorHex": "#f59e0b"
  },
  {
    "id": "store-43",
    "name": "隆昌生活百貨建國店",
    "address": "花蓮縣吉安鄉建國路1段191號",
    "lat": 23.986465,
    "lng": 121.589296,
    "estimated": false,
    "isPrimary488": true,
    "group": "主要488名單",
    "delta": -0.021568807692304404,
    "subGroup": "主要名單-西區",
    "region": "界線以西",
    "color": "紅色",
    "colorHex": "#e53e3e"
  },
  {
    "id": "store-44",
    "name": "順益汽車花蓮營業所",
    "address": "花蓮縣吉安鄉建國路1段197號",
    "lat": 23.986489,
    "lng": 121.588966,
    "estimated": false,
    "isPrimary488": true,
    "group": "主要488名單",
    "delta": -0.02192557692306707,
    "subGroup": "主要名單-西區",
    "region": "界線以西",
    "color": "紅色",
    "colorHex": "#e53e3e"
  },
  {
    "id": "store-45",
    "name": "黃世傑中醫診所",
    "address": "花蓮縣吉安鄉建國路1段61號",
    "lat": 23.985,
    "lng": 121.594541,
    "estimated": false,
    "isPrimary488": true,
    "group": "主要488名單",
    "delta": -0.014689769230756156,
    "subGroup": "主要名單-西區",
    "region": "界線以西",
    "color": "紅色",
    "colorHex": "#e53e3e"
  },
  {
    "id": "store-46",
    "name": "余登貴-2樓住家",
    "address": "花蓮縣吉安鄉建國路1段65號",
    "lat": 23.985048,
    "lng": 121.59454,
    "estimated": false,
    "isPrimary488": true,
    "group": "主要488名單",
    "delta": -0.014744307692311054,
    "subGroup": "主要名單-西區",
    "region": "界線以西",
    "color": "紅色",
    "colorHex": "#e53e3e"
  },
  {
    "id": "store-47",
    "name": "余登貴-藥局",
    "address": "花蓮縣吉安鄉建國路1段65號",
    "lat": 23.985048,
    "lng": 121.59454,
    "estimated": false,
    "isPrimary488": true,
    "group": "主要488名單",
    "delta": -0.014744307692311054,
    "subGroup": "主要名單-西區",
    "region": "界線以西",
    "color": "紅色",
    "colorHex": "#e53e3e"
  },
  {
    "id": "store-48",
    "name": "余登貴-3、4樓住家",
    "address": "花蓮縣吉安鄉建國路1段65號",
    "lat": 23.985048,
    "lng": 121.59454,
    "estimated": false,
    "isPrimary488": true,
    "group": "主要488名單",
    "delta": -0.014744307692311054,
    "subGroup": "主要名單-西區",
    "region": "界線以西",
    "color": "紅色",
    "colorHex": "#e53e3e"
  },
  {
    "id": "store-49",
    "name": "余登貴-診所",
    "address": "花蓮縣吉安鄉建國路1段65號",
    "lat": 23.985048,
    "lng": 121.59454,
    "estimated": false,
    "isPrimary488": true,
    "group": "主要488名單",
    "delta": -0.014744307692311054,
    "subGroup": "主要名單-西區",
    "region": "界線以西",
    "color": "紅色",
    "colorHex": "#e53e3e"
  },
  {
    "id": "store-50",
    "name": "宋威揚營業所建國店",
    "address": "花蓮縣吉安鄉建國路368號",
    "lat": 23.984819,
    "lng": 121.597069,
    "estimated": false,
    "isPrimary488": true,
    "group": "主要488名單",
    "delta": -0.011959884615379224,
    "subGroup": "主要名單-西區",
    "region": "界線以西",
    "color": "紅色",
    "colorHex": "#e53e3e"
  },
  {
    "id": "store-51",
    "name": "台灣真工國際股份有限公司",
    "address": "花蓮縣吉安鄉慈惠3街114號",
    "lat": 23.977408,
    "lng": 121.587133,
    "estimated": false,
    "isPrimary488": true,
    "group": "主要488名單",
    "delta": -0.013629769230774968,
    "subGroup": "主要名單-西區",
    "region": "界線以西",
    "color": "紅色",
    "colorHex": "#e53e3e"
  },
  {
    "id": "store-52",
    "name": "花悠家美食館",
    "address": "花蓮縣吉安鄉慈惠一街2號",
    "lat": 23.974357,
    "lng": 121.588348,
    "estimated": false,
    "isPrimary488": true,
    "group": "主要488名單",
    "delta": -0.009011730769231008,
    "subGroup": "主要名單-西區",
    "region": "界線以西",
    "color": "紅色",
    "colorHex": "#e53e3e"
  },
  {
    "id": "store-53",
    "name": "龔慧萍公館",
    "address": "花蓮縣吉安鄉慈惠二街43號",
    "lat": 23.9757,
    "lng": 121.586694,
    "estimated": false,
    "isPrimary488": true,
    "group": "主要488名單",
    "delta": -0.012163692307694873,
    "subGroup": "主要名單-西區",
    "region": "界線以西",
    "color": "紅色",
    "colorHex": "#e53e3e"
  },
  {
    "id": "store-54",
    "name": "花蓮一信ATM慈惠三街",
    "address": "花蓮縣吉安鄉慈惠三街188號",
    "lat": 23.9774,
    "lng": 121.585426,
    "estimated": false,
    "isPrimary488": true,
    "group": "主要488名單",
    "delta": -0.015327846153851965,
    "subGroup": "主要名單-西區",
    "region": "界線以西",
    "color": "紅色",
    "colorHex": "#e53e3e"
  },
  {
    "id": "store-55",
    "name": "葉善美公館",
    "address": "花蓮縣吉安鄉慈惠四街85巷2號",
    "lat": 23.977707,
    "lng": 121.58808,
    "estimated": false,
    "isPrimary488": true,
    "group": "主要488名單",
    "delta": -0.013016269230760713,
    "subGroup": "主要名單-西區",
    "region": "界線以西",
    "color": "紅色",
    "colorHex": "#e53e3e"
  },
  {
    "id": "store-56",
    "name": "台灣銀行花蓮分行倉庫",
    "address": "花蓮縣花蓮市三民街1號",
    "lat": 23.977502,
    "lng": 121.610994,
    "estimated": false,
    "isPrimary488": true,
    "group": "主要488名單",
    "delta": 0.010126384615389838,
    "subGroup": "主要名單-東區",
    "region": "界線以東",
    "color": "綠色",
    "colorHex": "#257a4e"
  },
  {
    "id": "store-57",
    "name": "普客二四花蓮三民街",
    "address": "花蓮縣花蓮市三民街1號",
    "lat": 23.977502,
    "lng": 121.610994,
    "estimated": false,
    "isPrimary488": false,
    "group": "非488名單(另色標記)",
    "subGroup": "非488名單",
    "region": "非488店家",
    "color": "黃色",
    "colorHex": "#f59e0b"
  },
  {
    "id": "store-58",
    "name": "花蓮縣縣定古蹟檢察長宿舍",
    "address": "花蓮縣花蓮市三民街23號",
    "lat": 23.979902,
    "lng": 121.608916,
    "estimated": false,
    "isPrimary488": true,
    "group": "主要488名單",
    "delta": 0.0053714615384592435,
    "subGroup": "主要名單-東區",
    "region": "界線以東",
    "color": "綠色",
    "colorHex": "#257a4e"
  },
  {
    "id": "store-59",
    "name": "花蓮縣私立佳姿幼兒園",
    "address": "花蓮縣花蓮市三民街82-6號",
    "lat": 23.980689,
    "lng": 121.608693,
    "estimated": false,
    "isPrimary488": true,
    "group": "主要488名單",
    "delta": 0.004270653846148775,
    "subGroup": "主要名單-東區",
    "region": "界線以東",
    "color": "綠色",
    "colorHex": "#257a4e"
  },
  {
    "id": "store-60",
    "name": "志成電子材料行",
    "address": "花蓮縣花蓮市三民街84之1號",
    "lat": 23.981589,
    "lng": 121.607924,
    "estimated": false,
    "isPrimary488": true,
    "group": "主要488名單",
    "delta": 0.002497807692307674,
    "subGroup": "主要名單-東區",
    "region": "界線以東",
    "color": "綠色",
    "colorHex": "#257a4e"
  },
  {
    "id": "store-61",
    "name": "承天佛具",
    "address": "花蓮縣花蓮市三民街88號",
    "lat": 23.981947,
    "lng": 121.60772,
    "estimated": false,
    "isPrimary488": true,
    "group": "主要488名單",
    "delta": 0.0018945000000059053,
    "subGroup": "主要名單-東區",
    "region": "界線以東",
    "color": "綠色",
    "colorHex": "#257a4e"
  },
  {
    "id": "store-62",
    "name": "鄒永強公館",
    "address": "花蓮縣花蓮市大同街110號",
    "lat": 23.982316,
    "lng": 121.608962,
    "estimated": false,
    "isPrimary488": true,
    "group": "主要488名單",
    "delta": 0.0027249230769257338,
    "subGroup": "主要名單-東區",
    "region": "界線以東",
    "color": "綠色",
    "colorHex": "#257a4e"
  },
  {
    "id": "store-63",
    "name": "郭彩玉洗衣店",
    "address": "花蓮縣花蓮市大同街31號",
    "lat": 23.9803,
    "lng": 121.605663,
    "estimated": false,
    "isPrimary488": true,
    "group": "主要488名單",
    "delta": 0.0016745384615433068,
    "subGroup": "主要名單-東區",
    "region": "界線以東",
    "color": "綠色",
    "colorHex": "#257a4e"
  },
  {
    "id": "store-64",
    "name": "信香行超市",
    "address": "花蓮縣花蓮市大同街33號B1",
    "lat": 23.980514,
    "lng": 121.605964,
    "estimated": false,
    "isPrimary488": true,
    "group": "主要488名單",
    "delta": 0.0017368461538467272,
    "subGroup": "主要名單-東區",
    "region": "界線以東",
    "color": "綠色",
    "colorHex": "#257a4e"
  },
  {
    "id": "store-65",
    "name": "利東工礦股份有限公司",
    "address": "花蓮縣花蓮市大同街5號1-3樓",
    "lat": 23.979416,
    "lng": 121.60432,
    "estimated": false,
    "isPrimary488": true,
    "group": "主要488名單",
    "delta": 0.0013175384615351504,
    "subGroup": "主要名單-東區",
    "region": "界線以東",
    "color": "綠色",
    "colorHex": "#257a4e"
  },
  {
    "id": "store-66",
    "name": "竇允辰公館",
    "address": "花蓮縣花蓮市大同街98巷5-1號",
    "lat": 23.985248,
    "lng": 121.581054,
    "estimated": false,
    "isPrimary488": true,
    "group": "主要488名單",
    "delta": -0.028453384615389155,
    "subGroup": "主要名單-西區",
    "region": "界線以西",
    "color": "紅色",
    "colorHex": "#e53e3e"
  },
  {
    "id": "store-67",
    "name": "吳明株營業所",
    "address": "花蓮縣花蓮市大禹街50號1.2樓",
    "lat": 23.977107,
    "lng": 121.608305,
    "estimated": false,
    "isPrimary488": false,
    "group": "非488名單(另色標記)",
    "subGroup": "非488名單",
    "region": "非488店家",
    "color": "黃色",
    "colorHex": "#f59e0b"
  },
  {
    "id": "store-68",
    "name": "吳明株營業所",
    "address": "花蓮縣花蓮市大禹街64號",
    "lat": 23.977215,
    "lng": 121.608141,
    "estimated": false,
    "isPrimary488": true,
    "group": "主要488名單",
    "delta": 0.0075934999999986985,
    "subGroup": "主要名單-東區",
    "region": "界線以東",
    "color": "綠色",
    "colorHex": "#257a4e"
  },
  {
    "id": "store-69",
    "name": "花蓮汽車客運股份有限公司",
    "address": "花蓮縣花蓮市中山路137號",
    "lat": 23.975647,
    "lng": 121.610354,
    "estimated": false,
    "isPrimary488": true,
    "group": "主要488名單",
    "delta": 0.011555423076927696,
    "subGroup": "主要名單-東區",
    "region": "界線以東",
    "color": "綠色",
    "colorHex": "#257a4e"
  },
  {
    "id": "store-70",
    "name": "日興珠寶有限公司",
    "address": "花蓮縣花蓮市中山路142-15號",
    "lat": 23.97652,
    "lng": 121.60996,
    "estimated": false,
    "isPrimary488": true,
    "group": "主要488名單",
    "delta": 0.010187692307695784,
    "subGroup": "主要名單-東區",
    "region": "界線以東",
    "color": "綠色",
    "colorHex": "#257a4e"
  },
  {
    "id": "store-71",
    "name": "曾記臻隆達有限公司",
    "address": "花蓮縣花蓮市中山路142號",
    "lat": 23.97652,
    "lng": 121.60996,
    "estimated": false,
    "isPrimary488": true,
    "group": "主要488名單",
    "delta": 0.010187692307695784,
    "subGroup": "主要名單-東區",
    "region": "界線以東",
    "color": "綠色",
    "colorHex": "#257a4e"
  },
  {
    "id": "store-72",
    "name": "台灣大哥大花蓮舊站直營店",
    "address": "花蓮縣花蓮市中山路157號",
    "lat": 23.976011,
    "lng": 121.610055,
    "estimated": false,
    "isPrimary488": true,
    "group": "主要488名單",
    "delta": 0.010850423076931293,
    "subGroup": "主要名單-東區",
    "region": "界線以東",
    "color": "綠色",
    "colorHex": "#257a4e"
  },
  {
    "id": "store-73",
    "name": "釗聖體育用品有限公司花蓮分公司",
    "address": "花蓮縣花蓮市中山路171、173、175號",
    "lat": 23.976459,
    "lng": 121.609711,
    "estimated": false,
    "isPrimary488": false,
    "group": "非488名單(另色標記)",
    "subGroup": "非488名單",
    "region": "非488店家",
    "color": "黃色",
    "colorHex": "#f59e0b"
  },
  {
    "id": "store-74",
    "name": "巨匠電腦花蓮認證",
    "address": "花蓮縣花蓮市中山路180號",
    "lat": 23.97795,
    "lng": 121.60879,
    "estimated": false,
    "isPrimary488": true,
    "group": "主要488名單",
    "delta": 0.007422692307684997,
    "subGroup": "主要名單-東區",
    "region": "界線以東",
    "color": "綠色",
    "colorHex": "#257a4e"
  },
  {
    "id": "store-75",
    "name": "大樹醫藥花蓮中山直營店",
    "address": "花蓮縣花蓮市中山路192號",
    "lat": 23.97859,
    "lng": 121.608287,
    "estimated": false,
    "isPrimary488": false,
    "group": "非488名單(另色標記)",
    "subGroup": "非488名單",
    "region": "非488店家",
    "color": "黃色",
    "colorHex": "#f59e0b"
  },
  {
    "id": "store-76",
    "name": "金六福珠寶銀樓",
    "address": "花蓮縣花蓮市中山路199-1號",
    "lat": 23.977124,
    "lng": 121.6092,
    "estimated": false,
    "isPrimary488": true,
    "group": "主要488名單",
    "delta": 0.008753999999996154,
    "subGroup": "主要名單-東區",
    "region": "界線以東",
    "color": "綠色",
    "colorHex": "#257a4e"
  },
  {
    "id": "store-77",
    "name": "科訊工坊花蓮店",
    "address": "花蓮縣花蓮市中山路207號",
    "lat": 23.977411,
    "lng": 121.608946,
    "estimated": false,
    "isPrimary488": false,
    "group": "非488名單(另色標記)",
    "subGroup": "非488名單",
    "region": "非488店家",
    "color": "黃色",
    "colorHex": "#f59e0b"
  },
  {
    "id": "store-78",
    "name": "Snapshot花蓮店",
    "address": "花蓮縣花蓮市中山路207號",
    "lat": 23.977411,
    "lng": 121.608946,
    "estimated": false,
    "isPrimary488": false,
    "group": "非488名單(另色標記)",
    "subGroup": "非488名單",
    "region": "非488店家",
    "color": "黃色",
    "colorHex": "#f59e0b"
  },
  {
    "id": "store-79",
    "name": "悠旅生活花蓮門市",
    "address": "花蓮縣花蓮市中山路231號1-2樓",
    "lat": 23.978093,
    "lng": 121.608365,
    "estimated": false,
    "isPrimary488": true,
    "group": "主要488名單",
    "delta": 0.006838192307697,
    "subGroup": "主要名單-東區",
    "region": "界線以東",
    "color": "綠色",
    "colorHex": "#257a4e"
  },
  {
    "id": "store-80",
    "name": "台灣大哥大花蓮民國直營店",
    "address": "花蓮縣花蓮市中山路240號",
    "lat": 23.979981,
    "lng": 121.607196,
    "estimated": false,
    "isPrimary488": true,
    "group": "主要488名單",
    "delta": 0.003563346153853786,
    "subGroup": "主要名單-東區",
    "region": "界線以東",
    "color": "綠色",
    "colorHex": "#257a4e"
  },
  {
    "id": "store-81",
    "name": "陳宏瑋營業所",
    "address": "花蓮縣花蓮市中山路253號",
    "lat": 23.978482,
    "lng": 121.608079,
    "estimated": false,
    "isPrimary488": false,
    "group": "非488名單(另色標記)",
    "subGroup": "非488名單",
    "region": "非488店家",
    "color": "黃色",
    "colorHex": "#f59e0b"
  },
  {
    "id": "store-82",
    "name": "科林國際助聽器花蓮中山店",
    "address": "花蓮縣花蓮市中山路276號1樓",
    "lat": 23.980704,
    "lng": 121.606701,
    "estimated": false,
    "isPrimary488": true,
    "group": "主要488名單",
    "delta": 0.0022619230769294063,
    "subGroup": "主要名單-東區",
    "region": "界線以東",
    "color": "綠色",
    "colorHex": "#257a4e"
  },
  {
    "id": "store-83",
    "name": "麵匡匡拉麵食堂花蓮中山店",
    "address": "花蓮縣花蓮市中山路280號1樓",
    "lat": 23.98081,
    "lng": 121.6067,
    "estimated": false,
    "isPrimary488": true,
    "group": "主要488名單",
    "delta": 0.0021426923077001447,
    "subGroup": "主要名單-東區",
    "region": "界線以東",
    "color": "綠色",
    "colorHex": "#257a4e"
  },
  {
    "id": "store-84",
    "name": "俞碧玲營業所",
    "address": "花蓮縣花蓮市中山路301號",
    "lat": 23.979529,
    "lng": 121.607265,
    "estimated": false,
    "isPrimary488": false,
    "group": "非488名單(另色標記)",
    "subGroup": "非488名單",
    "region": "非488店家",
    "color": "黃色",
    "colorHex": "#f59e0b"
  },
  {
    "id": "store-85",
    "name": "全球通訊行",
    "address": "花蓮縣花蓮市中山路311號",
    "lat": 23.97988,
    "lng": 121.60688,
    "estimated": false,
    "isPrimary488": true,
    "group": "主要488名單",
    "delta": 0.0033600000000006958,
    "subGroup": "主要名單-東區",
    "region": "界線以東",
    "color": "綠色",
    "colorHex": "#257a4e"
  },
  {
    "id": "store-86",
    "name": "科林國際助聽器花蓮中山門市",
    "address": "花蓮縣花蓮市中山路315號",
    "lat": 23.979965,
    "lng": 121.606937,
    "estimated": false,
    "isPrimary488": false,
    "group": "非488名單(另色標記)",
    "subGroup": "非488名單",
    "region": "非488店家",
    "color": "黃色",
    "colorHex": "#f59e0b"
  },
  {
    "id": "store-87",
    "name": "仁和堂藥房",
    "address": "花蓮縣花蓮市中山路316號",
    "lat": 23.981487,
    "lng": 121.606055,
    "estimated": false,
    "isPrimary488": true,
    "group": "主要488名單",
    "delta": 0.0007425769230735568,
    "subGroup": "主要名單-東區",
    "region": "界線以東",
    "color": "綠色",
    "colorHex": "#257a4e"
  },
  {
    "id": "store-88",
    "name": "順發電腦花蓮店",
    "address": "花蓮縣花蓮市中山路321號",
    "lat": 23.973521,
    "lng": 121.612302,
    "estimated": false,
    "isPrimary488": true,
    "group": "主要488名單",
    "delta": 0.01587473076922663,
    "subGroup": "主要名單-東區",
    "region": "界線以東",
    "color": "綠色",
    "colorHex": "#257a4e"
  },
  {
    "id": "store-89",
    "name": "佑全藥品1701花蓮中山店",
    "address": "花蓮縣花蓮市中山路326號",
    "lat": 23.981772,
    "lng": 121.605787,
    "estimated": false,
    "isPrimary488": true,
    "group": "主要488名單",
    "delta": 0.000156692307697881,
    "subGroup": "主要名單-東區",
    "region": "界線以東",
    "color": "綠色",
    "colorHex": "#257a4e"
  },
  {
    "id": "store-90",
    "name": "花蓮二信中山分社",
    "address": "花蓮縣花蓮市中山路344號",
    "lat": 23.982329,
    "lng": 121.605455,
    "estimated": false,
    "isPrimary488": true,
    "group": "主要488名單",
    "delta": -0.0007965769230651176,
    "subGroup": "主要名單-西區",
    "region": "界線以西",
    "color": "紅色",
    "colorHex": "#e53e3e"
  },
  {
    "id": "store-91",
    "name": "南山人壽花蓮分公司",
    "address": "花蓮縣花蓮市中山路356號4樓",
    "lat": 23.982497,
    "lng": 121.60529,
    "estimated": false,
    "isPrimary488": true,
    "group": "主要488名單",
    "delta": -0.00114896153846189,
    "subGroup": "主要名單-西區",
    "region": "界線以西",
    "color": "紅色",
    "colorHex": "#e53e3e"
  },
  {
    "id": "store-92",
    "name": "百悅休閒飯店有限公司",
    "address": "花蓮縣花蓮市中山路361號",
    "lat": 23.981226,
    "lng": 121.605926,
    "estimated": false,
    "isPrimary488": false,
    "group": "非488名單(另色標記)",
    "subGroup": "非488名單",
    "region": "非488店家",
    "color": "黃色",
    "colorHex": "#f59e0b"
  },
  {
    "id": "store-93",
    "name": "合作金庫北花蓮分行",
    "address": "花蓮縣花蓮市中山路371號1-2樓",
    "lat": 23.981376,
    "lng": 121.605797,
    "estimated": false,
    "isPrimary488": true,
    "group": "主要488名單",
    "delta": 0.0006083846153757122,
    "subGroup": "主要名單-東區",
    "region": "界線以東",
    "color": "綠色",
    "colorHex": "#257a4e"
  },
  {
    "id": "store-94",
    "name": "內政部入出國及移民署北區事務大",
    "address": "花蓮縣花蓮市中山路371號5樓",
    "lat": 23.981376,
    "lng": 121.605797,
    "estimated": false,
    "isPrimary488": true,
    "group": "主要488名單",
    "delta": 0.0006083846153757122,
    "subGroup": "主要名單-東區",
    "region": "界線以東",
    "color": "綠色",
    "colorHex": "#257a4e"
  },
  {
    "id": "store-95",
    "name": "外交部領事事務局花蓮辦事處",
    "address": "花蓮縣花蓮市中山路371號6樓",
    "lat": 23.981376,
    "lng": 121.605797,
    "estimated": false,
    "isPrimary488": true,
    "group": "主要488名單",
    "delta": 0.0006083846153757122,
    "subGroup": "主要名單-東區",
    "region": "界線以東",
    "color": "綠色",
    "colorHex": "#257a4e"
  },
  {
    "id": "store-96",
    "name": "行政院東部聯合服務中心",
    "address": "花蓮縣花蓮市中山路371號7-8樓",
    "lat": 23.981376,
    "lng": 121.605797,
    "estimated": false,
    "isPrimary488": true,
    "group": "主要488名單",
    "delta": 0.0006083846153757122,
    "subGroup": "主要名單-東區",
    "region": "界線以東",
    "color": "綠色",
    "colorHex": "#257a4e"
  },
  {
    "id": "store-97",
    "name": "張慧蘭營業所WeMe",
    "address": "花蓮縣花蓮市中山路372號",
    "lat": 23.982862,
    "lng": 121.604955,
    "estimated": false,
    "isPrimary488": false,
    "group": "非488名單(另色標記)",
    "subGroup": "非488名單",
    "region": "非488店家",
    "color": "黃色",
    "colorHex": "#f59e0b"
  },
  {
    "id": "store-98",
    "name": "爭鮮外帶花蓮中山店",
    "address": "花蓮縣花蓮市中山路373號",
    "lat": 23.981666,
    "lng": 121.605628,
    "estimated": false,
    "isPrimary488": true,
    "group": "主要488名單",
    "delta": 0.00011592307691898895,
    "subGroup": "主要名單-東區",
    "region": "界線以東",
    "color": "綠色",
    "colorHex": "#257a4e"
  },
  {
    "id": "store-99",
    "name": "焦點沖印行",
    "address": "花蓮縣花蓮市中山路374號",
    "lat": 23.982891,
    "lng": 121.604931,
    "estimated": false,
    "isPrimary488": true,
    "group": "主要488名單",
    "delta": -0.0019474230769276346,
    "subGroup": "主要名單-西區",
    "region": "界線以西",
    "color": "紅色",
    "colorHex": "#e53e3e"
  },
  {
    "id": "store-100",
    "name": "聯爵廣告代銷中山路383號",
    "address": "花蓮縣花蓮市中山路383號",
    "lat": 23.981866,
    "lng": 121.605464,
    "estimated": false,
    "isPrimary488": false,
    "group": "非488名單(另色標記)",
    "subGroup": "非488名單",
    "region": "非488店家",
    "color": "黃色",
    "colorHex": "#f59e0b"
  },
  {
    "id": "store-101",
    "name": "顏暐力營業所",
    "address": "花蓮縣花蓮市中山路406之2號",
    "lat": 23.972609,
    "lng": 121.612331,
    "estimated": false,
    "isPrimary488": true,
    "group": "主要488名單",
    "delta": 0.016920961538460233,
    "subGroup": "主要名單-東區",
    "region": "界線以東",
    "color": "綠色",
    "colorHex": "#257a4e"
  },
  {
    "id": "store-102",
    "name": "全球當鋪",
    "address": "花蓮縣花蓮市中山路408-7號",
    "lat": 23.98402,
    "lng": 121.604438,
    "estimated": false,
    "isPrimary488": false,
    "group": "非488名單(另色標記)",
    "subGroup": "非488名單",
    "region": "非488店家",
    "color": "黃色",
    "colorHex": "#f59e0b"
  },
  {
    "id": "store-103",
    "name": "台新銀行花蓮分行",
    "address": "花蓮縣花蓮市中山路408-9號",
    "lat": 23.98402,
    "lng": 121.604438,
    "estimated": false,
    "isPrimary488": true,
    "group": "主要488名單",
    "delta": -0.0036996923076912935,
    "subGroup": "主要名單-西區",
    "region": "界線以西",
    "color": "紅色",
    "colorHex": "#e53e3e"
  },
  {
    "id": "store-104",
    "name": "台新銀行ATM花蓮分行",
    "address": "花蓮縣花蓮市中山路408-9號",
    "lat": 23.98402,
    "lng": 121.604438,
    "estimated": false,
    "isPrimary488": true,
    "group": "主要488名單",
    "delta": -0.0036996923076912935,
    "subGroup": "主要名單-西區",
    "region": "界線以西",
    "color": "紅色",
    "colorHex": "#e53e3e"
  },
  {
    "id": "store-105",
    "name": "台新銀行花蓮分行金庫",
    "address": "花蓮縣花蓮市中山路408-9號",
    "lat": 23.98402,
    "lng": 121.604438,
    "estimated": false,
    "isPrimary488": true,
    "group": "主要488名單",
    "delta": -0.0036996923076912935,
    "subGroup": "主要名單-西區",
    "region": "界線以西",
    "color": "紅色",
    "colorHex": "#e53e3e"
  },
  {
    "id": "store-106",
    "name": "金壽砂石股份有限公司",
    "address": "花蓮縣花蓮市中山路412巷12號1樓",
    "lat": 23.985433,
    "lng": 121.603633,
    "estimated": false,
    "isPrimary488": true,
    "group": "主要488名單",
    "delta": -0.006080730769227216,
    "subGroup": "主要名單-西區",
    "region": "界線以西",
    "color": "紅色",
    "colorHex": "#e53e3e"
  },
  {
    "id": "store-107",
    "name": "姚定國公館",
    "address": "花蓮縣花蓮市中山路412巷2-2號",
    "lat": 23.972609,
    "lng": 121.612331,
    "estimated": false,
    "isPrimary488": true,
    "group": "主要488名單",
    "delta": 0.016920961538460233,
    "subGroup": "主要名單-東區",
    "region": "界線以東",
    "color": "綠色",
    "colorHex": "#257a4e"
  },
  {
    "id": "store-108",
    "name": "民豐車業行",
    "address": "花蓮縣花蓮市中山路413-1號",
    "lat": 23.982481,
    "lng": 121.604964,
    "estimated": false,
    "isPrimary488": true,
    "group": "主要488名單",
    "delta": -0.0014571153846247853,
    "subGroup": "主要名單-西區",
    "region": "界線以西",
    "color": "紅色",
    "colorHex": "#e53e3e"
  },
  {
    "id": "store-109",
    "name": "花蓮二信ATM花蓮商校",
    "address": "花蓮縣花蓮市中山路418號",
    "lat": 23.986369,
    "lng": 121.602348,
    "estimated": false,
    "isPrimary488": true,
    "group": "主要488名單",
    "delta": -0.00840973076921614,
    "subGroup": "主要名單-西區",
    "region": "界線以西",
    "color": "紅色",
    "colorHex": "#e53e3e"
  },
  {
    "id": "store-110",
    "name": "花商學生宿舍",
    "address": "花蓮縣花蓮市中山路418號",
    "lat": 23.986369,
    "lng": 121.602348,
    "estimated": false,
    "isPrimary488": true,
    "group": "主要488名單",
    "delta": -0.00840973076921614,
    "subGroup": "主要名單-西區",
    "region": "界線以西",
    "color": "紅色",
    "colorHex": "#e53e3e"
  },
  {
    "id": "store-111",
    "name": "金湯達人咖啡館",
    "address": "花蓮縣花蓮市中山路431號",
    "lat": 23.982809,
    "lng": 121.604619,
    "estimated": false,
    "isPrimary488": true,
    "group": "主要488名單",
    "delta": -0.002167961538461327,
    "subGroup": "主要名單-西區",
    "region": "界線以西",
    "color": "紅色",
    "colorHex": "#e53e3e"
  },
  {
    "id": "store-112",
    "name": "許庭崧-串串門子鴛鴦火鍋",
    "address": "花蓮縣花蓮市中山路434號",
    "lat": 23.988086,
    "lng": 121.601003,
    "estimated": false,
    "isPrimary488": false,
    "group": "非488名單(另色標記)",
    "subGroup": "非488名單",
    "region": "非488店家",
    "color": "黃色",
    "colorHex": "#f59e0b"
  },
  {
    "id": "store-113",
    "name": "宋威揚營業所中山店",
    "address": "花蓮縣花蓮市中山路472號",
    "lat": 23.988749,
    "lng": 121.600475,
    "estimated": false,
    "isPrimary488": true,
    "group": "主要488名單",
    "delta": -0.012937346153833573,
    "subGroup": "主要名單-西區",
    "region": "界線以西",
    "color": "紅色",
    "colorHex": "#e53e3e"
  },
  {
    "id": "store-114",
    "name": "花蓮一信ATM總社後側",
    "address": "花蓮縣花蓮市中山路473號",
    "lat": 23.983829,
    "lng": 121.603893,
    "estimated": false,
    "isPrimary488": true,
    "group": "主要488名單",
    "delta": -0.004031653846155336,
    "subGroup": "主要名單-西區",
    "region": "界線以西",
    "color": "紅色",
    "colorHex": "#e53e3e"
  },
  {
    "id": "store-115",
    "name": "花蓮一信總社",
    "address": "花蓮縣花蓮市中山路473號",
    "lat": 23.983829,
    "lng": 121.603893,
    "estimated": false,
    "isPrimary488": true,
    "group": "主要488名單",
    "delta": -0.004031653846155336,
    "subGroup": "主要名單-西區",
    "region": "界線以西",
    "color": "紅色",
    "colorHex": "#e53e3e"
  },
  {
    "id": "store-116",
    "name": "花蓮一信保管箱",
    "address": "花蓮縣花蓮市中山路473號2樓",
    "lat": 23.983829,
    "lng": 121.603893,
    "estimated": false,
    "isPrimary488": true,
    "group": "主要488名單",
    "delta": -0.004031653846155336,
    "subGroup": "主要名單-西區",
    "region": "界線以西",
    "color": "紅色",
    "colorHex": "#e53e3e"
  },
  {
    "id": "store-117",
    "name": "花蓮一信總社",
    "address": "花蓮縣花蓮市中山路473號3.5.7樓",
    "lat": 23.983829,
    "lng": 121.603893,
    "estimated": false,
    "isPrimary488": true,
    "group": "主要488名單",
    "delta": -0.004031653846155336,
    "subGroup": "主要名單-西區",
    "region": "界線以西",
    "color": "紅色",
    "colorHex": "#e53e3e"
  },
  {
    "id": "store-118",
    "name": "花蓮一信總社資訊室",
    "address": "花蓮縣花蓮市中山路473號4樓",
    "lat": 23.983829,
    "lng": 121.603893,
    "estimated": false,
    "isPrimary488": true,
    "group": "主要488名單",
    "delta": -0.004031653846155336,
    "subGroup": "主要名單-西區",
    "region": "界線以西",
    "color": "紅色",
    "colorHex": "#e53e3e"
  },
  {
    "id": "store-119",
    "name": "普客二四花蓮延平街",
    "address": "花蓮縣花蓮市中山路475-9號旁",
    "lat": 23.98419,
    "lng": 121.60351,
    "estimated": true,
    "isPrimary488": true,
    "group": "主要488名單",
    "delta": -0.004817307692306372,
    "subGroup": "主要名單-西區",
    "region": "界線以西",
    "color": "紅色",
    "colorHex": "#e53e3e"
  },
  {
    "id": "store-120",
    "name": "翹宇宙有限公司中山店",
    "address": "花蓮縣花蓮市中山路48.50號",
    "lat": 23.974133,
    "lng": 121.611771,
    "estimated": false,
    "isPrimary488": false,
    "group": "非488名單(另色標記)",
    "subGroup": "非488名單",
    "region": "非488店家",
    "color": "黃色",
    "colorHex": "#f59e0b"
  },
  {
    "id": "store-121",
    "name": "鍾思瑀營業所",
    "address": "花蓮縣花蓮市中山路50號(E73、E74攤位）",
    "lat": 23.97419,
    "lng": 121.611724,
    "estimated": true,
    "isPrimary488": false,
    "group": "非488名單(另色標記)",
    "subGroup": "非488名單",
    "region": "非488店家",
    "color": "黃色",
    "colorHex": "#f59e0b"
  },
  {
    "id": "store-122",
    "name": "政大書城公司花蓮分公司",
    "address": "花蓮縣花蓮市中山路547-2號3樓",
    "lat": 23.986502,
    "lng": 121.601811,
    "estimated": false,
    "isPrimary488": false,
    "group": "非488名單(另色標記)",
    "subGroup": "非488名單",
    "region": "非488店家",
    "color": "黃色",
    "colorHex": "#f59e0b"
  },
  {
    "id": "store-123",
    "name": "花蓮市農會購物中心",
    "address": "花蓮縣花蓮市中山路547號1樓",
    "lat": 23.986502,
    "lng": 121.601811,
    "estimated": false,
    "isPrimary488": true,
    "group": "主要488名單",
    "delta": -0.00909507692307443,
    "subGroup": "主要名單-西區",
    "region": "界線以西",
    "color": "紅色",
    "colorHex": "#e53e3e"
  },
  {
    "id": "store-124",
    "name": "安心食品花蓮中山店",
    "address": "花蓮縣花蓮市中山路549-1號",
    "lat": 23.986823,
    "lng": 121.6015,
    "estimated": false,
    "isPrimary488": true,
    "group": "主要488名單",
    "delta": -0.009764115384612637,
    "subGroup": "主要名單-西區",
    "region": "界線以西",
    "color": "紅色",
    "colorHex": "#e53e3e"
  },
  {
    "id": "store-125",
    "name": "台灣麥當勞和德昌花蓮中山店",
    "address": "花蓮縣花蓮市中山路549號",
    "lat": 23.986823,
    "lng": 121.6015,
    "estimated": false,
    "isPrimary488": true,
    "group": "主要488名單",
    "delta": -0.009764115384612637,
    "subGroup": "主要名單-西區",
    "region": "界線以西",
    "color": "紅色",
    "colorHex": "#e53e3e"
  },
  {
    "id": "store-126",
    "name": "義美食品花蓮店",
    "address": "花蓮縣花蓮市中山路551-2號",
    "lat": 23.987246,
    "lng": 121.601158,
    "estimated": false,
    "isPrimary488": true,
    "group": "主要488名單",
    "delta": -0.010577923076922957,
    "subGroup": "主要名單-西區",
    "region": "界線以西",
    "color": "紅色",
    "colorHex": "#e53e3e"
  },
  {
    "id": "store-127",
    "name": "台新銀行ATM丁丁藥局花蓮店內部",
    "address": "花蓮縣花蓮市中山路553號",
    "lat": 23.987607,
    "lng": 121.600927,
    "estimated": false,
    "isPrimary488": false,
    "group": "非488名單(另色標記)",
    "subGroup": "非488名單",
    "region": "非488店家",
    "color": "黃色",
    "colorHex": "#f59e0b"
  },
  {
    "id": "store-128",
    "name": "全發號陳金堂公館",
    "address": "花蓮縣花蓮市中山路571巷25號1-4樓",
    "lat": 23.98753,
    "lng": 121.600008,
    "estimated": false,
    "isPrimary488": true,
    "group": "主要488名單",
    "delta": -0.01204469230768268,
    "subGroup": "主要名單-西區",
    "region": "界線以西",
    "color": "紅色",
    "colorHex": "#e53e3e"
  },
  {
    "id": "store-129",
    "name": "英倫美學牙醫診所",
    "address": "花蓮縣花蓮市中山路585號",
    "lat": 23.988274,
    "lng": 121.600461,
    "estimated": false,
    "isPrimary488": true,
    "group": "主要488名單",
    "delta": -0.012421538461538262,
    "subGroup": "主要名單-西區",
    "region": "界線以西",
    "color": "紅色",
    "colorHex": "#e53e3e"
  },
  {
    "id": "store-130",
    "name": "和泰產險花蓮通訊處",
    "address": "花蓮縣花蓮市中山路603號2樓",
    "lat": 23.988596,
    "lng": 121.60001,
    "estimated": false,
    "isPrimary488": true,
    "group": "主要488名單",
    "delta": -0.013231692307698495,
    "subGroup": "主要名單-西區",
    "region": "界線以西",
    "color": "紅色",
    "colorHex": "#e53e3e"
  },
  {
    "id": "store-131",
    "name": "神將坊",
    "address": "花蓮縣花蓮市中山路6號",
    "lat": 23.97249,
    "lng": 121.612421,
    "estimated": false,
    "isPrimary488": true,
    "group": "主要488名單",
    "delta": 0.017143692307683978,
    "subGroup": "主要名單-東區",
    "region": "界線以東",
    "color": "綠色",
    "colorHex": "#257a4e"
  },
  {
    "id": "store-132",
    "name": "劉財來營業所",
    "address": "花蓮縣花蓮市中正路10號",
    "lat": 23.978944,
    "lng": 121.610115,
    "estimated": false,
    "isPrimary488": false,
    "group": "非488名單(另色標記)",
    "subGroup": "非488名單",
    "region": "非488店家",
    "color": "黃色",
    "colorHex": "#f59e0b"
  },
  {
    "id": "store-133",
    "name": "大樹醫藥花蓮中正直營店",
    "address": "花蓮縣花蓮市中正路123號",
    "lat": 23.9708,
    "lng": 121.597319,
    "estimated": false,
    "isPrimary488": false,
    "group": "非488名單(另色標記)",
    "subGroup": "非488名單",
    "region": "非488店家",
    "color": "黃色",
    "colorHex": "#f59e0b"
  },
  {
    "id": "store-134",
    "name": "劉品岑公館",
    "address": "花蓮縣花蓮市中正路125巷1號",
    "lat": 23.97099,
    "lng": 121.597262,
    "estimated": false,
    "isPrimary488": true,
    "group": "主要488名單",
    "delta": 0.0036577692307702137,
    "subGroup": "主要名單-東區",
    "region": "界線以東",
    "color": "綠色",
    "colorHex": "#257a4e"
  },
  {
    "id": "store-135",
    "name": "黃淳洋營業所",
    "address": "花蓮縣花蓮市中正路143巷7號",
    "lat": 23.971293,
    "lng": 121.597522,
    "estimated": false,
    "isPrimary488": true,
    "group": "主要488名單",
    "delta": 0.003579807692304371,
    "subGroup": "主要名單-東區",
    "region": "界線以東",
    "color": "綠色",
    "colorHex": "#257a4e"
  },
  {
    "id": "store-136",
    "name": "國泰聯合診所-中正路",
    "address": "花蓮縣花蓮市中正路194號",
    "lat": 23.971416,
    "lng": 121.598754,
    "estimated": false,
    "isPrimary488": true,
    "group": "主要488名單",
    "delta": 0.004674615384615777,
    "subGroup": "主要名單-東區",
    "region": "界線以東",
    "color": "綠色",
    "colorHex": "#257a4e"
  },
  {
    "id": "store-137",
    "name": "中正國小A系統",
    "address": "花蓮縣花蓮市中正路210號",
    "lat": 23.971884,
    "lng": 121.599571,
    "estimated": false,
    "isPrimary488": true,
    "group": "主要488名單",
    "delta": 0.004969615384609938,
    "subGroup": "主要名單-東區",
    "region": "界線以東",
    "color": "綠色",
    "colorHex": "#257a4e"
  },
  {
    "id": "store-138",
    "name": "中正國小B系統",
    "address": "花蓮縣花蓮市中正路210號",
    "lat": 23.971884,
    "lng": 121.599571,
    "estimated": false,
    "isPrimary488": true,
    "group": "主要488名單",
    "delta": 0.004969615384609938,
    "subGroup": "主要名單-東區",
    "region": "界線以東",
    "color": "綠色",
    "colorHex": "#257a4e"
  },
  {
    "id": "store-139",
    "name": "李學智婦產科診所",
    "address": "花蓮縣花蓮市中正路216號",
    "lat": 23.97209,
    "lng": 121.59999,
    "estimated": false,
    "isPrimary488": true,
    "group": "主要488名單",
    "delta": 0.0051588461538472075,
    "subGroup": "主要名單-東區",
    "region": "界線以東",
    "color": "綠色",
    "colorHex": "#257a4e"
  },
  {
    "id": "store-140",
    "name": "國泰世華銀行ATM全聯花蓮中正店",
    "address": "花蓮縣花蓮市中正路230號",
    "lat": 23.972414,
    "lng": 121.600308,
    "estimated": false,
    "isPrimary488": true,
    "group": "主要488名單",
    "delta": 0.005115461538451882,
    "subGroup": "主要名單-東區",
    "region": "界線以東",
    "color": "綠色",
    "colorHex": "#257a4e"
  },
  {
    "id": "store-141",
    "name": "全聯花蓮中正2111",
    "address": "花蓮縣花蓮市中正路230號",
    "lat": 23.972414,
    "lng": 121.600308,
    "estimated": false,
    "isPrimary488": true,
    "group": "主要488名單",
    "delta": 0.005115461538451882,
    "subGroup": "主要名單-東區",
    "region": "界線以東",
    "color": "綠色",
    "colorHex": "#257a4e"
  },
  {
    "id": "store-142",
    "name": "中正華廈速食店",
    "address": "花蓮縣花蓮市中正路231號",
    "lat": 24.02451,
    "lng": 121.61028,
    "estimated": false,
    "isPrimary488": false,
    "group": "非488名單(另色標記)",
    "subGroup": "非488名單",
    "region": "非488店家",
    "color": "黃色",
    "colorHex": "#f59e0b"
  },
  {
    "id": "store-143",
    "name": "東湳釣具行",
    "address": "花蓮縣花蓮市中正路29號",
    "lat": 24.024698,
    "lng": 121.611032,
    "estimated": false,
    "isPrimary488": true,
    "group": "主要488名單",
    "delta": -0.04247730769229463,
    "subGroup": "主要名單-西區",
    "region": "界線以西",
    "color": "紅色",
    "colorHex": "#e53e3e"
  },
  {
    "id": "store-144",
    "name": "呂外科診所",
    "address": "花蓮縣花蓮市中正路322號",
    "lat": 23.97366,
    "lng": 121.602182,
    "estimated": false,
    "isPrimary488": true,
    "group": "主要488名單",
    "delta": 0.005599692307697524,
    "subGroup": "主要名單-東區",
    "region": "界線以東",
    "color": "綠色",
    "colorHex": "#257a4e"
  },
  {
    "id": "store-145",
    "name": "花蓮一信國光分社",
    "address": "花蓮縣花蓮市中正路32號",
    "lat": 24.02413,
    "lng": 121.60936,
    "estimated": false,
    "isPrimary488": true,
    "group": "主要488名單",
    "delta": -0.0435157692307655,
    "subGroup": "主要名單-西區",
    "region": "界線以西",
    "color": "紅色",
    "colorHex": "#e53e3e"
  },
  {
    "id": "store-146",
    "name": "三商拿帕里花蓮店火災系統",
    "address": "花蓮縣花蓮市中正路342號",
    "lat": 23.973964,
    "lng": 121.602655,
    "estimated": false,
    "isPrimary488": true,
    "group": "主要488名單",
    "delta": 0.0057336153846136995,
    "subGroup": "主要名單-東區",
    "region": "界線以東",
    "color": "綠色",
    "colorHex": "#257a4e"
  },
  {
    "id": "store-147",
    "name": "三商拿帕里花蓮店",
    "address": "花蓮縣花蓮市中正路342號",
    "lat": 23.973964,
    "lng": 121.602655,
    "estimated": false,
    "isPrimary488": true,
    "group": "主要488名單",
    "delta": 0.0057336153846136995,
    "subGroup": "主要名單-東區",
    "region": "界線以東",
    "color": "綠色",
    "colorHex": "#257a4e"
  },
  {
    "id": "store-148",
    "name": "全國電子08020花蓮中正店",
    "address": "花蓮縣花蓮市中正路359.361.363號",
    "lat": 23.97413,
    "lng": 121.602409,
    "estimated": false,
    "isPrimary488": false,
    "group": "非488名單(另色標記)",
    "subGroup": "非488名單",
    "region": "非488店家",
    "color": "黃色",
    "colorHex": "#f59e0b"
  },
  {
    "id": "store-149",
    "name": "五州中醫診所",
    "address": "花蓮縣花蓮市中正路380號",
    "lat": 23.974493,
    "lng": 121.603421,
    "estimated": false,
    "isPrimary488": false,
    "group": "非488名單(另色標記)",
    "subGroup": "非488名單",
    "region": "非488店家",
    "color": "黃色",
    "colorHex": "#f59e0b"
  },
  {
    "id": "store-150",
    "name": "中英鐘錶眼鏡行",
    "address": "花蓮縣花蓮市中正路440號",
    "lat": 23.975414,
    "lng": 121.604763,
    "estimated": false,
    "isPrimary488": true,
    "group": "主要488名單",
    "delta": 0.006224307692306752,
    "subGroup": "主要名單-東區",
    "region": "界線以東",
    "color": "綠色",
    "colorHex": "#257a4e"
  },
  {
    "id": "store-151",
    "name": "曾記後山廣播電台3F.5F",
    "address": "花蓮縣花蓮市中正路446號",
    "lat": 23.97543,
    "lng": 121.60491,
    "estimated": false,
    "isPrimary488": true,
    "group": "主要488名單",
    "delta": 0.006353461538466831,
    "subGroup": "主要名單-東區",
    "region": "界線以東",
    "color": "綠色",
    "colorHex": "#257a4e"
  },
  {
    "id": "store-152",
    "name": "曾記麻糬公司中華門市1F.2F",
    "address": "花蓮縣花蓮市中正路446號",
    "lat": 23.97543,
    "lng": 121.60491,
    "estimated": false,
    "isPrimary488": true,
    "group": "主要488名單",
    "delta": 0.006353461538466831,
    "subGroup": "主要名單-東區",
    "region": "界線以東",
    "color": "綠色",
    "colorHex": "#257a4e"
  },
  {
    "id": "store-153",
    "name": "建華精品服飾店",
    "address": "花蓮縣花蓮市中正路462號",
    "lat": 23.976122,
    "lng": 121.605846,
    "estimated": false,
    "isPrimary488": true,
    "group": "主要488名單",
    "delta": 0.0065176153846095986,
    "subGroup": "主要名單-東區",
    "region": "界線以東",
    "color": "綠色",
    "colorHex": "#257a4e"
  },
  {
    "id": "store-154",
    "name": "職牛商號",
    "address": "花蓮縣花蓮市中正路474號2、3樓",
    "lat": 23.976479,
    "lng": 121.606436,
    "estimated": false,
    "isPrimary488": false,
    "group": "非488名單(另色標記)",
    "subGroup": "非488名單",
    "region": "非488店家",
    "color": "黃色",
    "colorHex": "#f59e0b"
  },
  {
    "id": "store-155",
    "name": "黃烈堂內科診所",
    "address": "花蓮縣花蓮市中正路482號",
    "lat": 23.976596,
    "lng": 121.606604,
    "estimated": false,
    "isPrimary488": true,
    "group": "主要488名單",
    "delta": 0.006746923076931921,
    "subGroup": "主要名單-東區",
    "region": "界線以東",
    "color": "綠色",
    "colorHex": "#257a4e"
  },
  {
    "id": "store-156",
    "name": "黛安芬直營花蓮中正店",
    "address": "花蓮縣花蓮市中正路499號",
    "lat": 23.97657,
    "lng": 121.606157,
    "estimated": false,
    "isPrimary488": false,
    "group": "非488名單(另色標記)",
    "subGroup": "非488名單",
    "region": "非488店家",
    "color": "黃色",
    "colorHex": "#f59e0b"
  },
  {
    "id": "store-157",
    "name": "瑞寶企業社",
    "address": "花蓮縣花蓮市中正路506號1樓",
    "lat": 23.976973,
    "lng": 121.607188,
    "estimated": false,
    "isPrimary488": true,
    "group": "主要488名單",
    "delta": 0.006910423076917027,
    "subGroup": "主要名單-東區",
    "region": "界線以東",
    "color": "綠色",
    "colorHex": "#257a4e"
  },
  {
    "id": "store-158",
    "name": "德誼數位科技花蓮店",
    "address": "花蓮縣花蓮市中正路511號1樓及3樓",
    "lat": 23.976733,
    "lng": 121.606404,
    "estimated": false,
    "isPrimary488": false,
    "group": "非488名單(另色標記)",
    "subGroup": "非488名單",
    "region": "非488店家",
    "color": "黃色",
    "colorHex": "#f59e0b"
  },
  {
    "id": "store-159",
    "name": "三立通訊配件行",
    "address": "花蓮縣花蓮市中正路520號1樓",
    "lat": 23.977438,
    "lng": 121.607868,
    "estimated": false,
    "isPrimary488": true,
    "group": "主要488名單",
    "delta": 0.00707176923076247,
    "subGroup": "主要名單-東區",
    "region": "界線以東",
    "color": "綠色",
    "colorHex": "#257a4e"
  },
  {
    "id": "store-160",
    "name": "冠毅服飾店住家",
    "address": "花蓮縣花蓮市中正路521號",
    "lat": 23.976853,
    "lng": 121.606611,
    "estimated": false,
    "isPrimary488": true,
    "group": "主要488名單",
    "delta": 0.0064672692307681245,
    "subGroup": "主要名單-東區",
    "region": "界線以東",
    "color": "綠色",
    "colorHex": "#257a4e"
  },
  {
    "id": "store-161",
    "name": "香港商捷時海外貿易花蓮門市",
    "address": "花蓮縣花蓮市中正路526號",
    "lat": 23.977552,
    "lng": 121.608051,
    "estimated": false,
    "isPrimary488": true,
    "group": "主要488名單",
    "delta": 0.00712761538461848,
    "subGroup": "主要名單-東區",
    "region": "界線以東",
    "color": "綠色",
    "colorHex": "#257a4e"
  },
  {
    "id": "store-162",
    "name": "三商行花蓮第二分公司",
    "address": "花蓮縣花蓮市中正路530號3樓",
    "lat": 23.977606,
    "lng": 121.608153,
    "estimated": false,
    "isPrimary488": true,
    "group": "主要488名單",
    "delta": 0.007169384615380636,
    "subGroup": "主要名單-東區",
    "region": "界線以東",
    "color": "綠色",
    "colorHex": "#257a4e"
  },
  {
    "id": "store-163",
    "name": "亞諾有限公司",
    "address": "花蓮縣花蓮市中正路549號",
    "lat": 23.977298,
    "lng": 121.607271,
    "estimated": false,
    "isPrimary488": true,
    "group": "主要488名單",
    "delta": 0.006630923076912154,
    "subGroup": "主要名單-東區",
    "region": "界線以東",
    "color": "綠色",
    "colorHex": "#257a4e"
  },
  {
    "id": "store-164",
    "name": "凱基銀行ATM花蓮分行",
    "address": "花蓮縣花蓮市中正路560-562號",
    "lat": 23.978077,
    "lng": 121.609049,
    "estimated": false,
    "isPrimary488": true,
    "group": "主要488名單",
    "delta": 0.0075400384615420535,
    "subGroup": "主要名單-東區",
    "region": "界線以東",
    "color": "綠色",
    "colorHex": "#257a4e"
  },
  {
    "id": "store-165",
    "name": "凱基銀行花蓮分行",
    "address": "花蓮縣花蓮市中正路560號(562號)",
    "lat": 23.978077,
    "lng": 121.609049,
    "estimated": false,
    "isPrimary488": true,
    "group": "主要488名單",
    "delta": 0.0075400384615420535,
    "subGroup": "主要名單-東區",
    "region": "界線以東",
    "color": "綠色",
    "colorHex": "#257a4e"
  },
  {
    "id": "store-166",
    "name": "金珍時尚珠寶",
    "address": "花蓮縣花蓮市中正路562-1號",
    "lat": 23.978098,
    "lng": 121.609086,
    "estimated": false,
    "isPrimary488": false,
    "group": "非488名單(另色標記)",
    "subGroup": "非488名單",
    "region": "非488店家",
    "color": "黃色",
    "colorHex": "#f59e0b"
  },
  {
    "id": "store-167",
    "name": "環球通訊行",
    "address": "花蓮縣花蓮市中正路566號",
    "lat": 23.978156,
    "lng": 121.609161,
    "estimated": false,
    "isPrimary488": true,
    "group": "主要488名單",
    "delta": 0.007563923076929768,
    "subGroup": "主要名單-東區",
    "region": "界線以東",
    "color": "綠色",
    "colorHex": "#257a4e"
  },
  {
    "id": "store-168",
    "name": "長榮通訊行",
    "address": "花蓮縣花蓮市中正路568號1樓",
    "lat": 23.97831,
    "lng": 121.609214,
    "estimated": false,
    "isPrimary488": true,
    "group": "主要488名單",
    "delta": 0.007445153846148855,
    "subGroup": "主要名單-東區",
    "region": "界線以東",
    "color": "綠色",
    "colorHex": "#257a4e"
  },
  {
    "id": "store-169",
    "name": "張家華營業所娃娃機商店",
    "address": "花蓮縣花蓮市中正路571號",
    "lat": 23.977744,
    "lng": 121.607918,
    "estimated": false,
    "isPrimary488": true,
    "group": "主要488名單",
    "delta": 0.006780461538454574,
    "subGroup": "主要名單-東區",
    "region": "界線以東",
    "color": "綠色",
    "colorHex": "#257a4e"
  },
  {
    "id": "store-170",
    "name": "宣寶石珊瑚行",
    "address": "花蓮縣花蓮市中正路572號",
    "lat": 23.978438,
    "lng": 121.609438,
    "estimated": false,
    "isPrimary488": true,
    "group": "主要488名單",
    "delta": 0.007526384615374582,
    "subGroup": "主要名單-東區",
    "region": "界線以東",
    "color": "綠色",
    "colorHex": "#257a4e"
  },
  {
    "id": "store-171",
    "name": "新光銀行花蓮分行司反偵測",
    "address": "花蓮縣花蓮市中正路588號",
    "lat": 23.978717,
    "lng": 121.609951,
    "estimated": false,
    "isPrimary488": true,
    "group": "主要488名單",
    "delta": 0.00772819230768107,
    "subGroup": "主要名單-東區",
    "region": "界線以東",
    "color": "綠色",
    "colorHex": "#257a4e"
  },
  {
    "id": "store-172",
    "name": "新光銀行花蓮分行",
    "address": "花蓮縣花蓮市中正路588號",
    "lat": 23.978717,
    "lng": 121.609951,
    "estimated": false,
    "isPrimary488": true,
    "group": "主要488名單",
    "delta": 0.00772819230768107,
    "subGroup": "主要名單-東區",
    "region": "界線以東",
    "color": "綠色",
    "colorHex": "#257a4e"
  },
  {
    "id": "store-173",
    "name": "花蓮二信中正分社",
    "address": "花蓮縣花蓮市中正路58號",
    "lat": 23.969425,
    "lng": 121.595692,
    "estimated": false,
    "isPrimary488": true,
    "group": "主要488名單",
    "delta": 0.0038333461538400115,
    "subGroup": "主要名單-東區",
    "region": "界線以東",
    "color": "綠色",
    "colorHex": "#257a4e"
  },
  {
    "id": "store-174",
    "name": "凌群電腦股份有限公司花蓮店",
    "address": "花蓮縣花蓮市中正路60號",
    "lat": 23.969446,
    "lng": 121.595785,
    "estimated": false,
    "isPrimary488": true,
    "group": "主要488名單",
    "delta": 0.0039029230769216383,
    "subGroup": "主要名單-東區",
    "region": "界線以東",
    "color": "綠色",
    "colorHex": "#257a4e"
  },
  {
    "id": "store-175",
    "name": "誠食生活家-將軍府-A棟",
    "address": "花蓮縣花蓮市中正路622巷6號",
    "lat": 23.980467,
    "lng": 121.61313,
    "estimated": false,
    "isPrimary488": true,
    "group": "主要488名單",
    "delta": 0.008955269230767726,
    "subGroup": "主要名單-東區",
    "region": "界線以東",
    "color": "綠色",
    "colorHex": "#257a4e"
  },
  {
    "id": "store-176",
    "name": "誠食生活家-將軍府C棟",
    "address": "花蓮縣花蓮市中正路622巷6號",
    "lat": 23.980467,
    "lng": 121.61313,
    "estimated": false,
    "isPrimary488": true,
    "group": "主要488名單",
    "delta": 0.008955269230767726,
    "subGroup": "主要名單-東區",
    "region": "界線以東",
    "color": "綠色",
    "colorHex": "#257a4e"
  },
  {
    "id": "store-177",
    "name": "誠食生活家-將軍府D棟",
    "address": "花蓮縣花蓮市中正路622巷6號",
    "lat": 23.980467,
    "lng": 121.61313,
    "estimated": false,
    "isPrimary488": true,
    "group": "主要488名單",
    "delta": 0.008955269230767726,
    "subGroup": "主要名單-東區",
    "region": "界線以東",
    "color": "綠色",
    "colorHex": "#257a4e"
  },
  {
    "id": "store-178",
    "name": "誠食生活家-將軍府E棟",
    "address": "花蓮縣花蓮市中正路622巷6號",
    "lat": 23.980467,
    "lng": 121.61313,
    "estimated": false,
    "isPrimary488": true,
    "group": "主要488名單",
    "delta": 0.008955269230767726,
    "subGroup": "主要名單-東區",
    "region": "界線以東",
    "color": "綠色",
    "colorHex": "#257a4e"
  },
  {
    "id": "store-179",
    "name": "誠食生活家-將軍府F棟",
    "address": "花蓮縣花蓮市中正路622巷6號",
    "lat": 23.980467,
    "lng": 121.61313,
    "estimated": false,
    "isPrimary488": true,
    "group": "主要488名單",
    "delta": 0.008955269230767726,
    "subGroup": "主要名單-東區",
    "region": "界線以東",
    "color": "綠色",
    "colorHex": "#257a4e"
  },
  {
    "id": "store-180",
    "name": "誠食-將軍府G棟(G1)",
    "address": "花蓮縣花蓮市中正路622巷6號",
    "lat": 23.980467,
    "lng": 121.61313,
    "estimated": false,
    "isPrimary488": true,
    "group": "主要488名單",
    "delta": 0.008955269230767726,
    "subGroup": "主要名單-東區",
    "region": "界線以東",
    "color": "綠色",
    "colorHex": "#257a4e"
  },
  {
    "id": "store-181",
    "name": "誠食-將軍府G棟(G2)",
    "address": "花蓮縣花蓮市中正路622巷6號",
    "lat": 23.980467,
    "lng": 121.61313,
    "estimated": false,
    "isPrimary488": true,
    "group": "主要488名單",
    "delta": 0.008955269230767726,
    "subGroup": "主要名單-東區",
    "region": "界線以東",
    "color": "綠色",
    "colorHex": "#257a4e"
  },
  {
    "id": "store-182",
    "name": "誠食-將軍府H棟(H1)",
    "address": "花蓮縣花蓮市中正路622巷6號",
    "lat": 23.980467,
    "lng": 121.61313,
    "estimated": false,
    "isPrimary488": true,
    "group": "主要488名單",
    "delta": 0.008955269230767726,
    "subGroup": "主要名單-東區",
    "region": "界線以東",
    "color": "綠色",
    "colorHex": "#257a4e"
  },
  {
    "id": "store-183",
    "name": "誠食-將軍府H棟(H2)",
    "address": "花蓮縣花蓮市中正路622巷6號",
    "lat": 23.980467,
    "lng": 121.61313,
    "estimated": false,
    "isPrimary488": true,
    "group": "主要488名單",
    "delta": 0.008955269230767726,
    "subGroup": "主要名單-東區",
    "region": "界線以東",
    "color": "綠色",
    "colorHex": "#257a4e"
  },
  {
    "id": "store-184",
    "name": "誠食生活家-將軍府儲藏室",
    "address": "花蓮縣花蓮市中正路622巷6號",
    "lat": 23.980467,
    "lng": 121.61313,
    "estimated": false,
    "isPrimary488": true,
    "group": "主要488名單",
    "delta": 0.008955269230767726,
    "subGroup": "主要名單-東區",
    "region": "界線以東",
    "color": "綠色",
    "colorHex": "#257a4e"
  },
  {
    "id": "store-185",
    "name": "誠食生活家-將軍府咖啡店",
    "address": "花蓮縣花蓮市中正路622巷6號",
    "lat": 23.980467,
    "lng": 121.61313,
    "estimated": false,
    "isPrimary488": true,
    "group": "主要488名單",
    "delta": 0.008955269230767726,
    "subGroup": "主要名單-東區",
    "region": "界線以東",
    "color": "綠色",
    "colorHex": "#257a4e"
  },
  {
    "id": "store-186",
    "name": "誠食生活家-將軍府(管理室)",
    "address": "花蓮縣花蓮市中正路622巷6號",
    "lat": 23.980467,
    "lng": 121.61313,
    "estimated": false,
    "isPrimary488": true,
    "group": "主要488名單",
    "delta": 0.008955269230767726,
    "subGroup": "主要名單-東區",
    "region": "界線以東",
    "color": "綠色",
    "colorHex": "#257a4e"
  },
  {
    "id": "store-187",
    "name": "誠食生活家-將軍府B棟",
    "address": "花蓮縣花蓮市中正路622巷6號",
    "lat": 23.980467,
    "lng": 121.61313,
    "estimated": false,
    "isPrimary488": true,
    "group": "主要488名單",
    "delta": 0.008955269230767726,
    "subGroup": "主要名單-東區",
    "region": "界線以東",
    "color": "綠色",
    "colorHex": "#257a4e"
  },
  {
    "id": "store-188",
    "name": "台灣麥當勞和德昌花蓮中正路店",
    "address": "花蓮縣花蓮市中正路625號B1-2樓",
    "lat": 23.979181,
    "lng": 121.610077,
    "estimated": false,
    "isPrimary488": true,
    "group": "主要488名單",
    "delta": 0.007336653846152785,
    "subGroup": "主要名單-東區",
    "region": "界線以東",
    "color": "綠色",
    "colorHex": "#257a4e"
  },
  {
    "id": "store-189",
    "name": "富利餐飲肯德基花蓮中正店",
    "address": "花蓮縣花蓮市中正路637-1號",
    "lat": 23.979404,
    "lng": 121.610426,
    "estimated": false,
    "isPrimary488": true,
    "group": "主要488名單",
    "delta": 0.007436923076923563,
    "subGroup": "主要名單-東區",
    "region": "界線以東",
    "color": "綠色",
    "colorHex": "#257a4e"
  },
  {
    "id": "store-190",
    "name": "嘉豐健康廚房",
    "address": "花蓮縣花蓮市中正路76號",
    "lat": 23.969758,
    "lng": 121.596212,
    "estimated": false,
    "isPrimary488": false,
    "group": "非488名單(另色標記)",
    "subGroup": "非488名單",
    "region": "非488店家",
    "color": "黃色",
    "colorHex": "#f59e0b"
  },
  {
    "id": "store-191",
    "name": "高橋動物醫院",
    "address": "花蓮縣花蓮市中正路91號",
    "lat": 24.023994,
    "lng": 121.60913,
    "estimated": false,
    "isPrimary488": true,
    "group": "主要488名單",
    "delta": -0.04359407692307116,
    "subGroup": "主要名單-西區",
    "region": "界線以西",
    "color": "紅色",
    "colorHex": "#e53e3e"
  },
  {
    "id": "store-192",
    "name": "馥邑京華大廈管理委員會",
    "address": "花蓮縣花蓮市中和街207號",
    "lat": 23.974469,
    "lng": 121.59813,
    "estimated": false,
    "isPrimary488": false,
    "group": "非488名單(另色標記)",
    "subGroup": "非488名單",
    "region": "非488店家",
    "color": "黃色",
    "colorHex": "#f59e0b"
  },
  {
    "id": "store-193",
    "name": "本居立有限公司",
    "address": "花蓮縣花蓮市中原路294號",
    "lat": 23.967238,
    "lng": 121.601165,
    "estimated": false,
    "isPrimary488": true,
    "group": "主要488名單",
    "delta": 0.011745692307684408,
    "subGroup": "主要名單-東區",
    "region": "界線以東",
    "color": "綠色",
    "colorHex": "#257a4e"
  },
  {
    "id": "store-194",
    "name": "美麗坊自助洗衣店",
    "address": "花蓮縣花蓮市中原路394巷6號",
    "lat": 23.969632,
    "lng": 121.599287,
    "estimated": false,
    "isPrimary488": true,
    "group": "主要488名單",
    "delta": 0.007197461538467564,
    "subGroup": "主要名單-東區",
    "region": "界線以東",
    "color": "綠色",
    "colorHex": "#257a4e"
  },
  {
    "id": "store-195",
    "name": "花蓮縣花蓮市中原國民小學",
    "address": "花蓮縣花蓮市中原路531號",
    "lat": 23.9721,
    "lng": 121.59641,
    "estimated": false,
    "isPrimary488": true,
    "group": "主要488名單",
    "delta": 0.0015676923076881621,
    "subGroup": "主要名單-東區",
    "region": "界線以東",
    "color": "綠色",
    "colorHex": "#257a4e"
  },
  {
    "id": "store-196",
    "name": "賴桑海鮮有限公司倉庫",
    "address": "花蓮縣花蓮市中原路575號",
    "lat": 23.973208,
    "lng": 121.595452,
    "estimated": false,
    "isPrimary488": false,
    "group": "非488名單(另色標記)",
    "subGroup": "非488名單",
    "region": "非488店家",
    "color": "黃色",
    "colorHex": "#f59e0b"
  },
  {
    "id": "store-197",
    "name": "黛恩尼服飾名店",
    "address": "花蓮縣花蓮市中原路622號",
    "lat": 23.974105,
    "lng": 121.594954,
    "estimated": false,
    "isPrimary488": true,
    "group": "主要488名單",
    "delta": -0.002124653846152569,
    "subGroup": "主要名單-西區",
    "region": "界線以西",
    "color": "紅色",
    "colorHex": "#e53e3e"
  },
  {
    "id": "store-198",
    "name": "提拉米蘇精緻蛋糕B棟",
    "address": "花蓮縣花蓮市中原路738號",
    "lat": 23.976975,
    "lng": 121.592062,
    "estimated": false,
    "isPrimary488": true,
    "group": "主要488名單",
    "delta": -0.008217807692304291,
    "subGroup": "主要名單-西區",
    "region": "界線以西",
    "color": "紅色",
    "colorHex": "#e53e3e"
  },
  {
    "id": "store-199",
    "name": "提拉米蘇精緻蛋糕",
    "address": "花蓮縣花蓮市中原路738號",
    "lat": 23.976975,
    "lng": 121.592062,
    "estimated": false,
    "isPrimary488": true,
    "group": "主要488名單",
    "delta": -0.008217807692304291,
    "subGroup": "主要名單-西區",
    "region": "界線以西",
    "color": "紅色",
    "colorHex": "#e53e3e"
  },
  {
    "id": "store-200",
    "name": "提拉米蘇精緻蛋糕防災系統",
    "address": "花蓮縣花蓮市中原路738號",
    "lat": 23.976975,
    "lng": 121.592062,
    "estimated": false,
    "isPrimary488": true,
    "group": "主要488名單",
    "delta": -0.008217807692304291,
    "subGroup": "主要名單-西區",
    "region": "界線以西",
    "color": "紅色",
    "colorHex": "#e53e3e"
  },
  {
    "id": "store-201",
    "name": "陳坤亮公館",
    "address": "花蓮縣花蓮市中強街126-1號",
    "lat": 23.96951,
    "lng": 121.601406,
    "estimated": false,
    "isPrimary488": true,
    "group": "主要488名單",
    "delta": 0.009452538461530935,
    "subGroup": "主要名單-東區",
    "region": "界線以東",
    "color": "綠色",
    "colorHex": "#257a4e"
  },
  {
    "id": "store-202",
    "name": "臻好翫休閒運動事業博愛門市",
    "address": "花蓮縣花蓮市中華路101號",
    "lat": 23.97571,
    "lng": 121.60716,
    "estimated": false,
    "isPrimary488": true,
    "group": "主要488名單",
    "delta": 0.008291153846144539,
    "subGroup": "主要名單-東區",
    "region": "界線以東",
    "color": "綠色",
    "colorHex": "#257a4e"
  },
  {
    "id": "store-203",
    "name": "林千仕營業所",
    "address": "花蓮縣花蓮市中華路104號",
    "lat": 23.975779,
    "lng": 121.604751,
    "estimated": false,
    "isPrimary488": false,
    "group": "非488名單(另色標記)",
    "subGroup": "非488名單",
    "region": "非488店家",
    "color": "黃色",
    "colorHex": "#f59e0b"
  },
  {
    "id": "store-204",
    "name": "黃詠翔彩卷行",
    "address": "花蓮縣花蓮市中華路112號",
    "lat": 23.97599,
    "lng": 121.6058,
    "estimated": false,
    "isPrimary488": true,
    "group": "主要488名單",
    "delta": 0.006618846153841673,
    "subGroup": "主要名單-東區",
    "region": "界線以東",
    "color": "綠色",
    "colorHex": "#257a4e"
  },
  {
    "id": "store-205",
    "name": "雅婷智慧股份有限公司",
    "address": "花蓮縣花蓮市中華路144號第19棟",
    "lat": 23.975833,
    "lng": 121.604534,
    "estimated": true,
    "isPrimary488": false,
    "group": "非488名單(另色標記)",
    "subGroup": "非488名單",
    "region": "非488店家",
    "color": "黃色",
    "colorHex": "#f59e0b"
  },
  {
    "id": "store-206",
    "name": "協源工具行",
    "address": "花蓮縣花蓮市中華路166號",
    "lat": 23.97582,
    "lng": 121.60357,
    "estimated": false,
    "isPrimary488": true,
    "group": "主要488名單",
    "delta": 0.004578461538471856,
    "subGroup": "主要名單-東區",
    "region": "界線以東",
    "color": "綠色",
    "colorHex": "#257a4e"
  },
  {
    "id": "store-207",
    "name": "陳韻琳營業所-娃娃機商店",
    "address": "花蓮縣花蓮市中華路18、20號",
    "lat": 23.97634,
    "lng": 121.60883,
    "estimated": false,
    "isPrimary488": false,
    "group": "非488名單(另色標記)",
    "subGroup": "非488名單",
    "region": "非488店家",
    "color": "黃色",
    "colorHex": "#f59e0b"
  },
  {
    "id": "store-208",
    "name": "花蓮一信信義分社",
    "address": "花蓮縣花蓮市中華路180號",
    "lat": 23.97578,
    "lng": 121.60304,
    "estimated": false,
    "isPrimary488": true,
    "group": "主要488名單",
    "delta": 0.004093076923069816,
    "subGroup": "主要名單-東區",
    "region": "界線以東",
    "color": "綠色",
    "colorHex": "#257a4e"
  },
  {
    "id": "store-209",
    "name": "肉肉豐盛餐飲有限公司",
    "address": "花蓮縣花蓮市中華路183號",
    "lat": 23.97533,
    "lng": 121.60384,
    "estimated": false,
    "isPrimary488": true,
    "group": "主要488名單",
    "delta": 0.005395000000007144,
    "subGroup": "主要名單-東區",
    "region": "界線以東",
    "color": "綠色",
    "colorHex": "#257a4e"
  },
  {
    "id": "store-210",
    "name": "張祖清營業所",
    "address": "花蓮縣花蓮市中華路24號",
    "lat": 23.976287,
    "lng": 121.608557,
    "estimated": false,
    "isPrimary488": true,
    "group": "主要488名單",
    "delta": 0.009044576923088243,
    "subGroup": "主要名單-東區",
    "region": "界線以東",
    "color": "綠色",
    "colorHex": "#257a4e"
  },
  {
    "id": "store-211",
    "name": "天泉銀樓",
    "address": "花蓮縣花蓮市中華路25號",
    "lat": 23.97604,
    "lng": 121.60933,
    "estimated": false,
    "isPrimary488": true,
    "group": "主要488名單",
    "delta": 0.010093076923070043,
    "subGroup": "主要名單-東區",
    "region": "界線以東",
    "color": "綠色",
    "colorHex": "#257a4e"
  },
  {
    "id": "store-212",
    "name": "陳韻琳-娃娃機",
    "address": "花蓮縣花蓮市中華路264之三號",
    "lat": 23.97554,
    "lng": 121.60044,
    "estimated": false,
    "isPrimary488": false,
    "group": "非488名單(另色標記)",
    "subGroup": "非488名單",
    "region": "非488店家",
    "color": "黃色",
    "colorHex": "#f59e0b"
  },
  {
    "id": "store-213",
    "name": "怡信鞋店企業社",
    "address": "花蓮縣花蓮市中華路27號",
    "lat": 23.97598,
    "lng": 121.6093,
    "estimated": false,
    "isPrimary488": true,
    "group": "主要488名單",
    "delta": 0.010130000000003747,
    "subGroup": "主要名單-東區",
    "region": "界線以東",
    "color": "綠色",
    "colorHex": "#257a4e"
  },
  {
    "id": "store-214",
    "name": "瀧寅醉居",
    "address": "花蓮縣花蓮市中華路289號",
    "lat": 23.975003,
    "lng": 121.599662,
    "estimated": false,
    "isPrimary488": false,
    "group": "非488名單(另色標記)",
    "subGroup": "非488名單",
    "region": "非488店家",
    "color": "黃色",
    "colorHex": "#f59e0b"
  },
  {
    "id": "store-215",
    "name": "印地安那服飾店",
    "address": "花蓮縣花蓮市中華路32號1樓",
    "lat": 23.97644,
    "lng": 121.60833,
    "estimated": false,
    "isPrimary488": false,
    "group": "非488名單(另色標記)",
    "subGroup": "非488名單",
    "region": "非488店家",
    "color": "黃色",
    "colorHex": "#f59e0b"
  },
  {
    "id": "store-216",
    "name": "中油中華加油站",
    "address": "花蓮縣花蓮市中華路333號",
    "lat": 23.974754,
    "lng": 121.597374,
    "estimated": false,
    "isPrimary488": true,
    "group": "主要488名單",
    "delta": -0.0004285384615343446,
    "subGroup": "主要名單-西區",
    "region": "界線以西",
    "color": "紅色",
    "colorHex": "#e53e3e"
  },
  {
    "id": "store-217",
    "name": "Sammi空間髮藝",
    "address": "花蓮縣花蓮市中華路366號",
    "lat": 23.97501,
    "lng": 121.59601,
    "estimated": false,
    "isPrimary488": false,
    "group": "非488名單(另色標記)",
    "subGroup": "非488名單",
    "region": "非488店家",
    "color": "黃色",
    "colorHex": "#f59e0b"
  },
  {
    "id": "store-218",
    "name": "新橋管理顧問有限公司",
    "address": "花蓮縣花蓮市中華路39號",
    "lat": 23.97603,
    "lng": 121.60898,
    "estimated": false,
    "isPrimary488": false,
    "group": "非488名單(另色標記)",
    "subGroup": "非488名單",
    "region": "非488店家",
    "color": "黃色",
    "colorHex": "#f59e0b"
  },
  {
    "id": "store-219",
    "name": "陳衍瑞營業所",
    "address": "花蓮縣花蓮市中華路407號",
    "lat": 23.97438,
    "lng": 121.59366,
    "estimated": false,
    "isPrimary488": true,
    "group": "主要488名單",
    "delta": -0.0037253846153788572,
    "subGroup": "主要名單-西區",
    "region": "界線以西",
    "color": "紅色",
    "colorHex": "#e53e3e"
  },
  {
    "id": "store-220",
    "name": "徐瑋志公館",
    "address": "花蓮縣花蓮市中華路410號",
    "lat": 23.97482,
    "lng": 121.59492,
    "estimated": false,
    "isPrimary488": true,
    "group": "主要488名單",
    "delta": -0.0029561538461564396,
    "subGroup": "主要名單-西區",
    "region": "界線以西",
    "color": "紅色",
    "colorHex": "#e53e3e"
  },
  {
    "id": "store-221",
    "name": "林秀香富美便當",
    "address": "花蓮縣花蓮市中華路435-3號",
    "lat": 23.974339,
    "lng": 121.592834,
    "estimated": false,
    "isPrimary488": false,
    "group": "非488名單(另色標記)",
    "subGroup": "非488名單",
    "region": "非488店家",
    "color": "黃色",
    "colorHex": "#f59e0b"
  },
  {
    "id": "store-222",
    "name": "花蓮市農會田埔生鮮超級市場",
    "address": "花蓮縣花蓮市中華路459號",
    "lat": 23.974037,
    "lng": 121.590389,
    "estimated": false,
    "isPrimary488": true,
    "group": "主要488名單",
    "delta": -0.006613807692303908,
    "subGroup": "主要名單-西區",
    "region": "界線以西",
    "color": "紅色",
    "colorHex": "#e53e3e"
  },
  {
    "id": "store-223",
    "name": "國泰聯合診所",
    "address": "花蓮縣花蓮市中華路459號",
    "lat": 23.974037,
    "lng": 121.590389,
    "estimated": false,
    "isPrimary488": true,
    "group": "主要488名單",
    "delta": -0.006613807692303908,
    "subGroup": "主要名單-西區",
    "region": "界線以西",
    "color": "紅色",
    "colorHex": "#e53e3e"
  },
  {
    "id": "store-224",
    "name": "花蓮縣農會生鮮超市",
    "address": "花蓮縣花蓮市中華路459號1-3樓",
    "lat": 23.974037,
    "lng": 121.590389,
    "estimated": false,
    "isPrimary488": true,
    "group": "主要488名單",
    "delta": -0.006613807692303908,
    "subGroup": "主要名單-西區",
    "region": "界線以西",
    "color": "紅色",
    "colorHex": "#e53e3e"
  },
  {
    "id": "store-225",
    "name": "花蓮一信中華分社",
    "address": "花蓮縣花蓮市中華路469號",
    "lat": 23.974084,
    "lng": 121.589841,
    "estimated": false,
    "isPrimary488": true,
    "group": "主要488名單",
    "delta": -0.0072142307692217855,
    "subGroup": "主要名單-西區",
    "region": "界線以西",
    "color": "紅色",
    "colorHex": "#e53e3e"
  },
  {
    "id": "store-226",
    "name": "名捷自行車有限公司",
    "address": "花蓮縣花蓮市中華路491號",
    "lat": 23.974013,
    "lng": 121.588849,
    "estimated": false,
    "isPrimary488": true,
    "group": "主要488名單",
    "delta": -0.00812703846153795,
    "subGroup": "主要名單-西區",
    "region": "界線以西",
    "color": "紅色",
    "colorHex": "#e53e3e"
  },
  {
    "id": "store-227",
    "name": "禾美建設有限公司",
    "address": "花蓮縣花蓮市中華路498巷1弄13號",
    "lat": 23.974755,
    "lng": 121.591166,
    "estimated": false,
    "isPrimary488": false,
    "group": "非488名單(另色標記)",
    "subGroup": "非488名單",
    "region": "非488店家",
    "color": "黃色",
    "colorHex": "#f59e0b"
  },
  {
    "id": "store-228",
    "name": "蔡素美御鑽銀樓",
    "address": "花蓮縣花蓮市中華路502號",
    "lat": 23.974518,
    "lng": 121.591184,
    "estimated": false,
    "isPrimary488": true,
    "group": "主要488名單",
    "delta": -0.006355307692317069,
    "subGroup": "主要名單-西區",
    "region": "界線以西",
    "color": "紅色",
    "colorHex": "#e53e3e"
  },
  {
    "id": "store-229",
    "name": "鈺發銀樓",
    "address": "花蓮縣花蓮市中華路50號",
    "lat": 23.97622,
    "lng": 121.60782,
    "estimated": false,
    "isPrimary488": true,
    "group": "主要488名單",
    "delta": 0.008382307692315294,
    "subGroup": "主要名單-東區",
    "region": "界線以東",
    "color": "綠色",
    "colorHex": "#257a4e"
  },
  {
    "id": "store-230",
    "name": "農糧署東區分署",
    "address": "花蓮縣花蓮市中華路512號",
    "lat": 23.974397,
    "lng": 121.590469,
    "estimated": false,
    "isPrimary488": true,
    "group": "主要488名單",
    "delta": -0.006935346153852606,
    "subGroup": "主要名單-西區",
    "region": "界線以西",
    "color": "紅色",
    "colorHex": "#e53e3e"
  },
  {
    "id": "store-231",
    "name": "農糧署東區分署",
    "address": "花蓮縣花蓮市中華路512號",
    "lat": 23.974397,
    "lng": 121.590469,
    "estimated": false,
    "isPrimary488": true,
    "group": "主要488名單",
    "delta": -0.006935346153852606,
    "subGroup": "主要名單-西區",
    "region": "界線以西",
    "color": "紅色",
    "colorHex": "#e53e3e"
  },
  {
    "id": "store-232",
    "name": "賴資凱彩券行-花蓮店",
    "address": "花蓮縣花蓮市中華路85號",
    "lat": 23.9759,
    "lng": 121.60772,
    "estimated": false,
    "isPrimary488": true,
    "group": "主要488名單",
    "delta": 0.008639230769233563,
    "subGroup": "主要名單-東區",
    "region": "界線以東",
    "color": "綠色",
    "colorHex": "#257a4e"
  },
  {
    "id": "store-233",
    "name": "瑪格莉特生活精品館",
    "address": "花蓮縣花蓮市中華路86號",
    "lat": 23.97623,
    "lng": 121.60672,
    "estimated": false,
    "isPrimary488": true,
    "group": "主要488名單",
    "delta": 0.007271153846147627,
    "subGroup": "主要名單-東區",
    "region": "界線以東",
    "color": "綠色",
    "colorHex": "#257a4e"
  },
  {
    "id": "store-234",
    "name": "王馥華稅務記帳士事務所",
    "address": "花蓮縣花蓮市中順街102號",
    "lat": 23.968983,
    "lng": 121.596974,
    "estimated": false,
    "isPrimary488": true,
    "group": "主要488名單",
    "delta": 0.0056083461538491974,
    "subGroup": "主要名單-東區",
    "region": "界線以東",
    "color": "綠色",
    "colorHex": "#257a4e"
  },
  {
    "id": "store-235",
    "name": "全人診所",
    "address": "花蓮縣花蓮市中福路158號",
    "lat": 23.972664,
    "lng": 121.601724,
    "estimated": false,
    "isPrimary488": true,
    "group": "主要488名單",
    "delta": 0.006252615384610749,
    "subGroup": "主要名單-東區",
    "region": "界線以東",
    "color": "綠色",
    "colorHex": "#257a4e"
  },
  {
    "id": "store-236",
    "name": "朱震雄倉庫",
    "address": "花蓮縣花蓮市中福路167號",
    "lat": 23.9727,
    "lng": 121.60142,
    "estimated": false,
    "isPrimary488": true,
    "group": "主要488名單",
    "delta": 0.005908461538467691,
    "subGroup": "主要名單-東區",
    "region": "界線以東",
    "color": "綠色",
    "colorHex": "#257a4e"
  },
  {
    "id": "store-237",
    "name": "朱美雪自助洗衣中福店",
    "address": "花蓮縣花蓮市中福路192號",
    "lat": 23.973665,
    "lng": 121.60093,
    "estimated": false,
    "isPrimary488": true,
    "group": "主要488名單",
    "delta": 0.004342115384616818,
    "subGroup": "主要名單-東區",
    "region": "界線以東",
    "color": "綠色",
    "colorHex": "#257a4e"
  },
  {
    "id": "store-238",
    "name": "林建中診所",
    "address": "花蓮縣花蓮市中福路193號",
    "lat": 23.973631,
    "lng": 121.600682,
    "estimated": false,
    "isPrimary488": true,
    "group": "主要488名單",
    "delta": 0.004132038461548859,
    "subGroup": "主要名單-東區",
    "region": "界線以東",
    "color": "綠色",
    "colorHex": "#257a4e"
  },
  {
    "id": "store-239",
    "name": "花蓮二信主里分社",
    "address": "花蓮縣花蓮市中福路232號",
    "lat": 23.974711,
    "lng": 121.60015,
    "estimated": false,
    "isPrimary488": true,
    "group": "主要488名單",
    "delta": 0.002395423076919201,
    "subGroup": "主要名單-東區",
    "region": "界線以東",
    "color": "綠色",
    "colorHex": "#257a4e"
  },
  {
    "id": "store-240",
    "name": "經濟部水利署第九河川分署",
    "address": "花蓮縣花蓮市仁愛街19號",
    "lat": 23.972932,
    "lng": 121.606163,
    "estimated": false,
    "isPrimary488": true,
    "group": "主要488名單",
    "delta": 0.0103926923076898,
    "subGroup": "主要名單-東區",
    "region": "界線以東",
    "color": "綠色",
    "colorHex": "#257a4e"
  },
  {
    "id": "store-241",
    "name": "台灣菸酒花蓮營業處",
    "address": "花蓮縣花蓮市仁愛街75號",
    "lat": 23.976527,
    "lng": 121.603222,
    "estimated": false,
    "isPrimary488": true,
    "group": "主要488名單",
    "delta": 0.003441884615384083,
    "subGroup": "主要名單-東區",
    "region": "界線以東",
    "color": "綠色",
    "colorHex": "#257a4e"
  },
  {
    "id": "store-242",
    "name": "甄饌私廚酒館有限公司",
    "address": "花蓮縣花蓮市仁愛街91、95號",
    "lat": 23.978307,
    "lng": 121.602049,
    "estimated": false,
    "isPrimary488": true,
    "group": "主要488名單",
    "delta": 0.00028349999999477404,
    "subGroup": "主要名單-東區",
    "region": "界線以東",
    "color": "綠色",
    "colorHex": "#257a4e"
  },
  {
    "id": "store-243",
    "name": "張啟仲營業所",
    "address": "花蓮縣花蓮市公正街6號",
    "lat": 23.97751,
    "lng": 121.61011,
    "estimated": false,
    "isPrimary488": true,
    "group": "主要488名單",
    "delta": 0.009233461538471488,
    "subGroup": "主要名單-東區",
    "region": "界線以東",
    "color": "綠色",
    "colorHex": "#257a4e"
  },
  {
    "id": "store-244",
    "name": "花蓮縣政府老人會館",
    "address": "花蓮縣花蓮市公園路39號",
    "lat": 23.976725,
    "lng": 121.610375,
    "estimated": false,
    "isPrimary488": true,
    "group": "主要488名單",
    "delta": 0.010374038461549162,
    "subGroup": "主要名單-東區",
    "region": "界線以東",
    "color": "綠色",
    "colorHex": "#257a4e"
  },
  {
    "id": "store-245",
    "name": "台灣銀行花蓮分行",
    "address": "花蓮縣花蓮市公園路3號",
    "lat": 23.976996,
    "lng": 121.610664,
    "estimated": false,
    "isPrimary488": true,
    "group": "主要488名單",
    "delta": 0.010360769230771893,
    "subGroup": "主要名單-東區",
    "region": "界線以東",
    "color": "綠色",
    "colorHex": "#257a4e"
  },
  {
    "id": "store-246",
    "name": "台灣銀行ATM反偵測花蓮分行",
    "address": "花蓮縣花蓮市公園路3號",
    "lat": 23.976996,
    "lng": 121.610664,
    "estimated": false,
    "isPrimary488": false,
    "group": "非488名單(另色標記)",
    "subGroup": "非488名單",
    "region": "非488店家",
    "color": "黃色",
    "colorHex": "#f59e0b"
  },
  {
    "id": "store-247",
    "name": "張禎峰公館",
    "address": "花蓮縣花蓮市北昌三街56號",
    "lat": 23.987183,
    "lng": 121.594172,
    "estimated": false,
    "isPrimary488": true,
    "group": "主要488名單",
    "delta": -0.017493653846159418,
    "subGroup": "主要名單-西區",
    "region": "界線以西",
    "color": "紅色",
    "colorHex": "#e53e3e"
  },
  {
    "id": "store-248",
    "name": "美崙聯安診所",
    "address": "花蓮縣花蓮市正義街8-2號",
    "lat": 23.983383,
    "lng": 121.606317,
    "estimated": false,
    "isPrimary488": true,
    "group": "主要488名單",
    "delta": -0.0011101923076921594,
    "subGroup": "主要名單-西區",
    "region": "界線以西",
    "color": "紅色",
    "colorHex": "#e53e3e"
  },
  {
    "id": "store-249",
    "name": "林子芹營業所",
    "address": "花蓮縣花蓮市民國路103之2號",
    "lat": 23.978665,
    "lng": 121.605184,
    "estimated": false,
    "isPrimary488": false,
    "group": "非488名單(另色標記)",
    "subGroup": "非488名單",
    "region": "非488店家",
    "color": "黃色",
    "colorHex": "#f59e0b"
  },
  {
    "id": "store-250",
    "name": "台灣菸酒公司花蓮營業所",
    "address": "花蓮縣花蓮市民國路16-1號",
    "lat": 23.977805,
    "lng": 121.60424,
    "estimated": false,
    "isPrimary488": true,
    "group": "主要488名單",
    "delta": 0.0030344230769259184,
    "subGroup": "主要名單-東區",
    "region": "界線以東",
    "color": "綠色",
    "colorHex": "#257a4e"
  },
  {
    "id": "store-251",
    "name": "國泰世華銀行花蓮分行",
    "address": "花蓮縣花蓮市民國路163號",
    "lat": 23.980642,
    "lng": 121.608266,
    "estimated": false,
    "isPrimary488": true,
    "group": "主要488名單",
    "delta": 0.003896076923084024,
    "subGroup": "主要名單-東區",
    "region": "界線以東",
    "color": "綠色",
    "colorHex": "#257a4e"
  },
  {
    "id": "store-252",
    "name": "國泰世華銀行花蓮分公司倉庫",
    "address": "花蓮縣花蓮市民國路163號3樓-1",
    "lat": 23.980642,
    "lng": 121.608266,
    "estimated": false,
    "isPrimary488": true,
    "group": "主要488名單",
    "delta": 0.003896076923084024,
    "subGroup": "主要名單-東區",
    "region": "界線以東",
    "color": "綠色",
    "colorHex": "#257a4e"
  },
  {
    "id": "store-253",
    "name": "新光銀行ATM元富證券反偵測",
    "address": "花蓮縣花蓮市民國路167號",
    "lat": 23.980949,
    "lng": 121.60868,
    "estimated": false,
    "isPrimary488": true,
    "group": "主要488名單",
    "delta": 0.0039676538461606015,
    "subGroup": "主要名單-東區",
    "region": "界線以東",
    "color": "綠色",
    "colorHex": "#257a4e"
  },
  {
    "id": "store-254",
    "name": "新光銀行ATM元富證券收付處",
    "address": "花蓮縣花蓮市民國路167號",
    "lat": 23.980949,
    "lng": 121.60868,
    "estimated": false,
    "isPrimary488": true,
    "group": "主要488名單",
    "delta": 0.0039676538461606015,
    "subGroup": "主要名單-東區",
    "region": "界線以東",
    "color": "綠色",
    "colorHex": "#257a4e"
  },
  {
    "id": "store-255",
    "name": "元富證券花蓮分公司",
    "address": "花蓮縣花蓮市民國路167號1-2號",
    "lat": 23.980949,
    "lng": 121.60868,
    "estimated": false,
    "isPrimary488": true,
    "group": "主要488名單",
    "delta": 0.0039676538461606015,
    "subGroup": "主要名單-東區",
    "region": "界線以東",
    "color": "綠色",
    "colorHex": "#257a4e"
  },
  {
    "id": "store-256",
    "name": "劉玉萍營業所",
    "address": "花蓮縣花蓮市民國路2-1號",
    "lat": 23.976336,
    "lng": 121.601983,
    "estimated": false,
    "isPrimary488": true,
    "group": "主要488名單",
    "delta": 0.0024159230769242868,
    "subGroup": "主要名單-東區",
    "region": "界線以東",
    "color": "綠色",
    "colorHex": "#257a4e"
  },
  {
    "id": "store-257",
    "name": "吳欣潔信欣中醫",
    "address": "花蓮縣花蓮市民國路26號",
    "lat": 23.97841,
    "lng": 121.60516,
    "estimated": false,
    "isPrimary488": true,
    "group": "主要488名單",
    "delta": 0.003279615384613521,
    "subGroup": "主要名單-東區",
    "region": "界線以東",
    "color": "綠色",
    "colorHex": "#257a4e"
  },
  {
    "id": "store-258",
    "name": "雄瑋饗樂有限公司",
    "address": "花蓮縣花蓮市民國路32-3號1樓",
    "lat": 23.978681,
    "lng": 121.60561,
    "estimated": false,
    "isPrimary488": true,
    "group": "主要488名單",
    "delta": 0.0034273461538418815,
    "subGroup": "主要名單-東區",
    "region": "界線以東",
    "color": "綠色",
    "colorHex": "#257a4e"
  },
  {
    "id": "store-259",
    "name": "頂興食品事業有限公司",
    "address": "花蓮縣花蓮市民國路69號1-2樓",
    "lat": 23.977432,
    "lng": 121.60333,
    "estimated": false,
    "isPrimary488": true,
    "group": "主要488名單",
    "delta": 0.0025404615384587714,
    "subGroup": "主要名單-東區",
    "region": "界線以東",
    "color": "綠色",
    "colorHex": "#257a4e"
  },
  {
    "id": "store-260",
    "name": "波提娜麗",
    "address": "花蓮縣花蓮市民國路80巷18號",
    "lat": 23.980407,
    "lng": 121.609473,
    "estimated": false,
    "isPrimary488": true,
    "group": "主要488名單",
    "delta": 0.005365192307692723,
    "subGroup": "主要名單-東區",
    "region": "界線以東",
    "color": "綠色",
    "colorHex": "#257a4e"
  },
  {
    "id": "store-261",
    "name": "徐名宏公館",
    "address": "花蓮縣花蓮市光復街126號",
    "lat": 23.98046,
    "lng": 121.604981,
    "estimated": false,
    "isPrimary488": true,
    "group": "主要488名單",
    "delta": 0.0008140769230777778,
    "subGroup": "主要名單-東區",
    "region": "界線以東",
    "color": "綠色",
    "colorHex": "#257a4e"
  },
  {
    "id": "store-262",
    "name": "曾鈺倫營業所",
    "address": "花蓮縣花蓮市光復街128號",
    "lat": 23.980508,
    "lng": 121.604823,
    "estimated": false,
    "isPrimary488": true,
    "group": "主要488名單",
    "delta": 0.0006025384615355733,
    "subGroup": "主要名單-東區",
    "region": "界線以東",
    "color": "綠色",
    "colorHex": "#257a4e"
  },
  {
    "id": "store-263",
    "name": "花蓮二信總社",
    "address": "花蓮縣花蓮市光復街36號1-2樓",
    "lat": 23.976366,
    "lng": 121.608011,
    "estimated": false,
    "isPrimary488": true,
    "group": "主要488名單",
    "delta": 0.008410461538460368,
    "subGroup": "主要名單-東區",
    "region": "界線以東",
    "color": "綠色",
    "colorHex": "#257a4e"
  },
  {
    "id": "store-264",
    "name": "艾萊實業",
    "address": "花蓮縣花蓮市光復街45號",
    "lat": 23.976513,
    "lng": 121.607734,
    "estimated": false,
    "isPrimary488": true,
    "group": "主要488名單",
    "delta": 0.007969499999987306,
    "subGroup": "主要名單-東區",
    "region": "界線以東",
    "color": "綠色",
    "colorHex": "#257a4e"
  },
  {
    "id": "store-265",
    "name": "精誠輪業有限公司",
    "address": "花蓮縣花蓮市光復街69號1-2樓",
    "lat": 23.978422,
    "lng": 121.606263,
    "estimated": false,
    "isPrimary488": true,
    "group": "主要488名單",
    "delta": 0.004369230769228238,
    "subGroup": "主要名單-東區",
    "region": "界線以東",
    "color": "綠色",
    "colorHex": "#257a4e"
  },
  {
    "id": "store-266",
    "name": "林勤龍娃娃機光復街店",
    "address": "花蓮縣花蓮市光復路60號",
    "lat": 24.024363,
    "lng": 121.608374,
    "estimated": false,
    "isPrimary488": true,
    "group": "主要488名單",
    "delta": -0.044761653846137506,
    "subGroup": "主要名單-西區",
    "region": "界線以西",
    "color": "紅色",
    "colorHex": "#e53e3e"
  },
  {
    "id": "store-267",
    "name": "宇泰汽車修理廠",
    "address": "花蓮縣花蓮市吉安鄉自強路438號",
    "lat": 23.984312,
    "lng": 121.590228,
    "estimated": false,
    "isPrimary488": true,
    "group": "主要488名單",
    "delta": -0.018235384615380212,
    "subGroup": "主要名單-西區",
    "region": "界線以西",
    "color": "紅色",
    "colorHex": "#e53e3e"
  },
  {
    "id": "store-268",
    "name": "郭香韻公館",
    "address": "花蓮縣花蓮市成功街225號",
    "lat": 23.974556,
    "lng": 121.607705,
    "estimated": false,
    "isPrimary488": false,
    "group": "非488名單(另色標記)",
    "subGroup": "非488名單",
    "region": "非488店家",
    "color": "黃色",
    "colorHex": "#f59e0b"
  },
  {
    "id": "store-269",
    "name": "台灣花蓮地方法院 院長宿舍",
    "address": "花蓮縣花蓮市自由街152號",
    "lat": 23.979811,
    "lng": 121.603517,
    "estimated": false,
    "isPrimary488": true,
    "group": "主要488名單",
    "delta": 0.00007396153846173092,
    "subGroup": "主要名單-西區",
    "region": "界線以西",
    "color": "紅色",
    "colorHex": "#e53e3e"
  },
  {
    "id": "store-270",
    "name": "花蓮一信自由分社",
    "address": "花蓮縣花蓮市自由街72號",
    "lat": 23.975298,
    "lng": 121.607034,
    "estimated": false,
    "isPrimary488": true,
    "group": "主要488名單",
    "delta": 0.008624692307691362,
    "subGroup": "主要名單-東區",
    "region": "界線以東",
    "color": "綠色",
    "colorHex": "#257a4e"
  },
  {
    "id": "store-271",
    "name": "楊淑芬公館",
    "address": "花蓮縣花蓮市自立路148巷70號",
    "lat": 23.987933,
    "lng": 121.596418,
    "estimated": false,
    "isPrimary488": false,
    "group": "非488名單(另色標記)",
    "subGroup": "非488名單",
    "region": "非488店家",
    "color": "黃色",
    "colorHex": "#f59e0b"
  },
  {
    "id": "store-272",
    "name": "臺富國際食品股份有限公司",
    "address": "花蓮縣花蓮市和平路1段110號",
    "lat": 23.967317,
    "lng": 121.60826,
    "estimated": false,
    "isPrimary488": true,
    "group": "主要488名單",
    "delta": 0.018752576923077413,
    "subGroup": "主要名單-東區",
    "region": "界線以東",
    "color": "綠色",
    "colorHex": "#257a4e"
  },
  {
    "id": "store-273",
    "name": "寶",
    "address": "花蓮縣花蓮市和平路285號",
    "lat": 23.970279,
    "lng": 121.606224,
    "estimated": false,
    "isPrimary488": true,
    "group": "主要488名單",
    "delta": 0.013412807692304796,
    "subGroup": "主要名單-東區",
    "region": "界線以東",
    "color": "綠色",
    "colorHex": "#257a4e"
  },
  {
    "id": "store-274",
    "name": "鑫饕寶炭烤鍋物",
    "address": "花蓮縣花蓮市和平路302-1號",
    "lat": 23.97279,
    "lng": 121.604576,
    "estimated": false,
    "isPrimary488": false,
    "group": "非488名單(另色標記)",
    "subGroup": "非488名單",
    "region": "非488店家",
    "color": "黃色",
    "colorHex": "#f59e0b"
  },
  {
    "id": "store-275",
    "name": "朱美雪鍋物",
    "address": "花蓮縣花蓮市和平路307號",
    "lat": 23.972753,
    "lng": 121.604285,
    "estimated": false,
    "isPrimary488": false,
    "group": "非488名單(另色標記)",
    "subGroup": "非488名單",
    "region": "非488店家",
    "color": "黃色",
    "colorHex": "#f59e0b"
  },
  {
    "id": "store-276",
    "name": "原相中醫診所",
    "address": "花蓮縣花蓮市和平路332-2號",
    "lat": 23.973561,
    "lng": 121.603972,
    "estimated": false,
    "isPrimary488": true,
    "group": "主要488名單",
    "delta": 0.007500115384615924,
    "subGroup": "主要名單-東區",
    "region": "界線以東",
    "color": "綠色",
    "colorHex": "#257a4e"
  },
  {
    "id": "store-277",
    "name": "朱美雪營業所無人自助洗",
    "address": "花蓮縣花蓮市和平路362號",
    "lat": 23.974115,
    "lng": 121.603517,
    "estimated": false,
    "isPrimary488": true,
    "group": "主要488名單",
    "delta": 0.006427192307683072,
    "subGroup": "主要名單-東區",
    "region": "界線以東",
    "color": "綠色",
    "colorHex": "#257a4e"
  },
  {
    "id": "store-278",
    "name": "冠潓商行",
    "address": "花蓮縣花蓮市和平路403號1樓",
    "lat": 23.974809,
    "lng": 121.602591,
    "estimated": false,
    "isPrimary488": true,
    "group": "主要488名單",
    "delta": 0.004727115384611125,
    "subGroup": "主要名單-東區",
    "region": "界線以東",
    "color": "綠色",
    "colorHex": "#257a4e"
  },
  {
    "id": "store-279",
    "name": "大同綜合訊電和平門市",
    "address": "花蓮縣花蓮市和平路415號(417號)",
    "lat": 23.974991,
    "lng": 121.602409,
    "estimated": false,
    "isPrimary488": false,
    "group": "非488名單(另色標記)",
    "subGroup": "非488名單",
    "region": "非488店家",
    "color": "黃色",
    "colorHex": "#f59e0b"
  },
  {
    "id": "store-280",
    "name": "泰多車行",
    "address": "花蓮縣花蓮市和平路511號",
    "lat": 23.976779,
    "lng": 121.601095,
    "estimated": false,
    "isPrimary488": true,
    "group": "主要488名單",
    "delta": 0.0010338076923090966,
    "subGroup": "主要名單-東區",
    "region": "界線以東",
    "color": "綠色",
    "colorHex": "#257a4e"
  },
  {
    "id": "store-281",
    "name": "新光銀行入金機花蓮大遠百",
    "address": "花蓮縣花蓮市和平路581號",
    "lat": 23.978381,
    "lng": 121.599729,
    "estimated": false,
    "isPrimary488": true,
    "group": "主要488名單",
    "delta": -0.0021190384615437097,
    "subGroup": "主要名單-西區",
    "region": "界線以西",
    "color": "紅色",
    "colorHex": "#e53e3e"
  },
  {
    "id": "store-282",
    "name": "新光銀行ATM花蓮遠東百貨",
    "address": "花蓮縣花蓮市和平路581號",
    "lat": 23.978381,
    "lng": 121.599729,
    "estimated": false,
    "isPrimary488": true,
    "group": "主要488名單",
    "delta": -0.0021190384615437097,
    "subGroup": "主要名單-西區",
    "region": "界線以西",
    "color": "紅色",
    "colorHex": "#e53e3e"
  },
  {
    "id": "store-283",
    "name": "新光銀行入金機花蓮遠東百貨",
    "address": "花蓮縣花蓮市和平路581號B1",
    "lat": 23.978381,
    "lng": 121.599729,
    "estimated": false,
    "isPrimary488": true,
    "group": "主要488名單",
    "delta": -0.0021190384615437097,
    "subGroup": "主要名單-西區",
    "region": "界線以西",
    "color": "紅色",
    "colorHex": "#e53e3e"
  },
  {
    "id": "store-284",
    "name": "台新銀行ATM愛買花蓮門市",
    "address": "花蓮縣花蓮市和平路581號B1",
    "lat": 23.978381,
    "lng": 121.599729,
    "estimated": false,
    "isPrimary488": true,
    "group": "主要488名單",
    "delta": -0.0021190384615437097,
    "subGroup": "主要名單-西區",
    "region": "界線以西",
    "color": "紅色",
    "colorHex": "#e53e3e"
  },
  {
    "id": "store-285",
    "name": "動吃動循環工作室",
    "address": "花蓮縣花蓮市和平路79號",
    "lat": 23.968829,
    "lng": 121.607343,
    "estimated": false,
    "isPrimary488": false,
    "group": "非488名單(另色標記)",
    "subGroup": "非488名單",
    "region": "非488店家",
    "color": "黃色",
    "colorHex": "#f59e0b"
  },
  {
    "id": "store-286",
    "name": "映達企業社",
    "address": "花蓮縣花蓮市和平路92-1號",
    "lat": 23.968645,
    "lng": 121.607744,
    "estimated": false,
    "isPrimary488": true,
    "group": "主要488名單",
    "delta": 0.01675534615384322,
    "subGroup": "主要名單-東區",
    "region": "界線以東",
    "color": "綠色",
    "colorHex": "#257a4e"
  },
  {
    "id": "store-287",
    "name": "崧煌工程行張哲維公館",
    "address": "花蓮縣花蓮市延平街210號",
    "lat": 23.985962,
    "lng": 121.597787,
    "estimated": false,
    "isPrimary488": false,
    "group": "非488名單(另色標記)",
    "subGroup": "非488名單",
    "region": "非488店家",
    "color": "黃色",
    "colorHex": "#f59e0b"
  },
  {
    "id": "store-288",
    "name": "豪謙國際有限公司丼物園",
    "address": "花蓮縣花蓮市延平路148巷22號",
    "lat": 24.023173,
    "lng": 121.60471,
    "estimated": false,
    "isPrimary488": true,
    "group": "主要488名單",
    "delta": -0.04709834615383102,
    "subGroup": "主要名單-西區",
    "region": "界線以西",
    "color": "紅色",
    "colorHex": "#e53e3e"
  },
  {
    "id": "store-289",
    "name": "重慶南北雜貨店",
    "address": "花蓮縣花蓮市忠孝街2號旁",
    "lat": 23.972152,
    "lng": 121.608191,
    "estimated": true,
    "isPrimary488": false,
    "group": "非488名單(另色標記)",
    "subGroup": "非488名單",
    "region": "非488店家",
    "color": "黃色",
    "colorHex": "#f59e0b"
  },
  {
    "id": "store-290",
    "name": "晶新企業社",
    "address": "花蓮縣花蓮市忠孝街60-1號1樓",
    "lat": 23.974201,
    "lng": 121.60641,
    "estimated": false,
    "isPrimary488": true,
    "group": "主要488名單",
    "delta": 0.009224269230770687,
    "subGroup": "主要名單-東區",
    "region": "界線以東",
    "color": "綠色",
    "colorHex": "#257a4e"
  },
  {
    "id": "store-291",
    "name": "華慶國際投資有限公司",
    "address": "花蓮縣花蓮市明心街1-16號1樓",
    "lat": 23.985338,
    "lng": 121.606603,
    "estimated": false,
    "isPrimary488": true,
    "group": "主要488名單",
    "delta": -0.00300476923075621,
    "subGroup": "主要名單-西區",
    "region": "界線以西",
    "color": "紅色",
    "colorHex": "#e53e3e"
  },
  {
    "id": "store-292",
    "name": "明正不動產有限公司",
    "address": "花蓮縣花蓮市明心街1-28號",
    "lat": 23.985338,
    "lng": 121.606603,
    "estimated": false,
    "isPrimary488": false,
    "group": "非488名單(另色標記)",
    "subGroup": "非488名單",
    "region": "非488店家",
    "color": "黃色",
    "colorHex": "#f59e0b"
  },
  {
    "id": "store-293",
    "name": "台新銀行ATM花蓮城中OK超商",
    "address": "花蓮縣花蓮市明心街1之18號",
    "lat": 23.985168,
    "lng": 121.606207,
    "estimated": false,
    "isPrimary488": false,
    "group": "非488名單(另色標記)",
    "subGroup": "非488名單",
    "region": "非488店家",
    "color": "黃色",
    "colorHex": "#f59e0b"
  },
  {
    "id": "store-294",
    "name": "綺麗實業社",
    "address": "花蓮縣花蓮市明心街2之1號",
    "lat": 23.985338,
    "lng": 121.606603,
    "estimated": false,
    "isPrimary488": true,
    "group": "主要488名單",
    "delta": -0.00300476923075621,
    "subGroup": "主要名單-西區",
    "region": "界線以西",
    "color": "紅色",
    "colorHex": "#e53e3e"
  },
  {
    "id": "store-295",
    "name": "劉國鈞公館",
    "address": "花蓮縣花蓮市明心街47號",
    "lat": 23.986292,
    "lng": 121.6081,
    "estimated": false,
    "isPrimary488": true,
    "group": "主要488名單",
    "delta": -0.0025718461538559723,
    "subGroup": "主要名單-西區",
    "region": "界線以西",
    "color": "紅色",
    "colorHex": "#e53e3e"
  },
  {
    "id": "store-296",
    "name": "祐薇企業社",
    "address": "花蓮縣花蓮市明智街101號",
    "lat": 23.985904,
    "lng": 121.605581,
    "estimated": false,
    "isPrimary488": true,
    "group": "主要488名單",
    "delta": -0.0046580769230786245,
    "subGroup": "主要名單-西區",
    "region": "界線以西",
    "color": "紅色",
    "colorHex": "#e53e3e"
  },
  {
    "id": "store-297",
    "name": "游時銘公館",
    "address": "花蓮縣花蓮市明智街108號",
    "lat": 23.985623,
    "lng": 121.605234,
    "estimated": false,
    "isPrimary488": true,
    "group": "主要488名單",
    "delta": -0.004691653846151667,
    "subGroup": "主要名單-西區",
    "region": "界線以西",
    "color": "紅色",
    "colorHex": "#e53e3e"
  },
  {
    "id": "store-298",
    "name": "花蔬生鮮超市供應中心",
    "address": "花蓮縣花蓮市明智街13號",
    "lat": 23.984103,
    "lng": 121.602809,
    "estimated": false,
    "isPrimary488": false,
    "group": "非488名單(另色標記)",
    "subGroup": "非488名單",
    "region": "非488店家",
    "color": "黃色",
    "colorHex": "#f59e0b"
  },
  {
    "id": "store-299",
    "name": "花蓮二信ATM明義國小",
    "address": "花蓮縣花蓮市明義街107號",
    "lat": 23.97051,
    "lng": 121.610205,
    "estimated": false,
    "isPrimary488": true,
    "group": "主要488名單",
    "delta": 0.017136153846138313,
    "subGroup": "主要名單-東區",
    "region": "界線以東",
    "color": "綠色",
    "colorHex": "#257a4e"
  },
  {
    "id": "store-300",
    "name": "明義國小-健康中心",
    "address": "花蓮縣花蓮市明義街107號",
    "lat": 23.97051,
    "lng": 121.610205,
    "estimated": false,
    "isPrimary488": true,
    "group": "主要488名單",
    "delta": 0.017136153846138313,
    "subGroup": "主要名單-東區",
    "region": "界線以東",
    "color": "綠色",
    "colorHex": "#257a4e"
  },
  {
    "id": "store-301",
    "name": "明義國小-學務處",
    "address": "花蓮縣花蓮市明義街107號",
    "lat": 23.97051,
    "lng": 121.610205,
    "estimated": false,
    "isPrimary488": true,
    "group": "主要488名單",
    "delta": 0.017136153846138313,
    "subGroup": "主要名單-東區",
    "region": "界線以東",
    "color": "綠色",
    "colorHex": "#257a4e"
  },
  {
    "id": "store-302",
    "name": "明義國小總務處",
    "address": "花蓮縣花蓮市明義街107號",
    "lat": 23.97051,
    "lng": 121.610205,
    "estimated": false,
    "isPrimary488": true,
    "group": "主要488名單",
    "delta": 0.017136153846138313,
    "subGroup": "主要名單-東區",
    "region": "界線以東",
    "color": "綠色",
    "colorHex": "#257a4e"
  },
  {
    "id": "store-303",
    "name": "明義國小-教務處",
    "address": "花蓮縣花蓮市明義街107號",
    "lat": 23.97051,
    "lng": 121.610205,
    "estimated": false,
    "isPrimary488": true,
    "group": "主要488名單",
    "delta": 0.017136153846138313,
    "subGroup": "主要名單-東區",
    "region": "界線以東",
    "color": "綠色",
    "colorHex": "#257a4e"
  },
  {
    "id": "store-304",
    "name": "明義國小-圖書館",
    "address": "花蓮縣花蓮市明義街107號",
    "lat": 23.97051,
    "lng": 121.610205,
    "estimated": false,
    "isPrimary488": true,
    "group": "主要488名單",
    "delta": 0.017136153846138313,
    "subGroup": "主要名單-東區",
    "region": "界線以東",
    "color": "綠色",
    "colorHex": "#257a4e"
  },
  {
    "id": "store-305",
    "name": "明義國小舊總務處",
    "address": "花蓮縣花蓮市明義街107號",
    "lat": 23.97051,
    "lng": 121.610205,
    "estimated": false,
    "isPrimary488": true,
    "group": "主要488名單",
    "delta": 0.017136153846138313,
    "subGroup": "主要名單-東區",
    "region": "界線以東",
    "color": "綠色",
    "colorHex": "#257a4e"
  },
  {
    "id": "store-306",
    "name": "明義國小-電腦教室",
    "address": "花蓮縣花蓮市明義街107號",
    "lat": 23.97051,
    "lng": 121.610205,
    "estimated": false,
    "isPrimary488": true,
    "group": "主要488名單",
    "delta": 0.017136153846138313,
    "subGroup": "主要名單-東區",
    "region": "界線以東",
    "color": "綠色",
    "colorHex": "#257a4e"
  },
  {
    "id": "store-307",
    "name": "明義國小-資訊研究室",
    "address": "花蓮縣花蓮市明義街107號",
    "lat": 23.97051,
    "lng": 121.610205,
    "estimated": false,
    "isPrimary488": true,
    "group": "主要488名單",
    "delta": 0.017136153846138313,
    "subGroup": "主要名單-東區",
    "region": "界線以東",
    "color": "綠色",
    "colorHex": "#257a4e"
  },
  {
    "id": "store-308",
    "name": "明義國小民國樓",
    "address": "花蓮縣花蓮市明義街107號",
    "lat": 23.97051,
    "lng": 121.610205,
    "estimated": false,
    "isPrimary488": true,
    "group": "主要488名單",
    "delta": 0.017136153846138313,
    "subGroup": "主要名單-東區",
    "region": "界線以東",
    "color": "綠色",
    "colorHex": "#257a4e"
  },
  {
    "id": "store-309",
    "name": "明義國小演藝廳",
    "address": "花蓮縣花蓮市明義街107號",
    "lat": 23.97051,
    "lng": 121.610205,
    "estimated": false,
    "isPrimary488": true,
    "group": "主要488名單",
    "delta": 0.017136153846138313,
    "subGroup": "主要名單-東區",
    "region": "界線以東",
    "color": "綠色",
    "colorHex": "#257a4e"
  },
  {
    "id": "store-310",
    "name": "明義國小-音樂教室3樓",
    "address": "花蓮縣花蓮市明義街107號",
    "lat": 23.97051,
    "lng": 121.610205,
    "estimated": false,
    "isPrimary488": true,
    "group": "主要488名單",
    "delta": 0.017136153846138313,
    "subGroup": "主要名單-東區",
    "region": "界線以東",
    "color": "綠色",
    "colorHex": "#257a4e"
  },
  {
    "id": "store-311",
    "name": "明義國小藝才樓1樓",
    "address": "花蓮縣花蓮市明義街107號",
    "lat": 23.97051,
    "lng": 121.610205,
    "estimated": false,
    "isPrimary488": true,
    "group": "主要488名單",
    "delta": 0.017136153846138313,
    "subGroup": "主要名單-東區",
    "region": "界線以東",
    "color": "綠色",
    "colorHex": "#257a4e"
  },
  {
    "id": "store-312",
    "name": "明義國小-藝才班4樓",
    "address": "花蓮縣花蓮市明義街107號",
    "lat": 23.97051,
    "lng": 121.610205,
    "estimated": false,
    "isPrimary488": true,
    "group": "主要488名單",
    "delta": 0.017136153846138313,
    "subGroup": "主要名單-東區",
    "region": "界線以東",
    "color": "綠色",
    "colorHex": "#257a4e"
  },
  {
    "id": "store-313",
    "name": "明義國小仁愛樓",
    "address": "花蓮縣花蓮市明義街107號",
    "lat": 23.97051,
    "lng": 121.610205,
    "estimated": false,
    "isPrimary488": true,
    "group": "主要488名單",
    "delta": 0.017136153846138313,
    "subGroup": "主要名單-東區",
    "region": "界線以東",
    "color": "綠色",
    "colorHex": "#257a4e"
  },
  {
    "id": "store-314",
    "name": "明義國小林森樓",
    "address": "花蓮縣花蓮市明義街107號",
    "lat": 23.97051,
    "lng": 121.610205,
    "estimated": false,
    "isPrimary488": true,
    "group": "主要488名單",
    "delta": 0.017136153846138313,
    "subGroup": "主要名單-東區",
    "region": "界線以東",
    "color": "綠色",
    "colorHex": "#257a4e"
  },
  {
    "id": "store-315",
    "name": "陳子軒-無限可能國際有限公司",
    "address": "花蓮縣花蓮市明義街21號",
    "lat": 23.971291,
    "lng": 121.609568,
    "estimated": false,
    "isPrimary488": true,
    "group": "主要488名單",
    "delta": 0.015628038461528604,
    "subGroup": "主要名單-東區",
    "region": "界線以東",
    "color": "綠色",
    "colorHex": "#257a4e"
  },
  {
    "id": "store-316",
    "name": "Mia藝術美甲",
    "address": "花蓮縣花蓮市明義街61-1號",
    "lat": 23.971291,
    "lng": 121.609568,
    "estimated": false,
    "isPrimary488": false,
    "group": "非488名單(另色標記)",
    "subGroup": "非488名單",
    "region": "非488店家",
    "color": "黃色",
    "colorHex": "#f59e0b"
  },
  {
    "id": "store-317",
    "name": "林郡淵-營業所",
    "address": "花蓮縣花蓮市明禮路109號",
    "lat": 23.983729,
    "lng": 121.607258,
    "estimated": false,
    "isPrimary488": false,
    "group": "非488名單(另色標記)",
    "subGroup": "非488名單",
    "region": "非488店家",
    "color": "黃色",
    "colorHex": "#f59e0b"
  },
  {
    "id": "store-318",
    "name": "款款廚藝有限公司",
    "address": "花蓮縣花蓮市明禮路22-2號",
    "lat": 23.982615,
    "lng": 121.608368,
    "estimated": false,
    "isPrimary488": false,
    "group": "非488名單(另色標記)",
    "subGroup": "非488名單",
    "region": "非488店家",
    "color": "黃色",
    "colorHex": "#f59e0b"
  },
  {
    "id": "store-319",
    "name": "李淑惠公館",
    "address": "花蓮縣花蓮市明禮路67巷6號",
    "lat": 23.981414,
    "lng": 121.608776,
    "estimated": false,
    "isPrimary488": true,
    "group": "主要488名單",
    "delta": 0.003545000000002574,
    "subGroup": "主要名單-東區",
    "region": "界線以東",
    "color": "綠色",
    "colorHex": "#257a4e"
  },
  {
    "id": "store-320",
    "name": "許漢中營業所",
    "address": "花蓮縣花蓮市東大門夜市各省一條街G42",
    "lat": 23.974012,
    "lng": 121.61352,
    "estimated": true,
    "isPrimary488": true,
    "group": "主要488名單",
    "delta": 0.01654507692306595,
    "subGroup": "主要名單-東區",
    "region": "界線以東",
    "color": "綠色",
    "colorHex": "#257a4e"
  },
  {
    "id": "store-321",
    "name": "謝金玲公館",
    "address": "花蓮縣花蓮市林政街19巷1弄9號",
    "lat": 23.97995,
    "lng": 121.59639,
    "estimated": false,
    "isPrimary488": true,
    "group": "主要488名單",
    "delta": -0.007208076923078011,
    "subGroup": "主要名單-西區",
    "region": "界線以西",
    "color": "紅色",
    "colorHex": "#e53e3e"
  },
  {
    "id": "store-322",
    "name": "張群公館",
    "address": "花蓮縣花蓮市林政街22巷1號",
    "lat": 23.978146,
    "lng": 121.5962,
    "estimated": false,
    "isPrimary488": true,
    "group": "主要488名單",
    "delta": -0.005385923076929089,
    "subGroup": "主要名單-西區",
    "region": "界線以西",
    "color": "紅色",
    "colorHex": "#e53e3e"
  },
  {
    "id": "store-323",
    "name": "賴志明公館",
    "address": "花蓮縣花蓮市林政街33巷13號",
    "lat": 23.979748,
    "lng": 121.596099,
    "estimated": false,
    "isPrimary488": true,
    "group": "主要488名單",
    "delta": -0.007273769230778271,
    "subGroup": "主要名單-西區",
    "region": "界線以西",
    "color": "紅色",
    "colorHex": "#e53e3e"
  },
  {
    "id": "store-324",
    "name": "國風國中",
    "address": "花蓮縣花蓮市林政街7號",
    "lat": 23.978479,
    "lng": 121.595377,
    "estimated": false,
    "isPrimary488": true,
    "group": "主要488名單",
    "delta": -0.006580346153853611,
    "subGroup": "主要名單-西區",
    "region": "界線以西",
    "color": "紅色",
    "colorHex": "#e53e3e"
  },
  {
    "id": "store-325",
    "name": "林務局花蓮林區管理處女生宿舍",
    "address": "花蓮縣花蓮市林政街8巷5弄8.9號",
    "lat": 23.976306,
    "lng": 121.595273,
    "estimated": false,
    "isPrimary488": true,
    "group": "主要488名單",
    "delta": -0.004260615384609423,
    "subGroup": "主要名單-西區",
    "region": "界線以西",
    "color": "紅色",
    "colorHex": "#e53e3e"
  },
  {
    "id": "store-326",
    "name": "台灣大哥大花蓮林森直營店",
    "address": "花蓮縣花蓮市林森路157號",
    "lat": 23.977807,
    "lng": 121.599716,
    "estimated": false,
    "isPrimary488": false,
    "group": "非488名單(另色標記)",
    "subGroup": "非488名單",
    "region": "非488店家",
    "color": "黃色",
    "colorHex": "#f59e0b"
  },
  {
    "id": "store-327",
    "name": "焱世紀寵物百貨商行",
    "address": "花蓮縣花蓮市林森路164號",
    "lat": 23.9782,
    "lng": 121.600766,
    "estimated": false,
    "isPrimary488": true,
    "group": "主要488名單",
    "delta": -0.0008801538461682412,
    "subGroup": "主要名單-西區",
    "region": "界線以西",
    "color": "紅色",
    "colorHex": "#e53e3e"
  },
  {
    "id": "store-328",
    "name": "曾琇瑛公館",
    "address": "花蓮縣花蓮市林森路167號",
    "lat": 23.977932,
    "lng": 121.599921,
    "estimated": false,
    "isPrimary488": false,
    "group": "非488名單(另色標記)",
    "subGroup": "非488名單",
    "region": "非488店家",
    "color": "黃色",
    "colorHex": "#f59e0b"
  },
  {
    "id": "store-329",
    "name": "特力屋花蓮店",
    "address": "花蓮縣花蓮市林森路176號",
    "lat": 23.978545,
    "lng": 121.601265,
    "estimated": false,
    "isPrimary488": true,
    "group": "主要488名單",
    "delta": -0.000765961538462534,
    "subGroup": "主要名單-西區",
    "region": "界線以西",
    "color": "紅色",
    "colorHex": "#e53e3e"
  },
  {
    "id": "store-330",
    "name": "安泰診所",
    "address": "花蓮縣花蓮市林森路195之1",
    "lat": 23.978335,
    "lng": 121.600561,
    "estimated": false,
    "isPrimary488": true,
    "group": "主要488名單",
    "delta": -0.0012357307692383301,
    "subGroup": "主要名單-西區",
    "region": "界線以西",
    "color": "紅色",
    "colorHex": "#e53e3e"
  },
  {
    "id": "store-331",
    "name": "伊蕾名店花蓮林森店",
    "address": "花蓮縣花蓮市林森路202號",
    "lat": 23.980297,
    "lng": 121.603962,
    "estimated": false,
    "isPrimary488": true,
    "group": "主要488名單",
    "delta": -0.000023115384621519297,
    "subGroup": "主要名單-西區",
    "region": "界線以西",
    "color": "紅色",
    "colorHex": "#e53e3e"
  },
  {
    "id": "store-332",
    "name": "永豐銀行花蓮分行",
    "address": "花蓮縣花蓮市林森路218-1號",
    "lat": 23.980906,
    "lng": 121.604911,
    "estimated": false,
    "isPrimary488": true,
    "group": "主要488名單",
    "delta": 0.0002466153846114594,
    "subGroup": "主要名單-東區",
    "region": "界線以東",
    "color": "綠色",
    "colorHex": "#257a4e"
  },
  {
    "id": "store-333",
    "name": "台灣知識庫股份有限公司",
    "address": "花蓮縣花蓮市林森路242號",
    "lat": 23.981723,
    "lng": 121.606136,
    "estimated": false,
    "isPrimary488": true,
    "group": "主要488名單",
    "delta": 0.0005603461538612464,
    "subGroup": "主要名單-東區",
    "region": "界線以東",
    "color": "綠色",
    "colorHex": "#257a4e"
  },
  {
    "id": "store-334",
    "name": "彭貴春營業所",
    "address": "花蓮縣花蓮市林森路243巷1號",
    "lat": 23.978972,
    "lng": 121.601117,
    "estimated": false,
    "isPrimary488": true,
    "group": "主要488名單",
    "delta": -0.0013902307692319482,
    "subGroup": "主要名單-西區",
    "region": "界線以西",
    "color": "紅色",
    "colorHex": "#e53e3e"
  },
  {
    "id": "store-335",
    "name": "霧川商行",
    "address": "花蓮縣花蓮市林森路273巷1弄26號",
    "lat": 23.980258,
    "lng": 121.601332,
    "estimated": false,
    "isPrimary488": true,
    "group": "主要488名單",
    "delta": -0.002609615384614017,
    "subGroup": "主要名單-西區",
    "region": "界線以西",
    "color": "紅色",
    "colorHex": "#e53e3e"
  },
  {
    "id": "store-336",
    "name": "台灣銀行ATM反偵測北花蓮分行",
    "address": "花蓮縣花蓮市林森路281號",
    "lat": 23.97962,
    "lng": 121.60224,
    "estimated": false,
    "isPrimary488": false,
    "group": "非488名單(另色標記)",
    "subGroup": "非488名單",
    "region": "非488店家",
    "color": "黃色",
    "colorHex": "#f59e0b"
  },
  {
    "id": "store-337",
    "name": "台灣銀行北花蓮分行",
    "address": "花蓮縣花蓮市林森路281號1-2樓",
    "lat": 23.97962,
    "lng": 121.60224,
    "estimated": false,
    "isPrimary488": true,
    "group": "主要488名單",
    "delta": -0.0009900000000016007,
    "subGroup": "主要名單-西區",
    "region": "界線以西",
    "color": "紅色",
    "colorHex": "#e53e3e"
  },
  {
    "id": "store-338",
    "name": "路邊烤肉wildbbq奇萊店",
    "address": "花蓮縣花蓮市林森路295號",
    "lat": 23.9798,
    "lng": 121.602764,
    "estimated": false,
    "isPrimary488": true,
    "group": "主要488名單",
    "delta": -0.0006667692307757989,
    "subGroup": "主要名單-西區",
    "region": "界線以西",
    "color": "紅色",
    "colorHex": "#e53e3e"
  },
  {
    "id": "store-339",
    "name": "大同綜合訊電花蓮林森展售中心",
    "address": "花蓮縣花蓮市林森路317號",
    "lat": 23.980363,
    "lng": 121.603645,
    "estimated": false,
    "isPrimary488": false,
    "group": "非488名單(另色標記)",
    "subGroup": "非488名單",
    "region": "非488店家",
    "color": "黃色",
    "colorHex": "#f59e0b"
  },
  {
    "id": "store-340",
    "name": "新鮮葉仔樂活休閒館",
    "address": "花蓮縣花蓮市林森路365.367號",
    "lat": 23.981377,
    "lng": 121.605208,
    "estimated": false,
    "isPrimary488": true,
    "group": "主要488名單",
    "delta": 0.000018269230778855672,
    "subGroup": "主要名單-西區",
    "region": "界線以西",
    "color": "紅色",
    "colorHex": "#e53e3e"
  },
  {
    "id": "store-341",
    "name": "臺灣花蓮地方法院所屬民間公證人",
    "address": "花蓮縣花蓮市林森路399號",
    "lat": 23.982327,
    "lng": 121.606619,
    "estimated": false,
    "isPrimary488": true,
    "group": "主要488名單",
    "delta": 0.0003696538461497312,
    "subGroup": "主要名單-東區",
    "region": "界線以東",
    "color": "綠色",
    "colorHex": "#257a4e"
  },
  {
    "id": "store-342",
    "name": "神田屋食品行",
    "address": "花蓮縣花蓮市林森路407號",
    "lat": 23.982531,
    "lng": 121.606853,
    "estimated": false,
    "isPrimary488": true,
    "group": "主要488名單",
    "delta": 0.00037611538461135297,
    "subGroup": "主要名單-東區",
    "region": "界線以東",
    "color": "綠色",
    "colorHex": "#257a4e"
  },
  {
    "id": "store-343",
    "name": "三井鞋坊",
    "address": "花蓮縣花蓮市林森路5號",
    "lat": 23.97508,
    "lng": 121.5962,
    "estimated": false,
    "isPrimary488": true,
    "group": "主要488名單",
    "delta": -0.001966153846154839,
    "subGroup": "主要名單-西區",
    "region": "界線以西",
    "color": "紅色",
    "colorHex": "#e53e3e"
  },
  {
    "id": "store-344",
    "name": "永旭室內裝修有限公司",
    "address": "花蓮縣花蓮市林森路77號",
    "lat": 23.976323,
    "lng": 121.597604,
    "estimated": false,
    "isPrimary488": false,
    "group": "非488名單(另色標記)",
    "subGroup": "非488名單",
    "region": "非488店家",
    "color": "黃色",
    "colorHex": "#f59e0b"
  },
  {
    "id": "store-345",
    "name": "賴桑海鮮有限公司",
    "address": "花蓮縣花蓮市林森路90號",
    "lat": 23.977303,
    "lng": 121.599416,
    "estimated": false,
    "isPrimary488": true,
    "group": "主要488名單",
    "delta": -0.0012296538461527007,
    "subGroup": "主要名單-西區",
    "region": "界線以西",
    "color": "紅色",
    "colorHex": "#e53e3e"
  },
  {
    "id": "store-346",
    "name": "賴桑海鮮有限公司",
    "address": "花蓮縣花蓮市林森路90號",
    "lat": 23.977303,
    "lng": 121.599416,
    "estimated": false,
    "isPrimary488": true,
    "group": "主要488名單",
    "delta": -0.0012296538461527007,
    "subGroup": "主要名單-西區",
    "region": "界線以西",
    "color": "紅色",
    "colorHex": "#e53e3e"
  },
  {
    "id": "store-347",
    "name": "吳吉成公館",
    "address": "花蓮縣花蓮市花蓮市國風街21巷21號",
    "lat": 23.976343,
    "lng": 121.596935,
    "estimated": false,
    "isPrimary488": true,
    "group": "主要488名單",
    "delta": -0.0026398846153767863,
    "subGroup": "主要名單-西區",
    "region": "界線以西",
    "color": "紅色",
    "colorHex": "#e53e3e"
  },
  {
    "id": "store-348",
    "name": "穆心儀公館",
    "address": "花蓮縣花蓮市信義街21巷2號",
    "lat": 23.973114,
    "lng": 121.604823,
    "estimated": false,
    "isPrimary488": false,
    "group": "非488名單(另色標記)",
    "subGroup": "非488名單",
    "region": "非488店家",
    "color": "黃色",
    "colorHex": "#f59e0b"
  },
  {
    "id": "store-349",
    "name": "勇鉅輪胎商行",
    "address": "花蓮縣花蓮市信義街32號",
    "lat": 23.973842,
    "lng": 121.604734,
    "estimated": false,
    "isPrimary488": false,
    "group": "非488名單(另色標記)",
    "subGroup": "非488名單",
    "region": "非488店家",
    "color": "黃色",
    "colorHex": "#f59e0b"
  },
  {
    "id": "store-350",
    "name": "黃微翔公館",
    "address": "花蓮縣花蓮市信義街47號",
    "lat": 23.974016,
    "lng": 121.604397,
    "estimated": false,
    "isPrimary488": true,
    "group": "主要488名單",
    "delta": 0.007417615384625265,
    "subGroup": "主要名單-東區",
    "region": "界線以東",
    "color": "綠色",
    "colorHex": "#257a4e"
  },
  {
    "id": "store-351",
    "name": "台電電源開發處東部水文管理站",
    "address": "花蓮縣花蓮市南京街125號",
    "lat": 23.973525,
    "lng": 121.604606,
    "estimated": false,
    "isPrimary488": false,
    "group": "非488名單(另色標記)",
    "subGroup": "非488名單",
    "region": "非488店家",
    "color": "黃色",
    "colorHex": "#f59e0b"
  },
  {
    "id": "store-352",
    "name": "陳威廷營業所",
    "address": "花蓮縣花蓮市南京街302號",
    "lat": 23.975294,
    "lng": 121.607554,
    "estimated": false,
    "isPrimary488": false,
    "group": "非488名單(另色標記)",
    "subGroup": "非488名單",
    "region": "非488店家",
    "color": "黃色",
    "colorHex": "#f59e0b"
  },
  {
    "id": "store-353",
    "name": "陳麗華營業所",
    "address": "花蓮縣花蓮市南京街38號",
    "lat": 23.971572,
    "lng": 121.601945,
    "estimated": false,
    "isPrimary488": true,
    "group": "主要488名單",
    "delta": 0.007691615384615602,
    "subGroup": "主要名單-東區",
    "region": "界線以東",
    "color": "綠色",
    "colorHex": "#257a4e"
  },
  {
    "id": "store-354",
    "name": "三傢一活蝦料理",
    "address": "花蓮縣花蓮市南濱32號",
    "lat": 23.982173,
    "lng": 121.606847,
    "estimated": false,
    "isPrimary488": true,
    "group": "主要488名單",
    "delta": 0.0007694230769317301,
    "subGroup": "主要名單-東區",
    "region": "界線以東",
    "color": "綠色",
    "colorHex": "#257a4e"
  },
  {
    "id": "store-355",
    "name": "亭停車花蓮中山營業所",
    "address": "花蓮縣花蓮市建中街68號對面空地停車場",
    "lat": 23.986592,
    "lng": 121.601643,
    "estimated": true,
    "isPrimary488": false,
    "group": "非488名單(另色標記)",
    "subGroup": "非488名單",
    "region": "非488店家",
    "color": "黃色",
    "colorHex": "#f59e0b"
  },
  {
    "id": "store-356",
    "name": "明義國小博愛分校廚房",
    "address": "花蓮縣花蓮市建林街163之1號",
    "lat": 23.982278,
    "lng": 121.60082,
    "estimated": false,
    "isPrimary488": true,
    "group": "主要488名單",
    "delta": -0.00537469230769716,
    "subGroup": "主要名單-西區",
    "region": "界線以西",
    "color": "紅色",
    "colorHex": "#e53e3e"
  },
  {
    "id": "store-357",
    "name": "明義國小博愛分校總務處",
    "address": "花蓮縣花蓮市建林街163之1號",
    "lat": 23.982278,
    "lng": 121.60082,
    "estimated": false,
    "isPrimary488": true,
    "group": "主要488名單",
    "delta": -0.00537469230769716,
    "subGroup": "主要名單-西區",
    "region": "界線以西",
    "color": "紅色",
    "colorHex": "#e53e3e"
  },
  {
    "id": "store-358",
    "name": "陳坤隆公館",
    "address": "花蓮縣花蓮市建林街29號",
    "lat": 23.978826,
    "lng": 121.598074,
    "estimated": false,
    "isPrimary488": true,
    "group": "主要488名單",
    "delta": -0.004270384615395528,
    "subGroup": "主要名單-西區",
    "region": "界線以西",
    "color": "紅色",
    "colorHex": "#e53e3e"
  },
  {
    "id": "store-359",
    "name": "莫非廚房",
    "address": "花蓮縣花蓮市建國路105號",
    "lat": 23.983481,
    "lng": 121.602711,
    "estimated": false,
    "isPrimary488": false,
    "group": "非488名單(另色標記)",
    "subGroup": "非488名單",
    "region": "非488店家",
    "color": "黃色",
    "colorHex": "#f59e0b"
  },
  {
    "id": "store-360",
    "name": "義發號",
    "address": "花蓮縣花蓮市建國路107巷25號",
    "lat": 23.982764,
    "lng": 121.602127,
    "estimated": false,
    "isPrimary488": true,
    "group": "主要488名單",
    "delta": -0.004609769230768279,
    "subGroup": "主要名單-西區",
    "region": "界線以西",
    "color": "紅色",
    "colorHex": "#e53e3e"
  },
  {
    "id": "store-361",
    "name": "力銓電車股份有限公司",
    "address": "花蓮縣花蓮市建國路150號",
    "lat": 23.98398,
    "lng": 121.602,
    "estimated": false,
    "isPrimary488": true,
    "group": "主要488名單",
    "delta": -0.006093076923065155,
    "subGroup": "主要名單-西區",
    "region": "界線以西",
    "color": "紅色",
    "colorHex": "#e53e3e"
  },
  {
    "id": "store-362",
    "name": "花蓮二信ATM花蓮農校",
    "address": "花蓮縣花蓮市建國路161號",
    "lat": 23.983582,
    "lng": 121.598963,
    "estimated": false,
    "isPrimary488": true,
    "group": "主要488名單",
    "delta": -0.00868615384615623,
    "subGroup": "主要名單-西區",
    "region": "界線以西",
    "color": "紅色",
    "colorHex": "#e53e3e"
  },
  {
    "id": "store-363",
    "name": "萬達寵物花蓮建國店",
    "address": "花蓮縣花蓮市建國路168號",
    "lat": 23.983975,
    "lng": 121.601632,
    "estimated": false,
    "isPrimary488": false,
    "group": "非488名單(另色標記)",
    "subGroup": "非488名單",
    "region": "非488店家",
    "color": "黃色",
    "colorHex": "#f59e0b"
  },
  {
    "id": "store-364",
    "name": "遠東空調家電有限公司",
    "address": "花蓮縣花蓮市建國路1段103.105號1-B1樓",
    "lat": 23.985293,
    "lng": 121.593369,
    "estimated": false,
    "isPrimary488": true,
    "group": "主要488名單",
    "delta": -0.016188576923084952,
    "subGroup": "主要名單-西區",
    "region": "界線以西",
    "color": "紅色",
    "colorHex": "#e53e3e"
  },
  {
    "id": "store-365",
    "name": "遠東空調花蓮二店",
    "address": "花蓮縣花蓮市建國路1段111號",
    "lat": 23.985327,
    "lng": 121.593193,
    "estimated": false,
    "isPrimary488": true,
    "group": "主要488名單",
    "delta": -0.0164024999999981,
    "subGroup": "主要名單-西區",
    "region": "界線以西",
    "color": "紅色",
    "colorHex": "#e53e3e"
  },
  {
    "id": "store-366",
    "name": "曾欣玉不動產有限公司",
    "address": "花蓮縣花蓮市建國路1段129號",
    "lat": 23.985532,
    "lng": 121.591782,
    "estimated": false,
    "isPrimary488": false,
    "group": "非488名單(另色標記)",
    "subGroup": "非488名單",
    "region": "非488店家",
    "color": "黃色",
    "colorHex": "#f59e0b"
  },
  {
    "id": "store-367",
    "name": "禾太食品有限公司",
    "address": "花蓮縣花蓮市建國路1段53之3號",
    "lat": 23.984966,
    "lng": 121.594926,
    "estimated": false,
    "isPrimary488": false,
    "group": "非488名單(另色標記)",
    "subGroup": "非488名單",
    "region": "非488店家",
    "color": "黃色",
    "colorHex": "#f59e0b"
  },
  {
    "id": "store-368",
    "name": "酒棧商行",
    "address": "花蓮縣花蓮市建國路1段60號一樓",
    "lat": 23.985245,
    "lng": 121.594659,
    "estimated": false,
    "isPrimary488": true,
    "group": "主要488名單",
    "delta": -0.014845038461544391,
    "subGroup": "主要名單-西區",
    "region": "界線以西",
    "color": "紅色",
    "colorHex": "#e53e3e"
  },
  {
    "id": "store-369",
    "name": "興鈺科技有限公司",
    "address": "花蓮縣花蓮市建國路218號",
    "lat": 23.98417,
    "lng": 121.600554,
    "estimated": false,
    "isPrimary488": true,
    "group": "主要488名單",
    "delta": -0.007750999999998953,
    "subGroup": "主要名單-西區",
    "region": "界線以西",
    "color": "紅色",
    "colorHex": "#e53e3e"
  },
  {
    "id": "store-370",
    "name": "自然有漁有限公司",
    "address": "花蓮縣花蓮市建國路23號",
    "lat": 23.982823,
    "lng": 121.606205,
    "estimated": false,
    "isPrimary488": true,
    "group": "主要488名單",
    "delta": -0.0005975769230701644,
    "subGroup": "主要名單-西區",
    "region": "界線以西",
    "color": "紅色",
    "colorHex": "#e53e3e"
  },
  {
    "id": "store-371",
    "name": "同仁堂中醫診所",
    "address": "花蓮縣花蓮市建國路256號",
    "lat": 23.984358,
    "lng": 121.599658,
    "estimated": false,
    "isPrimary488": true,
    "group": "主要488名單",
    "delta": -0.008856692307688263,
    "subGroup": "主要名單-西區",
    "region": "界線以西",
    "color": "紅色",
    "colorHex": "#e53e3e"
  },
  {
    "id": "store-372",
    "name": "詠聖車業",
    "address": "花蓮縣花蓮市建國路88號",
    "lat": 23.983606,
    "lng": 121.603418,
    "estimated": false,
    "isPrimary488": true,
    "group": "主要488名單",
    "delta": -0.004257923076920633,
    "subGroup": "主要名單-西區",
    "region": "界線以西",
    "color": "紅色",
    "colorHex": "#e53e3e"
  },
  {
    "id": "store-373",
    "name": "重慶九九五金日用品",
    "address": "花蓮縣花蓮市重慶路105號",
    "lat": 23.967607,
    "lng": 121.601267,
    "estimated": false,
    "isPrimary488": true,
    "group": "主要488名單",
    "delta": 0.011436115384611867,
    "subGroup": "主要名單-東區",
    "region": "界線以東",
    "color": "綠色",
    "colorHex": "#257a4e"
  },
  {
    "id": "store-374",
    "name": "國泰世華銀行ATM全聯花蓮重慶店",
    "address": "花蓮縣花蓮市重慶路331.333號",
    "lat": 23.97192,
    "lng": 121.607468,
    "estimated": false,
    "isPrimary488": true,
    "group": "主要488名單",
    "delta": 0.01282646153845235,
    "subGroup": "主要名單-東區",
    "region": "界線以東",
    "color": "綠色",
    "colorHex": "#257a4e"
  },
  {
    "id": "store-375",
    "name": "全聯花蓮重慶2100",
    "address": "花蓮縣花蓮市重慶路331號,333號",
    "lat": 23.97192,
    "lng": 121.607468,
    "estimated": false,
    "isPrimary488": true,
    "group": "主要488名單",
    "delta": 0.01282646153845235,
    "subGroup": "主要名單-東區",
    "region": "界線以東",
    "color": "綠色",
    "colorHex": "#257a4e"
  },
  {
    "id": "store-376",
    "name": "繼陽建材行",
    "address": "花蓮縣花蓮市重慶路386號",
    "lat": 23.970772,
    "lng": 121.606765,
    "estimated": false,
    "isPrimary488": true,
    "group": "主要488名單",
    "delta": 0.013403923076921842,
    "subGroup": "主要名單-東區",
    "region": "界線以東",
    "color": "綠色",
    "colorHex": "#257a4e"
  },
  {
    "id": "store-377",
    "name": "台新銀行ATM統冠超市重慶店",
    "address": "花蓮縣花蓮市重慶路415號",
    "lat": 23.973287,
    "lng": 121.609895,
    "estimated": false,
    "isPrimary488": true,
    "group": "主要488名單",
    "delta": 0.013728730769230424,
    "subGroup": "主要名單-東區",
    "region": "界線以東",
    "color": "綠色",
    "colorHex": "#257a4e"
  },
  {
    "id": "store-378",
    "name": "統冠超商重慶店",
    "address": "花蓮縣花蓮市重慶路415號",
    "lat": 23.973287,
    "lng": 121.609895,
    "estimated": false,
    "isPrimary488": true,
    "group": "主要488名單",
    "delta": 0.013728730769230424,
    "subGroup": "主要名單-東區",
    "region": "界線以東",
    "color": "綠色",
    "colorHex": "#257a4e"
  },
  {
    "id": "store-379",
    "name": "李永康公館",
    "address": "花蓮縣花蓮市重慶路89號",
    "lat": 23.967119,
    "lng": 121.600563,
    "estimated": false,
    "isPrimary488": false,
    "group": "非488名單(另色標記)",
    "subGroup": "非488名單",
    "region": "非488店家",
    "color": "黃色",
    "colorHex": "#f59e0b"
  },
  {
    "id": "store-380",
    "name": "陳秉越營業所",
    "address": "花蓮縣花蓮市海濱街31號",
    "lat": 23.972778,
    "lng": 121.6143,
    "estimated": false,
    "isPrimary488": true,
    "group": "主要488名單",
    "delta": 0.018701461538455533,
    "subGroup": "主要名單-東區",
    "region": "界線以東",
    "color": "綠色",
    "colorHex": "#257a4e"
  },
  {
    "id": "store-381",
    "name": "衛生福利部中央健康保險署",
    "address": "花蓮縣花蓮市軒轅路36號",
    "lat": 23.976552,
    "lng": 121.612048,
    "estimated": false,
    "isPrimary488": true,
    "group": "主要488名單",
    "delta": 0.01224000000000558,
    "subGroup": "主要名單-東區",
    "region": "界線以東",
    "color": "綠色",
    "colorHex": "#257a4e"
  },
  {
    "id": "store-382",
    "name": "台灣銀行ATM花蓮健保局",
    "address": "花蓮縣花蓮市軒轅路36號",
    "lat": 23.976552,
    "lng": 121.612048,
    "estimated": false,
    "isPrimary488": true,
    "group": "主要488名單",
    "delta": 0.01224000000000558,
    "subGroup": "主要名單-東區",
    "region": "界線以東",
    "color": "綠色",
    "colorHex": "#257a4e"
  },
  {
    "id": "store-383",
    "name": "台灣銀行ATM反偵測健保局",
    "address": "花蓮縣花蓮市軒轅路36號",
    "lat": 23.976552,
    "lng": 121.612048,
    "estimated": false,
    "isPrimary488": false,
    "group": "非488名單(另色標記)",
    "subGroup": "非488名單",
    "region": "非488店家",
    "color": "黃色",
    "colorHex": "#f59e0b"
  },
  {
    "id": "store-384",
    "name": "衫溫暖洗衣店",
    "address": "花蓮縣花蓮市軒轅路7號",
    "lat": 23.974742,
    "lng": 121.611575,
    "estimated": false,
    "isPrimary488": false,
    "group": "非488名單(另色標記)",
    "subGroup": "非488名單",
    "region": "非488店家",
    "color": "黃色",
    "colorHex": "#f59e0b"
  },
  {
    "id": "store-385",
    "name": "八丼日式手作料理店",
    "address": "花蓮縣花蓮市商校街196號",
    "lat": 23.987118,
    "lng": 121.600653,
    "estimated": false,
    "isPrimary488": true,
    "group": "主要488名單",
    "delta": -0.01094015384614977,
    "subGroup": "主要名單-西區",
    "region": "界線以西",
    "color": "紅色",
    "colorHex": "#e53e3e"
  },
  {
    "id": "store-386",
    "name": "花蓮縣農會無毒農特產品展售中心",
    "address": "花蓮縣花蓮市商校街234號",
    "lat": 23.986951,
    "lng": 121.59976,
    "estimated": false,
    "isPrimary488": true,
    "group": "主要488名單",
    "delta": -0.011646884615373665,
    "subGroup": "主要名單-西區",
    "region": "界線以西",
    "color": "紅色",
    "colorHex": "#e53e3e"
  },
  {
    "id": "store-387",
    "name": "花蓮縣營造業職業工會",
    "address": "花蓮縣花蓮市商校街266號",
    "lat": 23.986684,
    "lng": 121.59832,
    "estimated": false,
    "isPrimary488": true,
    "group": "主要488名單",
    "delta": -0.012789076923070297,
    "subGroup": "主要名單-西區",
    "region": "界線以西",
    "color": "紅色",
    "colorHex": "#e53e3e"
  },
  {
    "id": "store-388",
    "name": "劉燕湖建築師事務所商校街",
    "address": "花蓮縣花蓮市商校街68號",
    "lat": 23.988013,
    "lng": 121.604397,
    "estimated": false,
    "isPrimary488": true,
    "group": "主要488名單",
    "delta": -0.008194423076915314,
    "subGroup": "主要名單-西區",
    "region": "界線以西",
    "color": "紅色",
    "colorHex": "#e53e3e"
  },
  {
    "id": "store-389",
    "name": "桃花源視聽伴唱",
    "address": "花蓮縣花蓮市國民八街51號1樓",
    "lat": 23.993366,
    "lng": 121.605182,
    "estimated": false,
    "isPrimary488": true,
    "group": "主要488名單",
    "delta": -0.013380076923070305,
    "subGroup": "主要名單-西區",
    "region": "界線以西",
    "color": "紅色",
    "colorHex": "#e53e3e"
  },
  {
    "id": "store-390",
    "name": "藶峰有限公司火車站右機車停車場",
    "address": "花蓮縣花蓮市國民段466、476、477-2、483地號",
    "lat": 23.9926,
    "lng": 121.6034,
    "estimated": true,
    "isPrimary488": false,
    "group": "非488名單(另色標記)",
    "subGroup": "非488名單",
    "region": "非488店家",
    "color": "黃色",
    "colorHex": "#f59e0b"
  },
  {
    "id": "store-391",
    "name": "藶峰有限公司火車站左汽車停車場",
    "address": "花蓮縣花蓮市國民段483-2地號",
    "lat": 23.9926,
    "lng": 121.6034,
    "estimated": true,
    "isPrimary488": false,
    "group": "非488名單(另色標記)",
    "subGroup": "非488名單",
    "region": "非488店家",
    "color": "黃色",
    "colorHex": "#f59e0b"
  },
  {
    "id": "store-392",
    "name": "福星行陳春香公館",
    "address": "花蓮縣花蓮市國風街21巷1號",
    "lat": 23.977012,
    "lng": 121.597632,
    "estimated": false,
    "isPrimary488": true,
    "group": "主要488名單",
    "delta": -0.0026890769230760725,
    "subGroup": "主要名單-西區",
    "region": "界線以西",
    "color": "紅色",
    "colorHex": "#e53e3e"
  },
  {
    "id": "store-393",
    "name": "啟業餐飲文化企業社",
    "address": "花蓮縣花蓮市國風街22號",
    "lat": 23.977453,
    "lng": 121.59779,
    "estimated": false,
    "isPrimary488": false,
    "group": "非488名單(另色標記)",
    "subGroup": "非488名單",
    "region": "非488店家",
    "color": "黃色",
    "colorHex": "#f59e0b"
  },
  {
    "id": "store-394",
    "name": "賴炤霖國風自助洗衣店",
    "address": "花蓮縣花蓮市國風街23號",
    "lat": 23.97722,
    "lng": 121.59777,
    "estimated": false,
    "isPrimary488": false,
    "group": "非488名單(另色標記)",
    "subGroup": "非488名單",
    "region": "非488店家",
    "color": "黃色",
    "colorHex": "#f59e0b"
  },
  {
    "id": "store-395",
    "name": "鍾向前公館",
    "address": "花蓮縣花蓮市國風街28號",
    "lat": 23.977554,
    "lng": 121.597664,
    "estimated": false,
    "isPrimary488": true,
    "group": "主要488名單",
    "delta": -0.0032616153846163343,
    "subGroup": "主要名單-西區",
    "region": "界線以西",
    "color": "紅色",
    "colorHex": "#e53e3e"
  },
  {
    "id": "store-396",
    "name": "富陽商號",
    "address": "花蓮縣花蓮市國風街66號",
    "lat": 23.978287,
    "lng": 121.596765,
    "estimated": false,
    "isPrimary488": true,
    "group": "主要488名單",
    "delta": -0.004978192307689255,
    "subGroup": "主要名單-西區",
    "region": "界線以西",
    "color": "紅色",
    "colorHex": "#e53e3e"
  },
  {
    "id": "store-397",
    "name": "統冠超商國聯店",
    "address": "花蓮縣花蓮市國盛里國聯五路150號",
    "lat": 23.992836,
    "lng": 121.605281,
    "estimated": false,
    "isPrimary488": true,
    "group": "主要488名單",
    "delta": -0.01268992307691974,
    "subGroup": "主要名單-西區",
    "region": "界線以西",
    "color": "紅色",
    "colorHex": "#e53e3e"
  },
  {
    "id": "store-398",
    "name": "呂紹仁公館",
    "address": "花蓮縣花蓮市國富三街48號",
    "lat": 23.988644,
    "lng": 121.598264,
    "estimated": false,
    "isPrimary488": true,
    "group": "主要488名單",
    "delta": -0.015031230769224635,
    "subGroup": "主要名單-西區",
    "region": "界線以西",
    "color": "紅色",
    "colorHex": "#e53e3e"
  },
  {
    "id": "store-399",
    "name": "劉莉淳公館",
    "address": "花蓮縣花蓮市國興一街129號",
    "lat": 23.996368,
    "lng": 121.604321,
    "estimated": false,
    "isPrimary488": false,
    "group": "非488名單(另色標記)",
    "subGroup": "非488名單",
    "region": "非488店家",
    "color": "黃色",
    "colorHex": "#f59e0b"
  },
  {
    "id": "store-400",
    "name": "集福行",
    "address": "花蓮縣花蓮市國興一街300號",
    "lat": 24.000819,
    "lng": 121.607259,
    "estimated": false,
    "isPrimary488": false,
    "group": "非488名單(另色標記)",
    "subGroup": "非488名單",
    "region": "非488店家",
    "color": "黃色",
    "colorHex": "#f59e0b"
  },
  {
    "id": "store-401",
    "name": "順豐速運花蓮理貨站",
    "address": "花蓮縣花蓮市國興一街76號",
    "lat": 23.997167,
    "lng": 121.604881,
    "estimated": false,
    "isPrimary488": true,
    "group": "主要488名單",
    "delta": -0.01792065384614716,
    "subGroup": "主要名單-西區",
    "region": "界線以西",
    "color": "紅色",
    "colorHex": "#e53e3e"
  },
  {
    "id": "store-402",
    "name": "元盛百貨股份有限公司",
    "address": "花蓮縣花蓮市國興二街148號",
    "lat": 23.989036,
    "lng": 121.601354,
    "estimated": false,
    "isPrimary488": true,
    "group": "主要488名單",
    "delta": -0.012378461538460783,
    "subGroup": "主要名單-西區",
    "region": "界線以西",
    "color": "紅色",
    "colorHex": "#e53e3e"
  },
  {
    "id": "store-403",
    "name": "朝瓏展示天地",
    "address": "花蓮縣花蓮市國興二街21號",
    "lat": 23.992515,
    "lng": 121.604212,
    "estimated": false,
    "isPrimary488": false,
    "group": "非488名單(另色標記)",
    "subGroup": "非488名單",
    "region": "非488店家",
    "color": "黃色",
    "colorHex": "#f59e0b"
  },
  {
    "id": "store-404",
    "name": "嚴吳明公館",
    "address": "花蓮縣花蓮市國興五街125號",
    "lat": 23.988108,
    "lng": 121.60326,
    "estimated": false,
    "isPrimary488": true,
    "group": "主要488名單",
    "delta": -0.009437384615381461,
    "subGroup": "主要名單-西區",
    "region": "界線以西",
    "color": "紅色",
    "colorHex": "#e53e3e"
  },
  {
    "id": "store-405",
    "name": "台灣鐵路局花蓮機廠廠房",
    "address": "花蓮縣花蓮市國聯一路100-5號",
    "lat": 23.992615,
    "lng": 121.601329,
    "estimated": false,
    "isPrimary488": true,
    "group": "主要488名單",
    "delta": -0.016395423076914994,
    "subGroup": "主要名單-西區",
    "region": "界線以西",
    "color": "紅色",
    "colorHex": "#e53e3e"
  },
  {
    "id": "store-406",
    "name": "台灣鐵路局花蓮機廠行政大樓",
    "address": "花蓮縣花蓮市國聯一路100-5號",
    "lat": 23.992615,
    "lng": 121.601329,
    "estimated": false,
    "isPrimary488": true,
    "group": "主要488名單",
    "delta": -0.016395423076914994,
    "subGroup": "主要名單-西區",
    "region": "界線以西",
    "color": "紅色",
    "colorHex": "#e53e3e"
  },
  {
    "id": "store-407",
    "name": "爭鮮外帶花蓮店",
    "address": "花蓮縣花蓮市國聯一路100號3樓",
    "lat": 23.992615,
    "lng": 121.601329,
    "estimated": false,
    "isPrimary488": true,
    "group": "主要488名單",
    "delta": -0.016395423076914994,
    "subGroup": "主要名單-西區",
    "region": "界線以西",
    "color": "紅色",
    "colorHex": "#e53e3e"
  },
  {
    "id": "store-408",
    "name": "曾記麻糬股份有限公司",
    "address": "花蓮縣花蓮市國聯一路100號花蓮火車站商業區域12號",
    "lat": 23.992888,
    "lng": 121.601303,
    "estimated": true,
    "isPrimary488": true,
    "group": "主要488名單",
    "delta": -0.016725923076919003,
    "subGroup": "主要名單-西區",
    "region": "界線以西",
    "color": "紅色",
    "colorHex": "#e53e3e"
  },
  {
    "id": "store-409",
    "name": "龍鳳視聽歌唱",
    "address": "花蓮縣花蓮市國聯一路13,15,17號",
    "lat": 23.993936,
    "lng": 121.604195,
    "estimated": false,
    "isPrimary488": true,
    "group": "主要488名單",
    "delta": -0.01500284615383407,
    "subGroup": "主要名單-西區",
    "region": "界線以西",
    "color": "紅色",
    "colorHex": "#e53e3e"
  },
  {
    "id": "store-410",
    "name": "金芙蓉視聽歌唱",
    "address": "花蓮縣花蓮市國聯一路143號1樓",
    "lat": 23.990172,
    "lng": 121.601835,
    "estimated": false,
    "isPrimary488": true,
    "group": "主要488名單",
    "delta": -0.0131645384615382,
    "subGroup": "主要名單-西區",
    "region": "界線以西",
    "color": "紅色",
    "colorHex": "#e53e3e"
  },
  {
    "id": "store-411",
    "name": "中國信託花蓮分行",
    "address": "花蓮縣花蓮市國聯一路167號1-2樓",
    "lat": 23.989764,
    "lng": 121.601386,
    "estimated": false,
    "isPrimary488": true,
    "group": "主要488名單",
    "delta": -0.01315846153845257,
    "subGroup": "主要名單-西區",
    "region": "界線以西",
    "color": "紅色",
    "colorHex": "#e53e3e"
  },
  {
    "id": "store-412",
    "name": "中國信託ATM花蓮分行",
    "address": "花蓮縣花蓮市國聯一路167號1-2樓",
    "lat": 23.989764,
    "lng": 121.601386,
    "estimated": false,
    "isPrimary488": true,
    "group": "主要488名單",
    "delta": -0.01315846153845257,
    "subGroup": "主要名單-西區",
    "region": "界線以西",
    "color": "紅色",
    "colorHex": "#e53e3e"
  },
  {
    "id": "store-413",
    "name": "花蓮野人行",
    "address": "花蓮縣花蓮市國聯一路169號",
    "lat": 23.989632,
    "lng": 121.601248,
    "estimated": false,
    "isPrimary488": true,
    "group": "主要488名單",
    "delta": -0.013149230769229803,
    "subGroup": "主要名單-西區",
    "region": "界線以西",
    "color": "紅色",
    "colorHex": "#e53e3e"
  },
  {
    "id": "store-414",
    "name": "全兵衛餐飲有限公司",
    "address": "花蓮縣花蓮市國聯一路173號",
    "lat": 23.989497,
    "lng": 121.601156,
    "estimated": false,
    "isPrimary488": true,
    "group": "主要488名單",
    "delta": -0.013090653846148825,
    "subGroup": "主要名單-西區",
    "region": "界線以西",
    "color": "紅色",
    "colorHex": "#e53e3e"
  },
  {
    "id": "store-415",
    "name": "佑昌企業社無人停車場",
    "address": "花蓮縣花蓮市國聯一路18號",
    "lat": 23.994077,
    "lng": 121.60386,
    "estimated": false,
    "isPrimary488": true,
    "group": "主要488名單",
    "delta": -0.015495115384609903,
    "subGroup": "主要名單-西區",
    "region": "界線以西",
    "color": "紅色",
    "colorHex": "#e53e3e"
  },
  {
    "id": "store-416",
    "name": "親親檳榔攤-國聯一路",
    "address": "花蓮縣花蓮市國聯一路18號",
    "lat": 23.994077,
    "lng": 121.60386,
    "estimated": false,
    "isPrimary488": true,
    "group": "主要488名單",
    "delta": -0.015495115384609903,
    "subGroup": "主要名單-西區",
    "region": "界線以西",
    "color": "紅色",
    "colorHex": "#e53e3e"
  },
  {
    "id": "store-417",
    "name": "新光產物保險花蓮分公司",
    "address": "花蓮縣花蓮市國聯一路34號7樓",
    "lat": 23.99383,
    "lng": 121.603904,
    "estimated": false,
    "isPrimary488": false,
    "group": "非488名單(另色標記)",
    "subGroup": "非488名單",
    "region": "非488店家",
    "color": "黃色",
    "colorHex": "#f59e0b"
  },
  {
    "id": "store-418",
    "name": "新光人壽花蓮大樓",
    "address": "花蓮縣花蓮市國聯一路37號",
    "lat": 23.993438,
    "lng": 121.603923,
    "estimated": false,
    "isPrimary488": true,
    "group": "主要488名單",
    "delta": -0.014719384615389686,
    "subGroup": "主要名單-西區",
    "region": "界線以西",
    "color": "紅色",
    "colorHex": "#e53e3e"
  },
  {
    "id": "store-419",
    "name": "新光銀行ATM新壽花蓮大樓",
    "address": "花蓮縣花蓮市國聯一路37號",
    "lat": 23.993438,
    "lng": 121.603923,
    "estimated": false,
    "isPrimary488": true,
    "group": "主要488名單",
    "delta": -0.014719384615389686,
    "subGroup": "主要名單-西區",
    "region": "界線以西",
    "color": "紅色",
    "colorHex": "#e53e3e"
  },
  {
    "id": "store-420",
    "name": "台灣世界展望會東區辦事處",
    "address": "花蓮縣花蓮市國聯一路37號3樓",
    "lat": 23.993438,
    "lng": 121.603923,
    "estimated": false,
    "isPrimary488": true,
    "group": "主要488名單",
    "delta": -0.014719384615389686,
    "subGroup": "主要名單-西區",
    "region": "界線以西",
    "color": "紅色",
    "colorHex": "#e53e3e"
  },
  {
    "id": "store-421",
    "name": "新光保全花蓮駐區",
    "address": "花蓮縣花蓮市國聯一路37號4樓",
    "lat": 23.993438,
    "lng": 121.603923,
    "estimated": false,
    "isPrimary488": true,
    "group": "主要488名單",
    "delta": -0.014719384615389686,
    "subGroup": "主要名單-西區",
    "region": "界線以西",
    "color": "紅色",
    "colorHex": "#e53e3e"
  },
  {
    "id": "store-422",
    "name": "台灣保全花蓮金庫",
    "address": "花蓮縣花蓮市國聯一路37號4樓",
    "lat": 23.993438,
    "lng": 121.603923,
    "estimated": false,
    "isPrimary488": true,
    "group": "主要488名單",
    "delta": -0.014719384615389686,
    "subGroup": "主要名單-西區",
    "region": "界線以西",
    "color": "紅色",
    "colorHex": "#e53e3e"
  },
  {
    "id": "store-423",
    "name": "新光人壽花市收費處",
    "address": "花蓮縣花蓮市國聯一路37號5樓",
    "lat": 23.993438,
    "lng": 121.603923,
    "estimated": false,
    "isPrimary488": true,
    "group": "主要488名單",
    "delta": -0.014719384615389686,
    "subGroup": "主要名單-西區",
    "region": "界線以西",
    "color": "紅色",
    "colorHex": "#e53e3e"
  },
  {
    "id": "store-424",
    "name": "新光人壽意外險",
    "address": "花蓮縣花蓮市國聯一路37號6F",
    "lat": 23.993438,
    "lng": 121.603923,
    "estimated": false,
    "isPrimary488": true,
    "group": "主要488名單",
    "delta": -0.014719384615389686,
    "subGroup": "主要名單-西區",
    "region": "界線以西",
    "color": "紅色",
    "colorHex": "#e53e3e"
  },
  {
    "id": "store-425",
    "name": "承億文旅股份有限公司花蓮分公司",
    "address": "花蓮縣花蓮市國聯一路39號",
    "lat": 23.99318,
    "lng": 121.603813,
    "estimated": false,
    "isPrimary488": true,
    "group": "主要488名單",
    "delta": -0.014541615384601414,
    "subGroup": "主要名單-西區",
    "region": "界線以西",
    "color": "紅色",
    "colorHex": "#e53e3e"
  },
  {
    "id": "store-426",
    "name": "和運租車花蓮站",
    "address": "花蓮縣花蓮市國聯一路69號",
    "lat": 23.99186,
    "lng": 121.60332,
    "estimated": false,
    "isPrimary488": false,
    "group": "非488名單(另色標記)",
    "subGroup": "非488名單",
    "region": "非488店家",
    "color": "黃色",
    "colorHex": "#f59e0b"
  },
  {
    "id": "store-427",
    "name": "梅貴汽車租賃有限公司",
    "address": "花蓮縣花蓮市國聯一路95-2號",
    "lat": 23.991216,
    "lng": 121.602723,
    "estimated": false,
    "isPrimary488": true,
    "group": "主要488名單",
    "delta": -0.013441000000000258,
    "subGroup": "主要名單-西區",
    "region": "界線以西",
    "color": "紅色",
    "colorHex": "#e53e3e"
  },
  {
    "id": "store-428",
    "name": "花蓮先期轉運站",
    "address": "花蓮縣花蓮市國聯一路98號",
    "lat": 23.99299,
    "lng": 121.601849,
    "estimated": false,
    "isPrimary488": true,
    "group": "主要488名單",
    "delta": -0.016293692307684182,
    "subGroup": "主要名單-西區",
    "region": "界線以西",
    "color": "紅色",
    "colorHex": "#e53e3e"
  },
  {
    "id": "store-429",
    "name": "全聯花蓮2110",
    "address": "花蓮縣花蓮市國聯二路110號",
    "lat": 23.988928,
    "lng": 121.602101,
    "estimated": false,
    "isPrimary488": true,
    "group": "主要488名單",
    "delta": -0.011510999999998717,
    "subGroup": "主要名單-西區",
    "region": "界線以西",
    "color": "紅色",
    "colorHex": "#e53e3e"
  },
  {
    "id": "store-430",
    "name": "國泰世華銀行ATM全聯國聯店",
    "address": "花蓮縣花蓮市國聯二路110號",
    "lat": 23.988928,
    "lng": 121.602101,
    "estimated": false,
    "isPrimary488": true,
    "group": "主要488名單",
    "delta": -0.011510999999998717,
    "subGroup": "主要名單-西區",
    "region": "界線以西",
    "color": "紅色",
    "colorHex": "#e53e3e"
  },
  {
    "id": "store-431",
    "name": "勝泰水電衛生材料行",
    "address": "花蓮縣花蓮市國聯二路132號",
    "lat": 23.988492,
    "lng": 121.601642,
    "estimated": false,
    "isPrimary488": true,
    "group": "主要488名單",
    "delta": -0.011483692307692195,
    "subGroup": "主要名單-西區",
    "region": "界線以西",
    "color": "紅色",
    "colorHex": "#e53e3e"
  },
  {
    "id": "store-432",
    "name": "台灣松下銷售股份有限公司花蓮處",
    "address": "花蓮縣花蓮市國聯二路153號",
    "lat": 23.988368,
    "lng": 121.601954,
    "estimated": false,
    "isPrimary488": false,
    "group": "非488名單(另色標記)",
    "subGroup": "非488名單",
    "region": "非488店家",
    "color": "黃色",
    "colorHex": "#f59e0b"
  },
  {
    "id": "store-433",
    "name": "富豪鍋物",
    "address": "花蓮縣花蓮市國聯二路19號",
    "lat": 23.991024,
    "lng": 121.604494,
    "estimated": false,
    "isPrimary488": true,
    "group": "主要488名單",
    "delta": -0.011455846153836546,
    "subGroup": "主要名單-西區",
    "region": "界線以西",
    "color": "紅色",
    "colorHex": "#e53e3e"
  },
  {
    "id": "store-434",
    "name": "健光陽光休閒館",
    "address": "花蓮縣花蓮市國聯二路39號",
    "lat": 23.990445,
    "lng": 121.603915,
    "estimated": false,
    "isPrimary488": true,
    "group": "主要488名單",
    "delta": -0.011389038461530276,
    "subGroup": "主要名單-西區",
    "region": "界線以西",
    "color": "紅色",
    "colorHex": "#e53e3e"
  },
  {
    "id": "store-435",
    "name": "灶咖餐廳",
    "address": "花蓮縣花蓮市國聯二路70-1號",
    "lat": 23.990015,
    "lng": 121.603039,
    "estimated": false,
    "isPrimary488": false,
    "group": "非488名單(另色標記)",
    "subGroup": "非488名單",
    "region": "非488店家",
    "color": "黃色",
    "colorHex": "#f59e0b"
  },
  {
    "id": "store-436",
    "name": "許永銳營業所",
    "address": "花蓮縣花蓮市國聯二路81號",
    "lat": 23.989709,
    "lng": 121.603267,
    "estimated": false,
    "isPrimary488": false,
    "group": "非488名單(另色標記)",
    "subGroup": "非488名單",
    "region": "非488店家",
    "color": "黃色",
    "colorHex": "#f59e0b"
  },
  {
    "id": "store-437",
    "name": "家豐家居生活館",
    "address": "花蓮縣花蓮市國聯二路90號",
    "lat": 23.989646,
    "lng": 121.602751,
    "estimated": false,
    "isPrimary488": false,
    "group": "非488名單(另色標記)",
    "subGroup": "非488名單",
    "region": "非488店家",
    "color": "黃色",
    "colorHex": "#f59e0b"
  },
  {
    "id": "store-438",
    "name": "欣欣水泥企業股份有限公司",
    "address": "花蓮縣花蓮市國聯二路95.97號",
    "lat": 23.989336,
    "lng": 121.602983,
    "estimated": false,
    "isPrimary488": true,
    "group": "主要488名單",
    "delta": -0.01108407692308333,
    "subGroup": "主要名單-西區",
    "region": "界線以西",
    "color": "紅色",
    "colorHex": "#e53e3e"
  },
  {
    "id": "store-439",
    "name": "欣欣水泥企業股份有限公司",
    "address": "花蓮縣花蓮市國聯二路95~97號",
    "lat": 23.989336,
    "lng": 121.602983,
    "estimated": false,
    "isPrimary488": true,
    "group": "主要488名單",
    "delta": -0.01108407692308333,
    "subGroup": "主要名單-西區",
    "region": "界線以西",
    "color": "紅色",
    "colorHex": "#e53e3e"
  },
  {
    "id": "store-440",
    "name": "宗泰食品站前店",
    "address": "花蓮縣花蓮市國聯三路2號",
    "lat": 23.991355,
    "lng": 121.603016,
    "estimated": false,
    "isPrimary488": true,
    "group": "主要488名單",
    "delta": -0.013303038461529582,
    "subGroup": "主要名單-西區",
    "region": "界線以西",
    "color": "紅色",
    "colorHex": "#e53e3e"
  },
  {
    "id": "store-441",
    "name": "照典電器材料有限公司",
    "address": "花蓮縣花蓮市國聯三路78號",
    "lat": 23.990125,
    "lng": 121.605098,
    "estimated": false,
    "isPrimary488": true,
    "group": "主要488名單",
    "delta": -0.009849115384611196,
    "subGroup": "主要名單-西區",
    "region": "界線以西",
    "color": "紅色",
    "colorHex": "#e53e3e"
  },
  {
    "id": "store-442",
    "name": "台新銀行ATM花蓮統冠超市國聯",
    "address": "花蓮縣花蓮市國聯五路150號",
    "lat": 23.992836,
    "lng": 121.605281,
    "estimated": false,
    "isPrimary488": true,
    "group": "主要488名單",
    "delta": -0.01268992307691974,
    "subGroup": "主要名單-西區",
    "region": "界線以西",
    "color": "紅色",
    "colorHex": "#e53e3e"
  },
  {
    "id": "store-443",
    "name": "谷田有限公司",
    "address": "花蓮縣花蓮市國聯五路219號",
    "lat": 23.994929,
    "lng": 121.604452,
    "estimated": false,
    "isPrimary488": true,
    "group": "主要488名單",
    "delta": -0.01585342307691917,
    "subGroup": "主要名單-西區",
    "region": "界線以西",
    "color": "紅色",
    "colorHex": "#e53e3e"
  },
  {
    "id": "store-444",
    "name": "豐眼鏡",
    "address": "花蓮縣花蓮市國聯五路44號",
    "lat": 23.989953,
    "lng": 121.605995,
    "estimated": false,
    "isPrimary488": false,
    "group": "非488名單(另色標記)",
    "subGroup": "非488名單",
    "region": "非488店家",
    "color": "黃色",
    "colorHex": "#f59e0b"
  },
  {
    "id": "store-445",
    "name": "全家來便利商店",
    "address": "花蓮縣花蓮市國聯五路51號1樓",
    "lat": 23.990097,
    "lng": 121.605407,
    "estimated": false,
    "isPrimary488": true,
    "group": "主要488名單",
    "delta": -0.009508884615385682,
    "subGroup": "主要名單-西區",
    "region": "界線以西",
    "color": "紅色",
    "colorHex": "#e53e3e"
  },
  {
    "id": "store-446",
    "name": "辰宇公共",
    "address": "花蓮縣花蓮市國聯五路5-1號旁",
    "lat": 23.988994,
    "lng": 121.605931,
    "estimated": true,
    "isPrimary488": true,
    "group": "主要488名單",
    "delta": -0.007754615384612862,
    "subGroup": "主要名單-西區",
    "region": "界線以西",
    "color": "紅色",
    "colorHex": "#e53e3e"
  },
  {
    "id": "store-447",
    "name": "林怡欣醫美診所",
    "address": "花蓮縣花蓮市國聯五路5-5號3樓",
    "lat": 23.988994,
    "lng": 121.605931,
    "estimated": false,
    "isPrimary488": true,
    "group": "主要488名單",
    "delta": -0.007754615384612862,
    "subGroup": "主要名單-西區",
    "region": "界線以西",
    "color": "紅色",
    "colorHex": "#e53e3e"
  },
  {
    "id": "store-448",
    "name": "辰宇7樓",
    "address": "花蓮縣花蓮市國聯五路5之5 7樓",
    "lat": 23.988994,
    "lng": 121.605931,
    "estimated": false,
    "isPrimary488": true,
    "group": "主要488名單",
    "delta": -0.007754615384612862,
    "subGroup": "主要名單-西區",
    "region": "界線以西",
    "color": "紅色",
    "colorHex": "#e53e3e"
  },
  {
    "id": "store-449",
    "name": "華蓮秀泰花蓮影城",
    "address": "花蓮縣花蓮市國聯五路69號4F",
    "lat": 23.99087,
    "lng": 121.605345,
    "estimated": false,
    "isPrimary488": true,
    "group": "主要488名單",
    "delta": -0.010433076923078488,
    "subGroup": "主要名單-西區",
    "region": "界線以西",
    "color": "紅色",
    "colorHex": "#e53e3e"
  },
  {
    "id": "store-450",
    "name": "華蓮秀泰花蓮影城",
    "address": "花蓮縣花蓮市國聯五路69號5樓",
    "lat": 23.99087,
    "lng": 121.605345,
    "estimated": false,
    "isPrimary488": true,
    "group": "主要488名單",
    "delta": -0.010433076923078488,
    "subGroup": "主要名單-西區",
    "region": "界線以西",
    "color": "紅色",
    "colorHex": "#e53e3e"
  },
  {
    "id": "store-451",
    "name": "華蓮秀泰花蓮影城",
    "address": "花蓮縣花蓮市國聯五路69號6樓",
    "lat": 23.99087,
    "lng": 121.605345,
    "estimated": false,
    "isPrimary488": true,
    "group": "主要488名單",
    "delta": -0.010433076923078488,
    "subGroup": "主要名單-西區",
    "region": "界線以西",
    "color": "紅色",
    "colorHex": "#e53e3e"
  },
  {
    "id": "store-452",
    "name": "新光銀行入金機花蓮秀泰影城",
    "address": "花蓮縣花蓮市國聯五路69號7樓",
    "lat": 23.99087,
    "lng": 121.605345,
    "estimated": false,
    "isPrimary488": true,
    "group": "主要488名單",
    "delta": -0.010433076923078488,
    "subGroup": "主要名單-西區",
    "region": "界線以西",
    "color": "紅色",
    "colorHex": "#e53e3e"
  },
  {
    "id": "store-453",
    "name": "華蓮秀泰花蓮影城",
    "address": "花蓮縣花蓮市國聯五路69號7樓",
    "lat": 23.99087,
    "lng": 121.605345,
    "estimated": false,
    "isPrimary488": true,
    "group": "主要488名單",
    "delta": -0.010433076923078488,
    "subGroup": "主要名單-西區",
    "region": "界線以西",
    "color": "紅色",
    "colorHex": "#e53e3e"
  },
  {
    "id": "store-454",
    "name": "廖治德營業所(秀泰影城B1)",
    "address": "花蓮縣花蓮市國聯五路69號B1",
    "lat": 23.99087,
    "lng": 121.605345,
    "estimated": false,
    "isPrimary488": false,
    "group": "非488名單(另色標記)",
    "subGroup": "非488名單",
    "region": "非488店家",
    "color": "黃色",
    "colorHex": "#f59e0b"
  },
  {
    "id": "store-455",
    "name": "彭德豪營業所",
    "address": "花蓮縣花蓮市國聯五路7-3號",
    "lat": 23.989101,
    "lng": 121.605736,
    "estimated": false,
    "isPrimary488": true,
    "group": "主要488名單",
    "delta": -0.008068961538469921,
    "subGroup": "主要名單-西區",
    "region": "界線以西",
    "color": "紅色",
    "colorHex": "#e53e3e"
  },
  {
    "id": "store-456",
    "name": "合作金庫ATM北花蓮分行國聯",
    "address": "花蓮縣花蓮市國聯五路80號",
    "lat": 23.991069,
    "lng": 121.605601,
    "estimated": false,
    "isPrimary488": true,
    "group": "主要488名單",
    "delta": -0.010399038461542887,
    "subGroup": "主要名單-西區",
    "region": "界線以西",
    "color": "紅色",
    "colorHex": "#e53e3e"
  },
  {
    "id": "store-457",
    "name": "國王光學眼鏡有限公司",
    "address": "花蓮縣花蓮市博愛街109-1號",
    "lat": 23.972928,
    "lng": 121.609769,
    "estimated": false,
    "isPrimary488": true,
    "group": "主要488名單",
    "delta": 0.014003153846147143,
    "subGroup": "主要名單-東區",
    "region": "界線以東",
    "color": "綠色",
    "colorHex": "#257a4e"
  },
  {
    "id": "store-458",
    "name": "博愛燒肉工房",
    "address": "花蓮縣花蓮市博愛街136",
    "lat": 23.977319,
    "lng": 121.606535,
    "estimated": false,
    "isPrimary488": true,
    "group": "主要488名單",
    "delta": 0.005871499999997809,
    "subGroup": "主要名單-東區",
    "region": "界線以東",
    "color": "綠色",
    "colorHex": "#257a4e"
  },
  {
    "id": "store-459",
    "name": "初樂美學空間",
    "address": "花蓮縣花蓮市博愛街146號1樓",
    "lat": 23.977547,
    "lng": 121.606343,
    "estimated": false,
    "isPrimary488": true,
    "group": "主要488名單",
    "delta": 0.005425192307683346,
    "subGroup": "主要名單-東區",
    "region": "界線以東",
    "color": "綠色",
    "colorHex": "#257a4e"
  },
  {
    "id": "store-460",
    "name": "阿色紀有限公司-藝托邦藝文空間",
    "address": "花蓮縣花蓮市博愛街156-5號",
    "lat": 23.978037,
    "lng": 121.605975,
    "estimated": false,
    "isPrimary488": true,
    "group": "主要488名單",
    "delta": 0.0045106538461539,
    "subGroup": "主要名單-東區",
    "region": "界線以東",
    "color": "綠色",
    "colorHex": "#257a4e"
  },
  {
    "id": "store-461",
    "name": "小村日壽司",
    "address": "花蓮縣花蓮市博愛街175號(平信)",
    "lat": 23.976268,
    "lng": 121.607166,
    "estimated": false,
    "isPrimary488": false,
    "group": "非488名單(另色標記)",
    "subGroup": "非488名單",
    "region": "非488店家",
    "color": "黃色",
    "colorHex": "#f59e0b"
  },
  {
    "id": "store-462",
    "name": "樂津廚坊",
    "address": "花蓮縣花蓮市博愛街185號",
    "lat": 23.976467,
    "lng": 121.606995,
    "estimated": false,
    "isPrimary488": true,
    "group": "主要488名單",
    "delta": 0.007281807692308462,
    "subGroup": "主要名單-東區",
    "region": "界線以東",
    "color": "綠色",
    "colorHex": "#257a4e"
  },
  {
    "id": "store-463",
    "name": "錦向有限公司",
    "address": "花蓮縣花蓮市博愛街257號旁",
    "lat": 23.980863,
    "lng": 121.603386,
    "estimated": true,
    "isPrimary488": true,
    "group": "主要488名單",
    "delta": -0.0012304230769188962,
    "subGroup": "主要名單-西區",
    "region": "界線以西",
    "color": "紅色",
    "colorHex": "#e53e3e"
  },
  {
    "id": "store-464",
    "name": "蘋果活動創意行銷",
    "address": "花蓮縣花蓮市博愛街301號",
    "lat": 23.981455,
    "lng": 121.602526,
    "estimated": false,
    "isPrimary488": true,
    "group": "主要488名單",
    "delta": -0.0027507307692360428,
    "subGroup": "主要名單-西區",
    "region": "界線以西",
    "color": "紅色",
    "colorHex": "#e53e3e"
  },
  {
    "id": "store-465",
    "name": "聯登物業有限公司",
    "address": "花蓮縣花蓮市富國段富國一小段4地號",
    "lat": 23.9938,
    "lng": 121.5995,
    "estimated": true,
    "isPrimary488": false,
    "group": "非488名單(另色標記)",
    "subGroup": "非488名單",
    "region": "非488店家",
    "color": "黃色",
    "colorHex": "#f59e0b"
  },
  {
    "id": "store-466",
    "name": "藶峰-站前",
    "address": "花蓮縣花蓮市富國段富國一小段7號地號",
    "lat": 23.9938,
    "lng": 121.5995,
    "estimated": true,
    "isPrimary488": false,
    "group": "非488名單(另色標記)",
    "subGroup": "非488名單",
    "region": "非488店家",
    "color": "黃色",
    "colorHex": "#f59e0b"
  },
  {
    "id": "store-467",
    "name": "懷德診所",
    "address": "花蓮縣花蓮市復興街10號",
    "lat": 23.975854,
    "lng": 121.611384,
    "estimated": false,
    "isPrimary488": true,
    "group": "主要488名單",
    "delta": 0.01235453846153689,
    "subGroup": "主要名單-東區",
    "region": "界線以東",
    "color": "綠色",
    "colorHex": "#257a4e"
  },
  {
    "id": "store-468",
    "name": "花蓮一信復興分社",
    "address": "花蓮縣花蓮市復興街17號",
    "lat": 23.97608,
    "lng": 121.611033,
    "estimated": false,
    "isPrimary488": false,
    "group": "非488名單(另色標記)",
    "subGroup": "非488名單",
    "region": "非488店家",
    "color": "黃色",
    "colorHex": "#f59e0b"
  },
  {
    "id": "store-469",
    "name": "莊謙信公館",
    "address": "花蓮縣花蓮市復興街61號",
    "lat": 23.979726,
    "lng": 121.608107,
    "estimated": false,
    "isPrimary488": true,
    "group": "主要488名單",
    "delta": 0.004758769230775783,
    "subGroup": "主要名單-東區",
    "region": "界線以東",
    "color": "綠色",
    "colorHex": "#257a4e"
  },
  {
    "id": "store-470",
    "name": "林木田聯合地政士事務所",
    "address": "花蓮縣花蓮市復興街96-1號",
    "lat": 23.981501,
    "lng": 121.606945,
    "estimated": false,
    "isPrimary488": false,
    "group": "非488名單(另色標記)",
    "subGroup": "非488名單",
    "region": "非488店家",
    "color": "黃色",
    "colorHex": "#f59e0b"
  },
  {
    "id": "store-471",
    "name": "花蓮女中總務室",
    "address": "花蓮縣花蓮市菁華街2號",
    "lat": 23.979622,
    "lng": 121.615071,
    "estimated": false,
    "isPrimary488": true,
    "group": "主要488名單",
    "delta": 0.011838769230777757,
    "subGroup": "主要名單-東區",
    "region": "界線以東",
    "color": "綠色",
    "colorHex": "#257a4e"
  },
  {
    "id": "store-472",
    "name": "花蓮女中學藝大樓",
    "address": "花蓮縣花蓮市菁華街2號",
    "lat": 23.979622,
    "lng": 121.615071,
    "estimated": false,
    "isPrimary488": true,
    "group": "主要488名單",
    "delta": 0.011838769230777757,
    "subGroup": "主要名單-東區",
    "region": "界線以東",
    "color": "綠色",
    "colorHex": "#257a4e"
  },
  {
    "id": "store-473",
    "name": "台新銀行ATM花蓮女中OK超商",
    "address": "花蓮縣花蓮市菁華街2號",
    "lat": 23.979622,
    "lng": 121.615071,
    "estimated": false,
    "isPrimary488": true,
    "group": "主要488名單",
    "delta": 0.011838769230777757,
    "subGroup": "主要名單-東區",
    "region": "界線以東",
    "color": "綠色",
    "colorHex": "#257a4e"
  },
  {
    "id": "store-474",
    "name": "花蓮女中女生宿舍",
    "address": "花蓮縣花蓮市菁華街2號B1-2樓",
    "lat": 23.979622,
    "lng": 121.615071,
    "estimated": false,
    "isPrimary488": true,
    "group": "主要488名單",
    "delta": 0.011838769230777757,
    "subGroup": "主要名單-東區",
    "region": "界線以東",
    "color": "綠色",
    "colorHex": "#257a4e"
  },
  {
    "id": "store-475",
    "name": "找到了可頌小賣所",
    "address": "花蓮縣花蓮市菁華街35巷1弄18號",
    "lat": 23.981038,
    "lng": 121.615415,
    "estimated": false,
    "isPrimary488": true,
    "group": "主要488名單",
    "delta": 0.010603384615379241,
    "subGroup": "主要名單-東區",
    "region": "界線以東",
    "color": "綠色",
    "colorHex": "#257a4e"
  },
  {
    "id": "store-476",
    "name": "菁華學苑大樓",
    "address": "花蓮縣花蓮市菁華街4巷1弄2號7樓(B2及頂樓)",
    "lat": 23.979802,
    "lng": 121.616551,
    "estimated": true,
    "isPrimary488": true,
    "group": "主要488名單",
    "delta": 0.013118000000005736,
    "subGroup": "主要名單-東區",
    "region": "界線以東",
    "color": "綠色",
    "colorHex": "#257a4e"
  },
  {
    "id": "store-477",
    "name": "國軍花蓮總醫院進豐分院",
    "address": "花蓮縣花蓮市進豐街100號",
    "lat": 23.987405,
    "lng": 121.60653,
    "estimated": false,
    "isPrimary488": true,
    "group": "主要488名單",
    "delta": -0.0053832692307622665,
    "subGroup": "主要名單-西區",
    "region": "界線以西",
    "color": "紅色",
    "colorHex": "#e53e3e"
  },
  {
    "id": "store-478",
    "name": "劉素英公館",
    "address": "花蓮縣花蓮市進豐街66巷27號",
    "lat": 23.986886,
    "lng": 121.607642,
    "estimated": false,
    "isPrimary488": true,
    "group": "主要488名單",
    "delta": -0.003692384615376909,
    "subGroup": "主要名單-西區",
    "region": "界線以西",
    "color": "紅色",
    "colorHex": "#e53e3e"
  },
  {
    "id": "store-479",
    "name": "星璇展業東洋營業所",
    "address": "花蓮縣花蓮市新港街11巷2號",
    "lat": 23.977932,
    "lng": 121.607328,
    "estimated": false,
    "isPrimary488": true,
    "group": "主要488名單",
    "delta": 0.005980769230760075,
    "subGroup": "主要名單-東區",
    "region": "界線以東",
    "color": "綠色",
    "colorHex": "#257a4e"
  },
  {
    "id": "store-480",
    "name": "Morni莫尼早餐花蓮新港店",
    "address": "花蓮縣花蓮市新港街13號",
    "lat": 23.97794,
    "lng": 121.607586,
    "estimated": false,
    "isPrimary488": false,
    "group": "非488名單(另色標記)",
    "subGroup": "非488名單",
    "region": "非488店家",
    "color": "黃色",
    "colorHex": "#f59e0b"
  },
  {
    "id": "store-481",
    "name": "李翼丞-娃娃機新港店",
    "address": "花蓮縣花蓮市新港街41號",
    "lat": 23.978527,
    "lng": 121.607096,
    "estimated": false,
    "isPrimary488": false,
    "group": "非488名單(另色標記)",
    "subGroup": "非488名單",
    "region": "非488店家",
    "color": "黃色",
    "colorHex": "#f59e0b"
  },
  {
    "id": "store-482",
    "name": "遠見通訊",
    "address": "花蓮縣花蓮市新港街52-2號",
    "lat": 23.979215,
    "lng": 121.606785,
    "estimated": false,
    "isPrimary488": true,
    "group": "主要488名單",
    "delta": 0.004006730769233968,
    "subGroup": "主要名單-東區",
    "region": "界線以東",
    "color": "綠色",
    "colorHex": "#257a4e"
  },
  {
    "id": "store-483",
    "name": "精誠輪業有限公司",
    "address": "花蓮縣花蓮市節約街12號1-5樓",
    "lat": 23.97795,
    "lng": 121.606491,
    "estimated": false,
    "isPrimary488": true,
    "group": "主要488名單",
    "delta": 0.005123692307691385,
    "subGroup": "主要名單-東區",
    "region": "界線以東",
    "color": "綠色",
    "colorHex": "#257a4e"
  },
  {
    "id": "store-484",
    "name": "健野戶外休閒用品",
    "address": "花蓮縣花蓮市節約街29號",
    "lat": 23.978418,
    "lng": 121.606916,
    "estimated": false,
    "isPrimary488": true,
    "group": "主要488名單",
    "delta": 0.005026692307694702,
    "subGroup": "主要名單-東區",
    "region": "界線以東",
    "color": "綠色",
    "colorHex": "#257a4e"
  },
  {
    "id": "store-485",
    "name": "日月亭停車場花蓮節約街",
    "address": "花蓮縣花蓮市節約街83號對面",
    "lat": 23.980238,
    "lng": 121.609605,
    "estimated": true,
    "isPrimary488": false,
    "group": "非488名單(另色標記)",
    "subGroup": "非488名單",
    "region": "非488店家",
    "color": "黃色",
    "colorHex": "#f59e0b"
  },
  {
    "id": "store-486",
    "name": "曾記米麻糬榮正街金碧華廈地下室",
    "address": "花蓮縣花蓮市榮正街125巷1之33號",
    "lat": 23.973752,
    "lng": 121.597364,
    "estimated": false,
    "isPrimary488": true,
    "group": "主要488名單",
    "delta": 0.0006790769230775595,
    "subGroup": "主要名單-東區",
    "region": "界線以東",
    "color": "綠色",
    "colorHex": "#257a4e"
  },
  {
    "id": "store-487",
    "name": "幸生外科診所",
    "address": "花蓮縣花蓮市榮正街87號",
    "lat": 23.973394,
    "lng": 121.597641,
    "estimated": false,
    "isPrimary488": true,
    "group": "主要488名單",
    "delta": 0.0013553846153797622,
    "subGroup": "主要名單-東區",
    "region": "界線以東",
    "color": "綠色",
    "colorHex": "#257a4e"
  },
  {
    "id": "store-488",
    "name": "東尼斯美容養生館",
    "address": "花蓮縣花蓮市福建街439號",
    "lat": 23.974397,
    "lng": 121.608864,
    "estimated": false,
    "isPrimary488": true,
    "group": "主要488名單",
    "delta": 0.011459653846145557,
    "subGroup": "主要名單-東區",
    "region": "界線以東",
    "color": "綠色",
    "colorHex": "#257a4e"
  },
  {
    "id": "store-489",
    "name": "建發行",
    "address": "花蓮縣花蓮市福建街525號",
    "lat": 23.97613,
    "lng": 121.611463,
    "estimated": false,
    "isPrimary488": true,
    "group": "主要488名單",
    "delta": 0.012125692307691338,
    "subGroup": "主要名單-東區",
    "region": "界線以東",
    "color": "綠色",
    "colorHex": "#257a4e"
  },
  {
    "id": "store-490",
    "name": "曾記臻匯饌資材室",
    "address": "花蓮縣花蓮市福德段91-1地號",
    "lat": 23.9705,
    "lng": 121.611,
    "estimated": true,
    "isPrimary488": true,
    "group": "主要488名單",
    "delta": 0.017942307692308646,
    "subGroup": "主要名單-東區",
    "region": "界線以東",
    "color": "綠色",
    "colorHex": "#257a4e"
  },
  {
    "id": "store-491",
    "name": "曾記雿來農產福德農舍",
    "address": "花蓮縣花蓮市福德段91地號",
    "lat": 23.9705,
    "lng": 121.611,
    "estimated": true,
    "isPrimary488": true,
    "group": "主要488名單",
    "delta": 0.017942307692308646,
    "subGroup": "主要名單-東區",
    "region": "界線以東",
    "color": "綠色",
    "colorHex": "#257a4e"
  },
  {
    "id": "store-492",
    "name": "明昌食品號(店面、倉庫)",
    "address": "花蓮縣花蓮市綜合市場內68號",
    "lat": 23.9785,
    "lng": 121.6075,
    "estimated": true,
    "isPrimary488": true,
    "group": "主要488名單",
    "delta": 0.005519230769237993,
    "subGroup": "主要名單-東區",
    "region": "界線以東",
    "color": "綠色",
    "colorHex": "#257a4e"
  },
  {
    "id": "store-493",
    "name": "顏清雲營業所",
    "address": "花蓮縣花蓮市綜合市場固定攤位8號",
    "lat": 23.9785,
    "lng": 121.6075,
    "estimated": true,
    "isPrimary488": true,
    "group": "主要488名單",
    "delta": 0.005519230769237993,
    "subGroup": "主要名單-東區",
    "region": "界線以東",
    "color": "綠色",
    "colorHex": "#257a4e"
  },
  {
    "id": "store-494",
    "name": "陳民宏公館",
    "address": "花蓮縣花蓮市廣東街210號",
    "lat": 23.971812,
    "lng": 121.606668,
    "estimated": false,
    "isPrimary488": false,
    "group": "非488名單(另色標記)",
    "subGroup": "非488名單",
    "region": "非488店家",
    "color": "黃色",
    "colorHex": "#f59e0b"
  },
  {
    "id": "store-495",
    "name": "錢記早餐店",
    "address": "花蓮縣花蓮市德安一街213號",
    "lat": 23.972062,
    "lng": 121.594151,
    "estimated": false,
    "isPrimary488": true,
    "group": "主要488名單",
    "delta": -0.0006489230769233245,
    "subGroup": "主要名單-西區",
    "region": "界線以西",
    "color": "紅色",
    "colorHex": "#e53e3e"
  },
  {
    "id": "store-496",
    "name": "奧莉維亞花園婚禮",
    "address": "花蓮縣花蓮市德安二街17號",
    "lat": 23.974275,
    "lng": 121.594057,
    "estimated": false,
    "isPrimary488": true,
    "group": "主要488名單",
    "delta": -0.0032112692307606494,
    "subGroup": "主要名單-西區",
    "region": "界線以西",
    "color": "紅色",
    "colorHex": "#e53e3e"
  },
  {
    "id": "store-497",
    "name": "福星行有限公司",
    "address": "花蓮縣花蓮市鎮國街126號",
    "lat": 23.976541,
    "lng": 121.598807,
    "estimated": false,
    "isPrimary488": true,
    "group": "主要488名單",
    "delta": -0.0009887307692366676,
    "subGroup": "主要名單-西區",
    "region": "界線以西",
    "color": "紅色",
    "colorHex": "#e53e3e"
  },
  {
    "id": "store-498",
    "name": "花蓮縣新象社區交流協會",
    "address": "花蓮縣花蓮市鎮國街15號",
    "lat": 23.975904,
    "lng": 121.60129,
    "estimated": false,
    "isPrimary488": true,
    "group": "主要488名單",
    "delta": 0.0022047692307722855,
    "subGroup": "主要名單-東區",
    "region": "界線以東",
    "color": "綠色",
    "colorHex": "#257a4e"
  },
  {
    "id": "store-499",
    "name": "社團法人花蓮縣新象社區交流協會",
    "address": "花蓮縣花蓮市鎮國街15號2樓",
    "lat": 23.975904,
    "lng": 121.60129,
    "estimated": false,
    "isPrimary488": true,
    "group": "主要488名單",
    "delta": 0.0022047692307722855,
    "subGroup": "主要名單-東區",
    "region": "界線以東",
    "color": "綠色",
    "colorHex": "#257a4e"
  },
  {
    "id": "store-500",
    "name": "陳雨晞早安山丘花市和平店",
    "address": "花蓮縣花蓮市鎮國街1號",
    "lat": 23.976137,
    "lng": 121.601568,
    "estimated": false,
    "isPrimary488": false,
    "group": "非488名單(另色標記)",
    "subGroup": "非488名單",
    "region": "非488店家",
    "color": "黃色",
    "colorHex": "#f59e0b"
  }
];
}
