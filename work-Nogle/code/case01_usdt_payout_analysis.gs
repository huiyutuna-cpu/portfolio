// ============================================================
// Case 01 — USDT Payout Profit & Capital Flow Analysis
// Google Apps Script | Author: Jas Yeh | Nogle, 2023–present
//
// SETUP: Replace the three spreadsheet ID placeholders below
// with your actual IDs before running.
// ============================================================

var CONFIG = {
  ANALYSIS_RESULT_SHEET_ID: "YOUR_ANALYSIS_RESULT_SPREADSHEET_ID",
  SOURCE_SHEET_ID:          "YOUR_SOURCE_USDT_PAYOUT_SPREADSHEET_ID",
  SOURCE_SHEET_NAME:        "HTPay報表_USDT下發",
  PAYIN_PAYOUT_SHEET_ID:    "YOUR_PAYIN_PAYOUT_SPREADSHEET_ID"
};

// ============================================================
// SYSTEM ENTRY POINTS
// ============================================================

function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('📊 收益自動化報表')
    .addItem('手動更新：下發U收益分析', 'runFullAutomation')
    .addItem('手動更新：商戶收付與回U綜合佔比', 'runComprehensiveAnalysis')
    .addToUi();

  try { runFullAutomation(); } catch(e) {}
  try { runComprehensiveAnalysis(); } catch(e) {}
}

// Web app entry point — loads the Chart.js dashboard (index.html)
function doGet() {
  return HtmlService.createHtmlOutputFromFile('index')
    .setTitle('📊 月度收益數據動態儀表板')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// ============================================================
// SYSTEM 1 — USDT PAYOUT PROFIT ANALYSIS
// ============================================================

function runFullAutomation() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var targetSheet = ss.getActiveSheet();
  updateMonthlyAnalysis(targetSheet);
  formatReportLayout(targetSheet);
}

// Called by the web dashboard frontend to fetch chart data as JSON
function getDashboardData() {
  var ss = SpreadsheetApp.openById(CONFIG.ANALYSIS_RESULT_SHEET_ID);
  var sheet = ss.getSheets()[0];

  var lastRow = sheet.getLastRow();
  if (lastRow <= 1) return [];

  var data = sheet.getRange(1, 1, lastRow, 9).getValues();
  var jsonData = [];

  for (var i = 1; i < data.length; i++) {
    var row = data[i];
    if (!row[0] || String(row[0]).trim() === "" || !row[1]) continue;
    jsonData.push({
      month:    String(row[0]).trim(),
      merchant: String(row[1]).trim(),
      currency: String(row[2]).trim(),
      count:    Number(row[3]) || 0,
      diffSum:  Number(row[4]) || 0,
      amountSum:Number(row[5]) || 0,
      rate3rd:  Number(row[6]) || 0,
      rate:     Number(row[7]) || 0,
      rateDiff: Number(row[8]) || 0
    });
  }
  return jsonData;
}

// Core: pulls source data, calculates rate spread per merchant×currency×month,
// writes incrementally (closed months are locked and never re-processed)
function updateMonthlyAnalysis(targetSheet) {
  var sourceSs;
  try {
    sourceSs = SpreadsheetApp.openById(CONFIG.SOURCE_SHEET_ID);
  } catch(e) { return; }

  var sourceSheet = sourceSs.getSheetByName(CONFIG.SOURCE_SHEET_NAME);
  if (!sourceSheet) return;

  var sourceData = sourceSheet.getDataRange().getValues();
  if (sourceData.length <= 1) return;

  var now = new Date();
  var currentMonthStr = Utilities.formatDate(now, Session.getScriptTimeZone(), "yyyy-MM");

  // Build closed-month lock index from existing target data
  var targetData = targetSheet.getDataRange().getValues();
  var closedKeys = {};
  for (var i = 1; i < targetData.length; i++) {
    var row = targetData[i];
    var rowMonth = row[0];
    var rowMerchant = String(row[1]).trim();
    var rowCurrency = String(row[2]).trim();
    if (rowMonth && rowMerchant && rowCurrency && rowMonth < currentMonthStr) {
      closedKeys[rowMonth + "_" + rowMerchant + "_" + rowCurrency] = true;
    }
  }

  // Column index mapping (0-based) for source sheet
  var idxDate = 0, idxMerchant = 1, idxRate3rd = 3, idxRate = 4,
      idxCurrency = 5, idxAmount = 6, idxDiff = 9;

  var summary = {};
  var maxDateInSource = new Date(0);

  for (var j = 1; j < sourceData.length; j++) {
    var sRow = sourceData[j];
    var rawDate = sRow[idxDate];
    if (!rawDate) continue;

    var dateObj = new Date(rawDate);
    if (isNaN(dateObj.getTime())) continue;
    if (dateObj > maxDateInSource) maxDateInSource = dateObj;

    var monthStr = Utilities.formatDate(dateObj, Session.getScriptTimeZone(), "yyyy-MM");
    var merchant = String(sRow[idxMerchant]).trim();
    var currency = String(sRow[idxCurrency]).trim();
    if (!merchant || !currency) continue;

    var dataKey = monthStr + "_" + merchant + "_" + currency;
    if (closedKeys[dataKey]) continue;

    if (!summary[dataKey]) {
      summary[dataKey] = {
        month: monthStr, merchant: merchant, currency: currency,
        count: 0, diffSum: 0, amountSum: 0,
        rate3rdSum: 0, rate3rdCount: 0, rateSum: 0, rateCount: 0
      };
    }

    summary[dataKey].count     += 1;
    summary[dataKey].diffSum   += Number(sRow[idxDiff]) || 0;
    summary[dataKey].amountSum += Number(sRow[idxAmount]) || 0;

    var r3rd = Number(sRow[idxRate3rd]);
    if (!isNaN(r3rd) && r3rd > 0) {
      summary[dataKey].rate3rdSum += r3rd;
      summary[dataKey].rate3rdCount += 1;
    }
    var r = Number(sRow[idxRate]);
    if (!isNaN(r) && r > 0) {
      summary[dataKey].rateSum += r;
      summary[dataKey].rateCount += 1;
    }
  }

  var rowsToWrite = [];
  for (var k in summary) {
    var item = summary[k];
    var avgRate3rd = item.rate3rdCount > 0 ? (item.rate3rdSum / item.rate3rdCount) : 0;
    var avgRate    = item.rateCount > 0    ? (item.rateSum    / item.rateCount)    : 0;
    rowsToWrite.push([
      item.month, item.merchant, item.currency,
      item.count, item.diffSum, item.amountSum,
      avgRate3rd, avgRate, avgRate - avgRate3rd
    ]);
  }

  rowsToWrite.sort(function(a, b) {
    if (a[0] !== b[0]) return a[0].localeCompare(b[0]);
    if (a[2] !== b[2]) return a[2].localeCompare(b[2]);
    return a[1].localeCompare(b[1]);
  });

  // Remove current-month rows before re-writing
  var freshTargetData = targetSheet.getDataRange().getValues();
  for (var m = freshTargetData.length - 1; m >= 1; m--) {
    var tMonth = freshTargetData[m][0];
    if (tMonth === currentMonthStr ||
        summary[tMonth + "_" + String(freshTargetData[m][1]).trim() + "_" + String(freshTargetData[m][2]).trim()]) {
      targetSheet.deleteRow(m + 1);
    }
  }

  if (rowsToWrite.length > 0) {
    var lastRow = targetSheet.getLastRow();
    if (lastRow === 0 || (lastRow === 1 && targetSheet.getRange(1, 1).getValue() === "")) {
      targetSheet.getRange(1, 1, 1, 9).setValues([[
        "年月", "分站号(商戶)", "提现币别", "提单笔数",
        "差额加总", "提现金额加总", "平均三方汇率", "平均汇率", "汇率差额"
      ]]);
      lastRow = 1;
    }
    targetSheet.getRange(lastRow + 1, 1, rowsToWrite.length, 9).setValues(rowsToWrite);
  }

  targetSheet.getRange("O1").setValue("最後整理執行時間").setFontWeight("bold").setBackground("#F3F3F3");
  targetSheet.getRange("P1").setValue(Utilities.formatDate(now, Session.getScriptTimeZone(), "yyyy-MM-dd HH:mm:ss"));
  targetSheet.getRange("O2").setValue("原始資料截至時間").setFontWeight("bold").setBackground("#F3F3F3");
  targetSheet.getRange("P2").setValue(
    maxDateInSource.getTime() > 0
      ? Utilities.formatDate(maxDateInSource, Session.getScriptTimeZone(), "yyyy-MM-dd HH:mm:ss")
      : "無最新資料"
  );
}

// Zebra-stripe formatting + column auto-resize
function formatReportLayout(targetSheet) {
  var finalData = targetSheet.getDataRange().getValues();
  if (finalData.length <= 1) return;

  targetSheet.getRange(1, 1, 1, 9)
    .setFontWeight("bold").setBackground("#34495E")
    .setFontColor("#FFFFFF").setHorizontalAlignment("center");

  var totalRows = finalData.length - 1;
  targetSheet.getRange(2, 4, totalRows, 1).setNumberFormat("#,##0");
  targetSheet.getRange(2, 5, totalRows, 5).setNumberFormat("#,##0.00");
  targetSheet.getRange(2, 1, totalRows, 9).setBackground(null).setBorder(false, false, false, false, false, false);

  var isColoredBlock = false;
  for (var r = 1; r < finalData.length; r++) {
    if (isColoredBlock) targetSheet.getRange(r + 1, 1, 1, 9).setBackground("#F4F7FA");
    if (r + 1 < finalData.length && finalData[r][0] !== finalData[r + 1][0]) {
      isColoredBlock = !isColoredBlock;
    }
  }
  for (var c = 1; c <= 9; c++) { targetSheet.autoResizeColumn(c); }
  targetSheet.autoResizeColumn(15);
  targetSheet.autoResizeColumn(16);
}

// ============================================================
// SYSTEM 2 — MERCHANT PAYMENT RATIO ANALYSIS
// ============================================================

function cleanStr(str) {
  if (!str) return "";
  return String(str).replace(/\s+/g, '').replace(/[\/\—\–\＿\_]/g, '-').trim();
}

// Normalizes any date format to "YYYY-MM"
function parseToYearMonth(dateVal) {
  if (!dateVal) return "";
  if (dateVal instanceof Date) {
    try {
      return Utilities.formatDate(dateVal, Session.getScriptTimeZone(), "yyyy-MM");
    } catch(e) {
      var y = dateVal.getFullYear();
      var m = ("0" + (dateVal.getMonth() + 1)).slice(-2);
      return y + "-" + m;
    }
  }
  var str = cleanStr(dateVal);
  if (str.indexOf("/") !== -1 || str.indexOf("-") !== -1) {
    var parts = str.split(/[-\/]/);
    if (parts.length >= 2) {
      var y = parts[0];
      var m = parts[1].length === 1 ? "0" + parts[1] : parts[1];
      return y + "-" + m;
    }
  }
  if (str.length > 7) return str.substring(0, 7);
  return str;
}

// Cross-references Payin/Payout sheets to calculate:
//   depositAmount = Payin - Payout - U_Return
//   depositRatio  = depositAmount / Payin  (capital holding ratio)
function runComprehensiveAnalysis() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  var targetSheetName = "商戶收付與回U綜合佔比";
  var reportSheet = ss.getSheetByName(targetSheetName) || ss.insertSheet(targetSheetName);

  // Locate the USDT analysis sheet (sheet immediately before this one)
  var sheets = ss.getSheets();
  var uSheet = null;
  for (var idx = 0; idx < sheets.length; idx++) {
    if (sheets[idx].getName() === targetSheetName) {
      if (idx > 0) uSheet = sheets[idx - 1];
      break;
    }
  }
  if (!uSheet) uSheet = sheets[0];

  // Step A: cache U-return amounts from System 1 output
  var uData = uSheet.getDataRange().getValues();
  var uMap = {};
  var allPossibleKeys = new Set();
  var uOriginalNames = {};

  for (var i = 1; i < uData.length; i++) {
    var uRow = uData[i];
    if (!uRow[0] || String(uRow[0]).trim() === "" || String(uRow[0]).trim() === "年月") continue;
    var monthStr = parseToYearMonth(uRow[0]);
    var rawMerchant = String(uRow[1]).trim();
    var merchant = rawMerchant.toLowerCase();
    var currency = cleanStr(uRow[2]).toUpperCase();
    var uAmount = Number(uRow[5]) || 0;
    if (monthStr && merchant && currency && monthStr.length === 7) {
      var combKey = monthStr + "_" + merchant + "_" + currency;
      uMap[combKey] = uAmount;
      allPossibleKeys.add(combKey);
      uOriginalNames[merchant] = rawMerchant;
    }
  }

  // Build closed-month lock index
  var existingData = reportSheet.getDataRange().getValues();
  var closedMonths = {};
  var now = new Date();
  var currentMonthStr = Utilities.formatDate(now, Session.getScriptTimeZone(), "yyyy-MM");

  for (var j = 1; j < existingData.length; j++) {
    var exMonth = parseToYearMonth(existingData[j][0]);
    if (exMonth && exMonth < currentMonthStr && exMonth.length === 7) {
      closedMonths[exMonth] = true;
    }
  }

  // Step B: read all Payin/Payout sheets from source spreadsheet
  var sourceSs;
  try {
    sourceSs = SpreadsheetApp.openById(CONFIG.PAYIN_PAYOUT_SHEET_ID);
  } catch(e) {
    SpreadsheetApp.getUi().alert("❌ 無法讀取代收付試算表，請確認連結權限！");
    return;
  }

  var summary = {};

  sourceSs.getSheets().forEach(function(sheet) {
    var sheetName = sheet.getName();
    var isPayin  = sheetName.indexOf("-Payin")  !== -1;
    var isPayout = sheetName.indexOf("-Payout") !== -1;
    if (!isPayin && !isPayout) return;

    var monthStr = parseToYearMonth(sheetName.split("-")[0]);
    if (closedMonths[monthStr]) return;

    var sData = sheet.getDataRange().getValues();
    if (sData.length <= 1) return;

    for (var r = 1; r < sData.length; r++) {
      var rawMer   = String(sData[r][2]).trim();
      var merchant = cleanStr(rawMer).toLowerCase();
      var currency = cleanStr(sData[r][isPayin ? 5 : 4]).toUpperCase();
      var amount   = Number(sData[r][6]) || 0;
      if (!merchant || !currency || rawMer === "undefined" || rawMer === "") continue;

      var key = monthStr + "_" + merchant + "_" + currency;
      if (!summary[key]) {
        summary[key] = { month: monthStr, merchant: rawMer, currency: currency, payin: 0, payout: 0 };
      }
      if (isPayin)  summary[key].payin  += amount;
      if (isPayout) summary[key].payout += amount;
    }
  });

  // Step C: fill in merchants that appear only in U-return data (no Payin/Payout record)
  allPossibleKeys.forEach(function(uKey) {
    var parts = uKey.split("_");
    if (closedMonths[parts[0]]) return;
    if (!summary[uKey]) {
      summary[uKey] = {
        month: parts[0],
        merchant: uOriginalNames[parts[1]] || parts[1],
        currency: parts[2],
        payin: 0, payout: 0
      };
    }
  });

  // Step D: calculate final metrics and build output rows
  var rowsToWrite = [];
  for (var k in summary) {
    var item = summary[k];
    var uAmountSum    = uMap[k] || 0;
    var payoutRatio   = item.payin > 0 ? (item.payout    / item.payin) : 0;
    var uRatio        = item.payin > 0 ? (uAmountSum     / item.payin) : 0;
    var depositAmount = item.payin - item.payout - uAmountSum;           // 存放金額
    var depositRatio  = item.payin > 0 ? (depositAmount  / item.payin) : 0; // 資金存放率

    if (item.payin === 0 && item.payout === 0 && uAmountSum === 0) continue;

    rowsToWrite.push([
      item.month, item.merchant, item.currency,
      item.payin, item.payout, payoutRatio,
      uAmountSum, uRatio,
      depositAmount, depositRatio
    ]);
  }

  rowsToWrite.sort(function(a, b) {
    if (a[0] !== b[0]) return a[0].localeCompare(b[0]);
    if (a[2] !== b[2]) return a[2].localeCompare(b[2]);
    return a[1].localeCompare(b[1]);
  });

  // Clear current-month rows before re-writing
  var freshTargetData = reportSheet.getDataRange().getValues();
  for (var m = freshTargetData.length - 1; m >= 1; m--) {
    var tMonth = parseToYearMonth(freshTargetData[m][0]);
    if (tMonth === currentMonthStr || closedMonths[tMonth] === undefined) {
      reportSheet.deleteRow(m + 1);
    }
  }

  if (rowsToWrite.length > 0) {
    var lastRow = reportSheet.getLastRow();
    if (lastRow === 0 || (lastRow === 1 && reportSheet.getRange(1, 1).getValue() === "")) {
      reportSheet.getRange(1, 1, 1, 10).setValues([[
        "年月", "分站号(商戶)", "幣別", "總代收金額", "總代付金額",
        "代付佔比(%)", "總回U金額", "回U佔比(%)", "存放金額", "資金存放率(%)"
      ]]);
      lastRow = 1;
    }
    reportSheet.getRange(lastRow + 1, 1, rowsToWrite.length, 10).setValues(rowsToWrite);

    var totalRows = reportSheet.getLastRow() - 1;
    reportSheet.getRange(1, 1, 1, 10)
      .setFontWeight("bold").setBackground("#2C3E50")
      .setFontColor("#FFFFFF").setHorizontalAlignment("center");

    reportSheet.getRange(2, 4, totalRows, 2).setNumberFormat("#,##0.00");
    reportSheet.getRange(2, 7, totalRows, 1).setNumberFormat("#,##0.00");
    reportSheet.getRange(2, 9, totalRows, 1).setNumberFormat("#,##0.00");
    reportSheet.getRange(2, 6, totalRows, 1).setNumberFormat("0.00%");
    reportSheet.getRange(2, 8, totalRows, 1).setNumberFormat("0.00%");
    reportSheet.getRange(2, 10, totalRows, 1).setNumberFormat("0.00%");

    reportSheet.getRange(2, 1, totalRows, 10).setBackground(null);
    var isColoredBlock = false;
    var finalData = reportSheet.getDataRange().getValues();
    for (var r = 1; r < finalData.length; r++) {
      if (isColoredBlock) reportSheet.getRange(r + 1, 1, 1, 10).setBackground("#F8F9FA");
      if (r + 1 < finalData.length &&
          parseToYearMonth(finalData[r][0]) !== parseToYearMonth(finalData[r + 1][0])) {
        isColoredBlock = !isColoredBlock;
      }
    }
  }

  for (var c = 1; c <= 10; c++) { reportSheet.autoResizeColumn(c); }
}
