/**
 * Case 03 — Daily OPS Data Analysis System
 * Module: Source Sheet Analysis (代收 Payin + 代付 Payout)
 *
 * Functions in this file:
 *   1. sendProcessingOrdersToTG_Payin   — 上分訂單（處理中/新申請）推 TG
 *   2. sendProcessingOrdersToTG_Payout  — 回分訂單（處理中/新申請）推 TG
 *   3. checkAllSheetsVolumeDrop         — 各幣別跑量下跌 >10% 告警
 *   4. sendSlowOrdersToTG_ByAnalysisA1  — VND 慢單 >10min >50筆 告警
 *   5. analyzePSPGaps_AllCurrencies     — 代付 PSP 兩單間隔 >100min（吊單率）偵測
 *   6. 分組分析_商戶幣別PSP統計完整      — 法幣代收 PSP 統計 → 寫入分析表
 *   7. 分組分析_商戶幣別PSP統計虛擬幣    — 虛擬幣代收 PSP 統計
 *   8. 執行所有商戶代付統計              — 代付統計入口（法幣 + 虛擬幣）
 *   9. 核心代付處理邏輯                 — 代付統計核心邏輯（共用函式，customFilter 分流）
 */

// ============================================================
// CONFIG — 填入你自己的 ID
// ============================================================
const PAYIN_SOURCE_SS_ID  = "YOUR_PAYIN_SOURCE_SS_ID";   // 上分訂單分析試算表 ID
const PAYOUT_SOURCE_SS_ID = "YOUR_PAYOUT_SOURCE_SS_ID";  // 回分訂單分析試算表 ID
const FIAT_ANALYSIS_SS_ID   = "YOUR_FIAT_ANALYSIS_SS_ID";   // 法幣分析試算表 ID
const CRYPTO_ANALYSIS_SS_ID = "YOUR_CRYPTO_ANALYSIS_SS_ID"; // 虛擬幣分析試算表 ID

const TG_BOT_TOKEN    = "YOUR_BOT_TOKEN";
const TG_CHAT_ID      = "YOUR_CHAT_ID";
const TG_CHAT_ID_VOL  = "YOUR_CHAT_ID_VOL_MONITOR";  // 跑量監控群組

const ALLOWED_FIAT_SHEETS = ["CNY", "VND", "THB", "IDR", "BDT", "INR-Worldpay", "PHP"];

// ============================================================
// 1. 上分訂單報告（處理中 & 新申請）→ TG
// ============================================================
function sendProcessingOrdersToTG_Payin() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const analysisSheet = ss.getSheetByName("Analysis");
  const sourceSheetName = analysisSheet.getRange("A1").getValue().toString().trim();
  if (!sourceSheetName) { SpreadsheetApp.getUi().alert("請在 Analysis!A1 輸入來源工作表名稱。"); return; }

  const sourceSheet = ss.getSheetByName(sourceSheetName);
  if (!sourceSheet) { SpreadsheetApp.getUi().alert(`找不到名稱為 "${sourceSheetName}" 的工作表`); return; }

  const data = sourceSheet.getDataRange().getValues();
  const processingList = [], newApplyList = [];

  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const stationName = row[2];   // C欄
    const txId        = row[7];   // H欄
    const currency    = row[10];  // K欄
    const status      = row[18];  // S欄
    if (!status) continue;
    const s = status.toString().trim();
    if (s === "处理中") processingList.push(`${stationName} - ${currency} - ${txId}`);
    if (s === "新申请") newApplyList.push(`${stationName} - ${currency} - ${txId}`);
  }

  if (processingList.length === 0 && newApplyList.length === 0) {
    SpreadsheetApp.getUi().alert("今天沒有處理中或新申請的訂單"); return;
  }

  let message = `📄 上分訂單報告（處理中 & 新申請）\n來源表：${sourceSheetName}\n\n`;
  message += `以下還在「處理中」訂單：\n`;
  message += processingList.length > 0 ? processingList.map((item, i) => `${i+1}. ${item}`).join("\n") : "無處理中";
  message += `\n\n以下還在「新申请」訂單：\n`;
  message += newApplyList.length > 0 ? newApplyList.map((item, i) => `${i+1}. ${item}`).join("\n") : "無新申请";

  sendTelegramMessage_(TG_BOT_TOKEN, TG_CHAT_ID, message);
  SpreadsheetApp.getUi().alert(`已發送到 Telegram：\n處理中 ${processingList.length} 筆、新申請 ${newApplyList.length} 筆`);
}

// ============================================================
// 2. 回分訂單報告（處理中 & 新申請）→ TG
// ============================================================
function sendProcessingOrdersToTG_Payout() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const analysisSheet = ss.getSheetByName("Analysis");
  const sourceSheetName = analysisSheet.getRange("A1").getValue().toString().trim();
  if (!sourceSheetName) { SpreadsheetApp.getUi().alert("請在 Analysis!A1 輸入來源工作表名稱。"); return; }

  const sourceSheet = ss.getSheetByName(sourceSheetName);
  if (!sourceSheet) { SpreadsheetApp.getUi().alert(`找不到名稱為 "${sourceSheetName}" 的工作表`); return; }

  const data = sourceSheet.getDataRange().getValues();
  const processingList = [], newApplyList = [];

  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const stationName = row[2];   // C欄
    const txId        = row[7];   // H欄
    const currency    = row[11];  // L欄（代付幣別欄位不同）
    const status      = row[18];  // S欄
    if (!status) continue;
    const s = status.toString().trim();
    if (s === "处理中") processingList.push(`${stationName} - ${currency} - ${txId}`);
    if (s === "新申请") newApplyList.push(`${stationName} - ${currency} - ${txId}`);
  }

  if (processingList.length === 0 && newApplyList.length === 0) {
    SpreadsheetApp.getUi().alert("今天沒有處理中或新申請的訂單"); return;
  }

  let message = `📄 回分訂單報告（處理中 & 新申請）\n來源表：${sourceSheetName}\n\n`;
  message += `以下還在「處理中」訂單：\n`;
  message += processingList.length > 0 ? processingList.map((item, i) => `${i+1}. ${item}`).join("\n") : "無處理中";
  message += `\n\n以下還在「新申请」訂單：\n`;
  message += newApplyList.length > 0 ? newApplyList.map((item, i) => `${i+1}. ${item}`).join("\n") : "無新申请";

  sendTelegramMessage_(TG_BOT_TOKEN, TG_CHAT_ID, message);
  SpreadsheetApp.getUi().alert(`已發送到 Telegram：\n處理中 ${processingList.length} 筆、新申請 ${newApplyList.length} 筆`);
}

// ============================================================
// 3. 各幣別跑量下跌 >10% 告警 → TG
// ============================================================
function checkAllSheetsVolumeDrop() {
  const sheets = SpreadsheetApp.getActiveSpreadsheet().getSheets();
  const allMessages = [];
  let comparisonDates = null;

  sheets.forEach(sheet => {
    if (!sheet) return;
    if (!ALLOWED_FIAT_SHEETS.includes(sheet.getName())) return;
    try {
      const { messages, dates } = checkVolumeInSheetWithDate_(sheet);
      if (!comparisonDates && dates) comparisonDates = dates;
      if (messages.length > 0) allMessages.push(`💱 *${sheet.getName()}*\n` + messages.join("\n\n"));
    } catch (e) { Logger.log(`⚠️ 處理工作表 ${sheet.getName()} 發生錯誤: ${e}`); }
  });

  if (allMessages.length > 0) {
    const dateHeader = comparisonDates
      ? `📅 比較日期：${formatDate_(comparisonDates.yesterday)} vs ${formatDate_(comparisonDates.today)}\n\n`
      : "";
    sendTelegramMessage_(TG_BOT_TOKEN, TG_CHAT_ID_VOL, dateHeader + allMessages.join("\n\n────────────\n\n"));
  }
}

function checkVolumeInSheetWithDate_(sheet) {
  const data = sheet.getDataRange().getValues();
  if (data.length < 3 || data[0].length < 4) return { messages: [], dates: null };

  const messages = [];
  const lastCol = data[0].length;
  const dateColumns = [];

  for (let col = 2; col <= lastCol; col += 3) {
    const displayText = sheet.getRange(1, col).getDisplayValue().trim();
    if (!displayText) continue;
    const dateObj = new Date(displayText.substring(0, 10));
    if (!isNaN(dateObj.getTime())) dateColumns.push({ date: dateObj, colIndex: col });
  }

  if (dateColumns.length < 2) return { messages: [], dates: null };

  const now = new Date();
  const validDates = dateColumns.filter(dc => dc.date <= now);
  if (validDates.length < 2) return { messages: [], dates: null };

  validDates.sort((a, b) => a.date - b.date);
  const yesterday = validDates[validDates.length - 3];
  const today     = validDates[validDates.length - 2];

  const yesterdayStr = Utilities.formatDate(yesterday.date, Session.getScriptTimeZone(), "yyyy-MM-dd");
  const todayStr     = Utilities.formatDate(today.date,     Session.getScriptTimeZone(), "yyyy-MM-dd");
  const skipKeywords = ["總和","總計","合計","總數","總覽","total","sum"];

  for (let i = 2; i < data.length; i++) {
    const merchant = data[i][0];
    if (!merchant || typeof merchant !== "string" || merchant.trim() === "") continue;
    if (skipKeywords.some(k => merchant.toLowerCase().includes(k.toLowerCase()))) break;

    const volumeYesterday = parseFloat(data[i][yesterday.colIndex]) || 0;
    const volumeToday     = parseFloat(data[i][today.colIndex])     || 0;
    if (volumeYesterday === 0) continue;

    const dropPercent = ((volumeYesterday - volumeToday) / volumeYesterday) * 100;
    if (dropPercent > 10) {
      messages.push(
        `⚠️ 商戶 *${merchant}* 跑量下跌 *${dropPercent.toFixed(2)}%*\n` +
        `📉 ${yesterdayStr}: ${formatNumber_(volumeYesterday)}\n` +
        `📉 ${todayStr}: ${formatNumber_(volumeToday)}`
      );
    }
  }
  return { messages, dates: { yesterday: yesterday.date, today: today.date } };
}

// ============================================================
// 4. VND 慢單告警（>10min & >50筆 才推 TG）
// ============================================================
const TARGET_CURRENCY     = "VND";
const THRESHOLD_MINUTES   = 10;
const MIN_COUNT_TO_NOTIFY = 50;
const MAX_ITEMS_IN_MESSAGE = 80;

function sendSlowOrdersToTG_ByAnalysisA1() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const analysisSheet = ss.getSheetByName("Analysis");
  if (!analysisSheet) { SpreadsheetApp.getUi().alert("找不到 Analysis 工作表"); return; }
  const sourceSheetName = String(analysisSheet.getRange("A1").getValue() || "").trim();
  if (!sourceSheetName) { SpreadsheetApp.getUi().alert("請在 Analysis!A1 輸入來源工作表名稱。"); return; }

  const sourceSheet = ss.getSheetByName(sourceSheetName);
  if (!sourceSheet) { SpreadsheetApp.getUi().alert(`找不到名稱為 "${sourceSheetName}" 的工作表`); return; }

  const data = sourceSheet.getDataRange().getValues();
  const flagged = [];

  for (let i = 1; i < data.length; i++) {
    const row      = data[i];
    const merchant = safeTrim_(row[2]);   // C
    const psp      = safeTrim_(row[6]);   // G
    const orderId  = safeTrim_(row[7]);   // H
    const currency = safeTrim_(row[11]);  // L
    const duration = row[25];             // Z

    if (currency !== TARGET_CURRENCY) continue;
    const minutes = parseDurationToMinutes_(duration);
    if (minutes === null || minutes <= THRESHOLD_MINUTES) continue;
    flagged.push({ currency, merchant: merchant || "-", psp: psp || "-", orderId: orderId || "-", minutes });
  }

  if (flagged.length <= MIN_COUNT_TO_NOTIFY) {
    SpreadsheetApp.getUi().alert(
      `來源表：${sourceSheetName}\n符合條件筆數：${flagged.length}\n未超過門檻（> ${MIN_COUNT_TO_NOTIFY}），不發 TG。`
    );
    return;
  }

  const now = new Date();
  const header =
    `🚨 慢單告警\n來源表：${sourceSheetName}\n` +
    `時間：${Utilities.formatDate(now, Session.getScriptTimeZone(), "yyyy-MM-dd HH:mm:ss")}\n` +
    `條件：${TARGET_CURRENCY} 且 Z > ${THRESHOLD_MINUTES} 分鐘，且總數 > ${MIN_COUNT_TO_NOTIFY}\n` +
    `目前超標總數：${flagged.length}\n\n`;

  const list = flagged.slice(0, MAX_ITEMS_IN_MESSAGE)
    .map((x, idx) => `${idx+1}. ${x.currency} | 商戶:${x.merchant} | PSP:${x.psp} | 訂單:${x.orderId}`)
    .join("\n");
  const tail = flagged.length > MAX_ITEMS_IN_MESSAGE
    ? `\n\n…其餘 ${flagged.length - MAX_ITEMS_IN_MESSAGE} 筆未列出` : "";

  sendTelegramMessage_(TG_BOT_TOKEN, TG_CHAT_ID, header + list + tail);
  SpreadsheetApp.getUi().alert(`✅ 已發送 TG\n慢單筆數：${flagged.length}`);
}

// ============================================================
// 5. 代付 PSP 吊單率偵測（兩單間隔 >100min）
// ============================================================
function analyzePSPGaps_AllCurrencies_WithTxnID() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const outputSheet   = ss.getSheetByName("代付超時");
  const analysisSheet = ss.getSheetByName("Analysis");
  const sourceSheetName = analysisSheet.getRange("A1").getValue().toString().trim();
  if (!sourceSheetName) { SpreadsheetApp.getUi().alert("請在 Analysis!A1 輸入來源工作表名稱。"); return; }

  const sourceSheet = ss.getSheetByName(sourceSheetName);
  if (!sourceSheet) { SpreadsheetApp.getUi().alert("找不到來源工作表"); return; }

  // 記住上次來源表名，換表時自動清空結果
  const scriptProperties = PropertiesService.getDocumentProperties();
  const previousSourceName = scriptProperties.getProperty("LAST_SOURCE_SHEET_NAME_PAY");
  if (previousSourceName !== sourceSheetName) {
    const lastRow = outputSheet.getLastRow();
    if (lastRow > 1) outputSheet.getRange(2, 1, lastRow - 1, outputSheet.getLastColumn()).clearContent();
    scriptProperties.setProperty("LAST_SOURCE_SHEET_NAME_PAY", sourceSheetName);
  }

  const data  = sourceSheet.getDataRange().getValues();
  const rows  = data.slice(1);
  const parsed = rows
    .map((r, i) => ({ rowIndex: i+2, time: r[4], psp: r[6], txn: r[7], currency: r[11] }))
    .filter(r => r.time instanceof Date && r.psp && r.psp !== "None" && r.psp !== "BTSEPAY")
    .sort((a, b) => a.time - b.time);

  const result   = [];
  const lastSeen = {};

  for (const entry of parsed) {
    const key  = `${entry.psp}|${entry.currency}`;
    const prev = lastSeen[key];
    if (prev) {
      const gapMinutes = (entry.time - prev.time) / (1000 * 100);
      if (gapMinutes > 100) {
        result.push([entry.psp, prev.time, entry.time, Math.round(gapMinutes), entry.currency, prev.txn]);
      }
    }
    lastSeen[key] = { time: entry.time, txn: entry.txn };
  }

  result.sort((a, b) => a[1] - b[1]);
  const header = ["PSP名稱", "前一筆時間", "本筆時間", "間隔（分鐘）", "幣別", "前一筆交易編號"];
  outputSheet.getRange(2, 1, 1, header.length).setValues([header]);
  if (result.length > 0) outputSheet.getRange(3, 1, result.length, 6).setValues(result);
  else outputSheet.getRange(3, 1).setValue("無間隔超過1小時的資料");
}

// ============================================================
// 6 & 7. 代收 PSP 統計（法幣 / 虛擬幣）
// ============================================================

/** 法幣代收：統計商戶 × PSP × 類型，寫入法幣分析表 */
function 分組分析_商戶幣別PSP統計完整() {
  _runPayinPSPStats_(FIAT_ANALYSIS_SS_ID, row => {
    const psp = (row[5] || "").toString().trim();
    const type = (row[4] || "").toString().trim().replace(/\s/g, "");
    const allowEmptyPSP = ["上分调帐加钱","上分调帐扣钱","提现"].includes(type);
    if (!allowEmptyPSP) {
      if (!psp) return false;
      if (psp.toLowerCase() === "btsepay") return false;
    }
    return true;
  });
}

/** 虛擬幣代收：統計商戶 × PSP × 類型，寫入虛擬幣分析表 */
function 分組分析_商戶幣別PSP統計虛擬幣() {
  _runPayinPSPStats_(CRYPTO_ANALYSIS_SS_ID, row => {
    const psp  = (row[5] || "").toString().trim();
    const type = (row[4] || "").toString().trim();
    const allowEmptyPSP = ["上分调帐加钱","上分调帐扣钱","提现"].includes(type);
    if (!allowEmptyPSP && (!psp || psp.includes("银行转帐"))) return false;
    return true;
  });
}

/** 代收 PSP 統計核心邏輯（法幣/虛擬幣共用，customFilter 分流） */
function _runPayinPSPStats_(targetSpreadsheetId, customFilter) {
  const sourceFile     = SpreadsheetApp.openById(PAYIN_SOURCE_SS_ID);
  const analysisSheet  = sourceFile.getSheetByName("Analysis");
  const sourceSheetName = analysisSheet.getRange("A1").getValues()[0][0];
  if (!sourceSheetName) { SpreadsheetApp.getUi().alert("Analysis!A1 沒有資料來源表名"); return; }

  const sourceSheet = sourceFile.getSheetByName(sourceSheetName);
  if (!sourceSheet) { SpreadsheetApp.getUi().alert("找不到來源工作表：" + sourceSheetName); return; }

  const lastRow  = sourceSheet.getLastRow();
  const lastCol  = sourceSheet.getLastColumn();
  if (lastRow <= 1) { SpreadsheetApp.getUi().alert("來源工作表沒有資料"); return; }

  const sourceData = sourceSheet.getRange(2, 1, lastRow-1, lastCol).getValues();
  const filtered   = sourceData.filter(row => {
    const merchant = (row[2] || "").toString().trim();
    const type     = (row[4] || "").toString().trim().replace(/\s/g, "");
    const currency = (row[10] || "").toString().trim();
    const amount   = row[12];
    const status   = (row[18] || "").toString().trim();
    if (!merchant || !currency || !type || amount === "" || amount == null || !status) return false;
    return customFilter(row);
  });

  const resultMap = new Map();
  const merchantCurrencyTotals = new Map();

  for (const row of filtered) {
    const type     = (row[4] || "").toString().trim().replace(/\s/g, "");
    const merchant = (row[2] || "").toString().trim();
    const psp      = (row[5] || "").toString().trim();
    const currency = (row[10] || "").toString().trim();
    const status   = (row[18] || "").toString().trim();
    const amount   = parseFloat(row[12]) || 0;
    const key      = `${merchant}||${currency}||${psp}||${type}`;
    const totalKey = `${merchant}||${currency}`;

    if (!resultMap.has(key)) resultMap.set(key, { total:0, success:0, fail:0, successAmount:0 });
    const record = resultMap.get(key);
    record.total++;
    if (status.includes("完成")) {
      record.success++;
      record.successAmount += (type === "上分调帐扣钱" ? -amount : amount);
    } else { record.fail++; }

    merchantCurrencyTotals.set(totalKey, (merchantCurrencyTotals.get(totalKey)||0)+1);
  }

  // 計算平均完成時間
  const typeOrder = { "上分":1,"回分":2,"提现":3,"上分调帐加钱":4,"上分调帐扣钱":5 };
  const output = [];
  for (const [key, stats] of resultMap.entries()) {
    const [merchant, currency, psp, type] = key.split("||");
    const totalKey  = `${merchant}||${currency}`;
    const totalCount = merchantCurrencyTotals.get(totalKey) || 0;
    const successRate = stats.total ? (stats.success/stats.total*100).toFixed(2)+"%" : "0%";
    const pspRatio    = totalCount  ? (stats.total/totalCount*100).toFixed(2)+"%" : "0%";

    const rowsForAvg = sourceData.filter(row => {
      const rMerchant = (row[2]||"").toString().trim();
      const rPsp      = (row[5]||"").toString().trim();
      const rType     = (row[4]||"").toString().trim().replace(/\s/g,"");
      const rStatus   = (row[18]||"").toString().trim();
      const rComplete = row[17];
      if (rMerchant!==merchant || rType!==type || !rStatus.includes("完成") || !rComplete) return false;
      if (psp && rPsp!==psp) return false;
      if (!psp && rPsp) return false;
      return true;
    });

    let totalSeconds = 0, validCount = 0;
    for (const r of rowsForAvg) {
      const diff = (new Date(r[17]) - new Date(r[3])) / 1000;
      if (diff >= 0) { totalSeconds += diff; validCount++; }
    }
    let avgTime = "-";
    if (validCount > 0) {
      const avg = totalSeconds/validCount;
      avgTime = `${Math.floor(avg/3600).toString().padStart(2,"0")}:${Math.floor((avg%3600)/60).toString().padStart(2,"0")}:${Math.round(avg%60).toString().padStart(2,"0")}`;
    }
    output.push([merchant, currency, type, psp, toThousands_(stats.total), toThousands_(stats.success), toThousands_(stats.fail), successRate, pspRatio, toThousands_(stats.successAmount), avgTime]);
  }

  // 寫入目標試算表
  const targetSS    = SpreadsheetApp.openById(targetSpreadsheetId);
  const outputSheet = targetSS.getSheetByName("商戶PSP統計-代收");
  if (!outputSheet) { SpreadsheetApp.getUi().alert("找不到『商戶PSP統計-代收』工作表"); return; }

  outputSheet.getRange(2,1,outputSheet.getMaxRows()-1,11).clearContent();
  outputSheet.getRange("A2").setValue(`來源資料表：${sourceSheetName}`);
  outputSheet.getRange("A3:K3").setValues([["商戶名稱","幣別","類型","PSP名稱","總筆數","成功筆數","失敗筆數","成功率","PSP佔比","成功金額","平均完成時間"]]);
  if (output.length > 0) outputSheet.getRange(4,1,output.length,11).setValues(output);

  // 寫入幣別分頁（按日期偏移欄）
  const dateMatch = sourceSheetName.match(/\d{4}-\d{2}-\d{2}/);
  if (!dateMatch) { SpreadsheetApp.getUi().alert("來源日期格式不正確"); return; }
  const day = new Date(dateMatch[0]).getDate();
  const offsetCol = (day-1)*7+4;

  output.forEach((rowArr, i) => {
    const [merchant,,type,psp,total,success,fail,rate,ratio,amount,avgTime] = rowArr;
    const typeClean = type.replace(/\s/g,"");
    const currencySheet = targetSS.getSheetByName(rowArr[1]);
    if (!currencySheet) { outputSheet.getRange(i+4,1).setBackground("#d9d9d9"); return; }

    const mPSPRange = currencySheet.getRange("A3:C").getValues();
    let targetRowIdx = mPSPRange.findIndex(([m,p,t]) =>
      (m||"").toString().trim()===merchant && (p||"").toString().trim()===psp &&
      (t||"").toString().trim().replace(/\s/g,"")===typeClean
    );

    if (targetRowIdx === -1) {
      let insertIdx = 0, foundGroup = false;
      for (let j=0; j<mPSPRange.length; j++) {
        const [m,p,t] = mPSPRange[j].map(v=>(v||"").toString().trim());
        if (m===merchant) {
          if ((typeOrder[typeClean]||99)<(typeOrder[t.replace(/\s/g,"")]||99)) { insertIdx=j; foundGroup=true; break; }
          else { insertIdx=j+1; foundGroup=true; }
        }
      }
      if (!foundGroup) insertIdx=0;
      currencySheet.insertRowBefore(insertIdx+3);
      currencySheet.getRange(insertIdx+3,1,1,3).setValues([[merchant,psp,type]]);
      targetRowIdx = insertIdx;
    }
    try {
      currencySheet.getRange(targetRowIdx+3,offsetCol,1,7).setValues([[total,success,fail,rate,ratio,amount,avgTime]]);
    } catch(e) { outputSheet.getRange(i+4,1).setBackground("#d9d9d9"); }
  });
}

// ============================================================
// 8 & 9. 代付 PSP 統計
// ============================================================

const CONFIG_PAYOUT = {
  SOURCE_SS_ID:   "YOUR_PAYOUT_SOURCE_SS_ID",
  TARGET_FULL_ID: "YOUR_FIAT_ANALYSIS_SS_ID",
  TARGET_CRYPTO_ID: "YOUR_CRYPTO_ANALYSIS_SS_ID",
  SHEET_NAME_ANALYSIS: "Analysis",
  SHEET_NAME_OUTPUT:   "商戶PSP統計-代付"
};

function 執行所有商戶代付統計() {
  核心代付處理邏輯(CONFIG_PAYOUT.TARGET_FULL_ID, row => {
    const type = (row[5]||"").toString().trim().replace(/\s/g,"");
    const psp  = (row[6]||"").toString().trim();
    const allowEmptyPSP = ["调帐加钱","调帐扣钱"].includes(type);
    if (!allowEmptyPSP && (!psp || psp.includes("银行转帐") || psp.toLowerCase()==="btsepay")) return false;
    return true;
  });
  核心代付處理邏輯(CONFIG_PAYOUT.TARGET_CRYPTO_ID, row => {
    const type = (row[5]||"").toString().trim().replace(/\s/g,"");
    const psp  = (row[6]||"").toString().trim();
    const allowEmptyPSP = ["调帐加钱","调帐扣钱"].includes(type);
    if (!allowEmptyPSP && (!psp || psp.includes("银行转帐"))) return false;
    return true;
  });
  SpreadsheetApp.getUi().alert("✅ 所有「代付」統計已同步完成！");
}

function 核心代付處理邏輯(targetId, customFilter) {
  const sourceFile    = SpreadsheetApp.openById(CONFIG_PAYOUT.SOURCE_SS_ID);
  const analysisSheet = sourceFile.getSheetByName(CONFIG_PAYOUT.SHEET_NAME_ANALYSIS);
  const sourceSheetName = analysisSheet.getRange("A1").getValue();
  if (!sourceSheetName) return;

  const sourceSheet = sourceFile.getSheetByName(sourceSheetName);
  if (!sourceSheet) return;

  const rawData = sourceSheet.getRange("A2:Z"+sourceSheet.getLastRow()).getValues();
  const resultMap = new Map(), merchantCurrencyTotals = new Map();

  rawData.filter(row => {
    const merchant = (row[2]||"").toString().trim();
    const type     = (row[5]||"").toString().trim().replace(/\s/g,"");
    const currency = (row[11]||"").toString().trim();
    const amount   = row[12];
    const status   = (row[18]||"").toString().trim();
    if (!merchant||!currency||!type||amount==null||!status) return false;
    return customFilter(row);
  }).forEach(row => {
    const merchant = (row[2]||"").toString().trim();
    const type     = (row[5]||"").toString().trim().replace(/\s/g,"");
    const psp      = (row[6]||"").toString().trim();
    const currency = (row[11]||"").toString().trim();
    const status   = (row[18]||"").toString().trim();
    const amount   = parseFloat(row[12])||0;
    const key      = `${merchant}||${currency}||${psp}||${type}`;
    const totalKey = `${merchant}||${currency}`;

    if (!resultMap.has(key)) resultMap.set(key, {total:0,success:0,fail:0,successAmount:0,totalTime:0,completeCount:0});
    const record = resultMap.get(key);
    record.total++;
    if (status.includes("完成")) {
      record.success++;
      record.successAmount += (type==="调帐扣钱"?-amount:amount);
      const diff = (new Date(row[17])-new Date(row[4]))/1000;
      if (diff>=0) { record.totalTime+=diff; record.completeCount++; }
    } else { record.fail++; }
    merchantCurrencyTotals.set(totalKey, (merchantCurrencyTotals.get(totalKey)||0)+1);
  });

  const targetSS    = SpreadsheetApp.openById(targetId);
  const outputSheet = targetSS.getSheetByName(CONFIG_PAYOUT.SHEET_NAME_OUTPUT);
  if (!outputSheet) return;

  outputSheet.getRange(2,1,Math.max(1,outputSheet.getMaxRows()-1),11).clearContent();
  outputSheet.getRange("A2").setValue(`來源資料表：${sourceSheetName}`);
  outputSheet.getRange("A3:K3").setValues([["商戶名稱","幣別","類型","PSP名稱","總筆數","成功筆數","失敗筆數","成功率","PSP佔比","成功金額","平均完成時間"]]);

  const typeOrder = {"上分":1,"回分":2,"提现":3,"调帐加钱":4,"调帐扣钱":5};
  const outputBuffer = [];

  for (const [key,stats] of resultMap.entries()) {
    const [merchant,currency,psp,type] = key.split("||");
    const totalCount = merchantCurrencyTotals.get(`${merchant}||${currency}`);
    const avgTimeStr = stats.completeCount>0 ? formatSeconds_(stats.totalTime/stats.completeCount) : "-";
    outputBuffer.push([
      merchant,currency,type,psp,
      toThousands_(stats.total),toThousands_(stats.success),toThousands_(stats.fail),
      (stats.total?(stats.success/stats.total*100).toFixed(2)+"%":"0%"),
      (totalCount?(stats.total/totalCount*100).toFixed(2)+"%":"0%"),
      toThousands_(stats.successAmount),avgTimeStr
    ]);
  }

  if (outputBuffer.length>0) outputSheet.getRange(4,1,outputBuffer.length,11).setValues(outputBuffer);

  const dateMatch = sourceSheetName.match(/\d{4}-\d{2}-\d{2}/);
  if (!dateMatch) return;
  const day = new Date(dateMatch[0]).getDate();
  const offsetCol = (day-1)*7+4;

  outputBuffer.forEach((rowArr,i) => {
    const [merchant,currency,type,psp,total,success,fail,rate,ratio,amount,avgTime] = rowArr;
    const typeClean = type.replace(/\s/g,"");
    const currencySheet = targetSS.getSheetByName(currency);
    if (!currencySheet) { outputSheet.getRange(i+4,1).setBackground("#d9d9d9"); return; }

    const mPSPRange = currencySheet.getRange("A3:C"+Math.max(3,currencySheet.getLastRow())).getValues();
    let targetRowIdx = mPSPRange.findIndex(([m,p,t])=>
      String(m).trim()===merchant&&String(p).trim()===psp&&String(t).trim().replace(/\s/g,"")===typeClean
    );
    if (targetRowIdx===-1) {
      let insertIdx = mPSPRange.findIndex(([m,p,t])=>{
        if (String(m).trim()!==merchant) return false;
        return (typeOrder[typeClean]||99)<(typeOrder[String(t).trim().replace(/\s/g,"")]||99);
      });
      if (insertIdx===-1) {
        for(let k=mPSPRange.length-1;k>=0;k--) {
          if(String(mPSPRange[k][0]).trim()===merchant){insertIdx=k+1;break;}
        }
      }
      if (insertIdx===-1) insertIdx=mPSPRange.length;
      currencySheet.insertRowBefore(insertIdx+3);
      currencySheet.getRange(insertIdx+3,1,1,3).setValues([[merchant,psp,type]]);
      targetRowIdx=insertIdx;
    }
    currencySheet.getRange(targetRowIdx+3,offsetCol,1,7).setValues([[total,success,fail,rate,ratio,amount,avgTime]]);
  });
}

// ============================================================
// Shared Helpers
// ============================================================
function sendTelegramMessage_(token, chatId, text) {
  UrlFetchApp.fetch(`https://api.telegram.org/bot${token}/sendMessage`, {
    method:"post", contentType:"application/json",
    payload: JSON.stringify({ chat_id: chatId, text, parse_mode:"Markdown" }),
    muteHttpExceptions: true
  });
}
function toThousands_(num) {
  const n=Math.round(Number(num)||0);
  return (n<0?"-":"")+Math.abs(n).toString().replace(/\B(?=(\d{3})+(?!\d))/g,",");
}
function formatSeconds_(avg) {
  return `${Math.floor(avg/3600).toString().padStart(2,"0")}:${Math.floor((avg%3600)/60).toString().padStart(2,"0")}:${Math.round(avg%60).toString().padStart(2,"0")}`;
}
function formatDate_(dateObj) {
  return Utilities.formatDate(dateObj, Session.getScriptTimeZone(), "yyyy-MM-dd");
}
function formatNumber_(num) {
  return num.toLocaleString("en-US",{minimumFractionDigits:0,maximumFractionDigits:0});
}
function parseDurationToMinutes_(v) {
  if (v===""||v===null||v===undefined) return null;
  if (typeof v==="number") return v>=0&&v<=1?v*24*60:v;
  if (v instanceof Date) return v.getHours()*60+v.getMinutes()+v.getSeconds()/60;
  const s=String(v).trim(), parts=s.split(":").map(x=>x.trim());
  if (parts.length<2||parts.length>3) return null;
  let h=0,m=0,sec=0;
  if (parts.length===2){m=Number(parts[0]);sec=Number(parts[1]);}
  else{h=Number(parts[0]);m=Number(parts[1]);sec=Number(parts[2]);}
  return [h,m,sec].some(n=>Number.isNaN(n))?null:h*60+m+sec/60;
}
function safeTrim_(v) { return (v===null||v===undefined)?"":String(v).trim(); }
