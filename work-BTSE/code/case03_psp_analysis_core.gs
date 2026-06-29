/**
 * Case 03 — Daily OPS Data Analysis System
 * Module: PSP Analysis Core — Scoring, Heatmaps, History, Full Automation Pipeline
 *
 * Functions in this file:
 *   [Analysis]
 *   - updateDailyAnalysis()         主入口（手動執行）
 *   - processBusinessData()         核心數據解析（代收/代付）
 *   - renderToSheetHorizontal()     24h 逐小時分頁報表
 *   - updateDashboardSummary()      Dashboard 摘要更新
 *   [Smart_Insights]
 *   - generateAdvancedAnalysis()    深度分析報告入口
 *   - renderScoringTable()          5 維度評分表（時效/成功/超時/金額/筆數）
 *   - renderHeatmap()               24h 筆數/金額熱點圖
 *   - renderPercentageHeatmap()     PSP 佔比熱點圖
 *   [History]
 *   - recordPspHistory()            每日評分存外部月度歷史表
 *   - calculateCombinedScores()     合併代收/代付計算總分
 *   - findOrInsertDateColumn()      確保日期欄存在（按日期排序）
 *   - findOrInsertPspRow()          確保 PSP 列存在（按幣別+名稱排序）
 *   [Automation Pipeline]
 *   - installTrigger()              安裝每 5 分鐘 Drive 掃描觸發器
 *   - pollDriveFolder()             掃描 Drive 資料夾，偵測代收/代付 CSV
 *   - runStep_A ~ runStep_D         4-step 自動化流程
 *   - _importCsvAsSheet()           CSV → Google Sheet 新分頁（含代付前處理）
 *   - _runUploadTransfer()          代收 raw data → 幣別分頁
 *   - _runDownloadTransfer()        代付 raw data → 幣別分頁
 *   - _runPSPMaster()               觸發 PSP 分析（自動化版本，無 UI.alert）
 */

// ============================================================
// CONFIG — 全域設定，填入你自己的 ID
// ============================================================
const CONFIG = {
  PARENT_FOLDER_ID:     "YOUR_PARENT_FOLDER_ID",  // 月份資料夾的上層資料夾 ID
  DASHBOARD_SHEET_NAME: "Dashboard",
  DATE_CELL:            "B1",
  TIMEOUT_SETTING_RANGE: "C2:E100",

  // 幣別分流：INR / PHP 依 PSP 名稱路由到不同分頁
  CURRENCY_MAPPING: {
    "INR": { "INR-喚醒": ["JHPAY","TPAY"], "DEFAULT": "INR-QR" },
    "PHP": { "PHP-Gcash": ["NEXPAYGCASH","SUNNYPAY-PHP-GCASH"], "DEFAULT": "PHP-Maya" }
  }
};

const PRO_CONFIG = {
  INSIGHT_SHEET_NAME: "Smart_Insights",
  WEIGHTS: { TIME:30, SUCCESS:30, OVERTIME:30, AMOUNT:5, COUNT:5 }
};

const HISTORY_CONFIG = {
  TARGET_SS_ID: "YOUR_HISTORY_SS_URL_OR_ID",  // PSP 月度歷史績效表（網址或 ID 皆可）
  WEIGHTS: { TIME:30, SUCCESS:30, OVERTIME:30, AMOUNT:5, COUNT:5 }
};

const CFG = {
  WATCH_FOLDER_ID:    "YOUR_WATCH_FOLDER_ID",   // 每天上傳 CSV 的 Drive 資料夾
  UPLOAD_SS_ID:       "YOUR_UPLOAD_SS_ID",       // 代收訂單分析試算表 ID
  DOWNLOAD_SS_ID:     "YOUR_DOWNLOAD_SS_ID",     // 代付訂單分析試算表 ID
  FILE_KEYWORDS:      { UPLOAD:"代收", DOWNLOAD:"代付" },
  PARENT_FOLDER_ID:   "YOUR_PARENT_FOLDER_ID",
  DASHBOARD_SHEET_NAME: "Dashboard",
  DATE_CELL:          "B1",
  TIMEOUT_SETTING_RANGE: "C2:E100",
  WAIT_TIMEOUT_MINUTES: 30,
  TG: { BOT_TOKEN: "YOUR_BOT_TOKEN", CHAT_ID: "YOUR_CHAT_ID" },
  TIMEZONE: "GMT+8"
};

// ============================================================
// ANALYSIS — 主入口與核心解析
// ============================================================

function updateDailyAnalysis() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const dashboard = ss.getSheetByName(CONFIG.DASHBOARD_SHEET_NAME);
  if (!dashboard) return;

  dashboard.getRange("B2").clearContent();
  const lastRowClear = dashboard.getLastRow();
  if (lastRowClear >= 3) dashboard.getRange(3,8,lastRowClear,9).setBackground(null).clearNote();

  const dateValue = dashboard.getRange(CONFIG.DATE_CELL).getValue();
  if (!(dateValue instanceof Date)) return;
  const targetDateStr = Utilities.formatDate(dateValue, "GMT+8", "yyyy-MM-dd");

  const timeoutData = dashboard.getRange(CONFIG.TIMEOUT_SETTING_RANGE).getValues();
  const timeoutMap  = {};
  timeoutData.forEach(row => {
    if (row && row[0]) timeoutMap[String(row[0]).trim()] = { collect:parseFloat(row[1])||5, payout:parseFloat(row[2])||5 };
  });

  const newCurrenciesFound = [];
  const collectAnalysis = processBusinessData(targetDateStr, CONFIG.PARENT_FOLDER_ID, "代收", "上分訂單分析", timeoutMap, newCurrenciesFound);
  const payoutAnalysis  = processBusinessData(targetDateStr, CONFIG.PARENT_FOLDER_ID, "代付", "回分訂單分析", timeoutMap, newCurrenciesFound);

  updateDashboardSummary(dashboard, collectAnalysis, payoutAnalysis);
  if (newCurrenciesFound.length>0) syncCurrencyToDashboard(dashboard, newCurrenciesFound);

  const ui = SpreadsheetApp.getUi();
  const response = ui.alert(`✅ ${targetDateStr} 分析更新完成！`, "是否要產出 Smart_Insights 深度分析報告？", ui.ButtonSet.YES_NO);
  if (response==ui.Button.YES && typeof generateAdvancedAnalysis==="function") {
    generateAdvancedAnalysis(targetDateStr, collectAnalysis, payoutAnalysis, timeoutMap);
    SpreadsheetApp.flush();
    ui.alert("🚀 Smart_Insights 深度報告已產出！");
  }

  const historyResponse = ui.alert("📊 歷史記錄同步", "是否要將本次評分存入 PSP 月度歷史績效表？", ui.ButtonSet.YES_NO);
  if (historyResponse==ui.Button.YES && typeof recordPspHistory==="function") {
    recordPspHistory(targetDateStr, collectAnalysis, payoutAnalysis, timeoutMap);
    ui.alert("✅ 歷史分數已成功同步！");
  }
}

function processBusinessData(dateStr, parentFolderId, bizType, fileNameKeyword, timeoutMap, newCurrenciesFound) {
  const [year, month] = dateStr.split("-");
  const subFolderName = `${year}-${month}`;
  const exactFileName = `${year}_${month} ${fileNameKeyword}`;

  const parentFolder = DriveApp.getFolderById(parentFolderId);
  const subFolders   = parentFolder.getFoldersByName(subFolderName);
  if (!subFolders.hasNext()) { console.warn(`找不到月份資料夾 [${subFolderName}]`); return null; }

  const files = subFolders.next().searchFiles(
    `title = '${exactFileName}' and mimeType = 'application/vnd.google-apps.spreadsheet' and trashed = false`
  );
  let sourceSS = null;
  if (files.hasNext()) sourceSS = SpreadsheetApp.openById(files.next().getId());
  if (!sourceSS) { console.warn(`找不到檔案 [${exactFileName}]`); return null; }

  const sheets      = sourceSS.getSheets();
  const targetSheet = sheets.find(s => s.getName().includes(dateStr) && s.getName().includes(bizType));
  if (!targetSheet) { console.warn(`找不到工作表（含 ${dateStr} 及 ${bizType}）`); return null; }

  const rawData = targetSheet.getDataRange().getDisplayValues();
  if (rawData.length < 2) return null;
  const data = rawData.slice(1);

  const idx = (bizType==="代付")
    ? { applyTime:4, type:5, psp:6, transId:7, currency:11, status:18, rawDuration:25, amount:12 }
    : { applyTime:3, type:4, psp:5, transId:7, currency:10, status:18, rawDuration:22, amount:12 };

  const analysis = {};
  data.forEach(row => {
    let originalCurrency = row[idx.currency].trim();
    const psp    = row[idx.psp].trim();
    const status = row[idx.status];
    const transId = row[idx.transId].trim();
    const type   = row[idx.type].trim();
    const amount = parseFloat(row[idx.amount].replace(/,/g,"")) || 0;

    if (bizType==="代收" && !type.includes("上分")) return;
    if (bizType==="代付" && !type.includes("回分")) return;
    if (!originalCurrency||!psp||psp.includes("银行转帐")) return;

    // 幣別分流
    let currency = originalCurrency;
    const upperPsp = psp.toUpperCase();
    if (CONFIG.CURRENCY_MAPPING[originalCurrency]) {
      const rules = CONFIG.CURRENCY_MAPPING[originalCurrency];
      let matched = false;
      for (let subType in rules) {
        if (subType==="DEFAULT") continue;
        if (rules[subType].map(v=>v.toUpperCase()).includes(upperPsp)) { currency=subType; matched=true; break; }
      }
      if (!matched && rules["DEFAULT"]) currency = rules["DEFAULT"];
    }

    if (!analysis[currency]) analysis[currency] = {};
    if (!analysis[currency][psp]) {
      analysis[currency][psp] = {};
      for (let i=0;i<24;i++) analysis[currency][psp][i] = {total:0,success:0,totalAmount:0,successAmount:0,totalSec:0,overtimeCount:0,overtimeIds:[]};
    }

    const hour = getHourByStringSplit_(row[idx.applyTime]);
    const stat  = analysis[currency][psp][hour];
    stat.total++;
    stat.totalAmount += amount;

    if (status.includes("完成")) {
      stat.success++;
      stat.successAmount += amount;
      const durationSec = parseDurationToSeconds_(row[idx.rawDuration]);
      stat.totalSec += durationSec;
      const limitMin = (bizType==="代收")
        ? (timeoutMap[currency]?.collect||timeoutMap[originalCurrency]?.collect||5)
        : (timeoutMap[currency]?.payout ||timeoutMap[originalCurrency]?.payout ||5);
      if (durationSec>(limitMin*60)) { stat.overtimeCount++; stat.overtimeIds.push(transId); }
    }
  });

  renderToSheetHorizontal_(dateStr, bizType, analysis, timeoutMap);
  return analysis;
}

function renderToSheetHorizontal_(dateStr, bizType, analysis, timeoutMap) {
  if (!analysis) return;
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheetName = `${dateStr}-${bizType}分析`;
  const fixedSheets = ["Smart_Insights","PSP Cost","評分Note","Dashboard"];

  let sheet = ss.getSheetByName(sheetName);
  if (!sheet) {
    const allSheets = ss.getSheets();
    let insertIndex = allSheets.length;
    for (let i=0;i<allSheets.length;i++) {
      const currentName = allSheets[i].getName();
      if (fixedSheets.includes(currentName)) continue;
      if (currentName>sheetName) { insertIndex=i; break; }
    }
    sheet = ss.insertSheet(sheetName, insertIndex);
  }
  sheet.clear().clearNotes();

  let currentRow = 1;
  Object.keys(analysis).sort().forEach(curr => {
    const psps = Object.keys(analysis[curr]).sort();
    const header1 = ["幣別: "+curr];
    const header2 = ["時段區間"];
    psps.forEach(psp => {
      header1.push(psp,"","","","","","","");
      header2.push("總筆數","成功數","總金額","成功金額","成功率(%)","處理時間","超時筆數","超時率(%)");
    });

    const colCount = header2.length;
    sheet.getRange(currentRow,1,1,colCount).setValues([header1]).setBackground("#333333").setFontColor("white").setFontWeight("bold");
    sheet.getRange(currentRow+1,1,1,colCount).setValues([header2]).setBackground("#f3f3f3").setFontWeight("bold");

    const dataBlock = [];
    for (let h=0;h<24;h++) {
      const rowData = [`${String(h).padStart(2,"0")}:00 - ${String(h).padStart(2,"0")}:59`];
      psps.forEach(psp => {
        const s = analysis[curr][psp][h];
        rowData.push(
          s.total, s.success, s.totalAmount, s.successAmount,
          s.total>0?s.success/s.total:0,
          formatTime_(s.success>0?s.totalSec/s.success:0),
          s.overtimeCount,
          s.success>0?s.overtimeCount/s.success:0
        );
      });
      dataBlock.push(rowData);
    }
    sheet.getRange(currentRow+2,1,24,colCount).setValues(dataBlock).setHorizontalAlignment("center");

    psps.forEach((psp,pIdx) => {
      const colStart = 2+(pIdx*8);
      sheet.getRange(currentRow+2,colStart,24,2).setNumberFormat("#,##0");
      sheet.getRange(currentRow+2,colStart+2,24,2).setNumberFormat("#,##0");
      sheet.getRange(currentRow+2,colStart+6,24,1).setNumberFormat("#,##0");
      sheet.getRange(currentRow+2,colStart+4,24,1).setNumberFormat("0.00%");
      sheet.getRange(currentRow+2,colStart+7,24,1).setNumberFormat("0.00%");
      sheet.getRange(currentRow+2,colStart+5,24,1).setNumberFormat("[hh]:mm:ss");
      for (let h=0;h<24;h++) {
        const s = analysis[curr][psp][h];
        if (s.overtimeCount>0) {
          sheet.getRange(currentRow+2+h,colStart,1,8).setBackground("#f4cccc");
          sheet.getRange(currentRow+2+h,colStart+6).setNote("超時單號:\n"+s.overtimeIds.join("\n"));
        }
      }
    });
    currentRow += 28;
  });
  sheet.autoResizeColumns(1, sheet.getLastColumn());
}

function updateDashboardSummary(dashboard, collectData, payoutData) {
  const lastCol = dashboard.getLastColumn();
  if (lastCol>=8) dashboard.getRange(1,8,dashboard.getMaxRows(),lastCol-7).clear().clearNote();

  const subHeader = [["類型","幣別","PSP名稱","總筆數","成功數","提單總金額","成功總金額","成功平均單筆","總超時數","平均時效"]];
  const summaryBody = [], summaryNotes = [];

  const processForSummary = (data, typeName) => {
    if (!data) return;
    Object.keys(data).sort().forEach(curr => {
      Object.keys(data[curr]).sort().forEach(psp => {
        let dayTotal=0,daySuccess=0,dayOvertime=0,dayTotalSec=0,dayAmount=0,daySuccessAmount=0,allOvertimeIds=[];
        for (let h=0;h<24;h++) {
          const s=data[curr][psp][h];
          dayTotal+=s.total; daySuccess+=s.success; dayOvertime+=s.overtimeCount;
          dayTotalSec+=s.totalSec; dayAmount+=s.totalAmount; daySuccessAmount+=s.successAmount;
          if (s.overtimeIds.length>0) allOvertimeIds.push(...s.overtimeIds);
        }
        if (dayTotal>0) {
          summaryBody.push([typeName,curr,psp,dayTotal,daySuccess,dayAmount,daySuccessAmount,
            daySuccess>0?daySuccessAmount/daySuccess:0,dayOvertime,
            formatTime_(daySuccess>0?dayTotalSec/daySuccess:0)]);
          summaryNotes.push(allOvertimeIds.length>0?"當日超時清單：\n"+allOvertimeIds.join("\n"):null);
        }
      });
    });
  };
  processForSummary(collectData,"代收");
  processForSummary(payoutData,"代付");

  if (summaryBody.length>0) {
    const colCount = 10;
    dashboard.getRange(1,8,1,colCount).setValues([["當日營運摘要 (Summary)","","","","","","","","",""]]).merge().setBackground("#444444").setFontColor("white").setFontWeight("bold").setHorizontalAlignment("center");
    dashboard.getRange(2,8,1,colCount).setValues(subHeader).setBackground("#f3f3f3").setFontWeight("bold").setHorizontalAlignment("center");
    const bodyRange = dashboard.getRange(3,8,summaryBody.length,colCount);
    bodyRange.setValues(summaryBody).setHorizontalAlignment("center");
    for (let i=0;i<summaryBody.length;i++) {
      if (summaryBody[i][8]>0) {
        dashboard.getRange(3+i,8,1,colCount).setBackground("#f4cccc");
        if (summaryNotes[i]) dashboard.getRange(3+i,17).setNote(summaryNotes[i]);
      }
    }
  }
}

// ============================================================
// SMART_INSIGHTS — 評分 + 熱點圖
// ============================================================

function generateAdvancedAnalysis(dateStr, collectAnalysis, payoutAnalysis, timeoutMap) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(PRO_CONFIG.INSIGHT_SHEET_NAME) || ss.insertSheet(PRO_CONFIG.INSIGHT_SHEET_NAME);
  const requiredCols = 32;
  if (sheet.getMaxColumns()<requiredCols) sheet.insertColumnsAfter(sheet.getMaxColumns(), requiredCols-sheet.getMaxColumns());
  sheet.clear();
  sheet.getRange(1,1,sheet.getMaxRows(),sheet.getMaxColumns()).clearNote().setBackground(null).setBorder(false,false,false,false,false,false);

  let currentLine = 1;
  sheet.getRange(currentLine,1).setValue(`🚀 營運洞察報告 (${dateStr})`).setFontSize(18).setFontWeight("bold").setFontColor("#1a73e8");
  currentLine += 2;

  currentLine = renderScoringTable_(sheet, currentLine, collectAnalysis, "代收", timeoutMap);
  currentLine += 2;
  currentLine = renderScoringTable_(sheet, currentLine, payoutAnalysis,  "代付", timeoutMap);
  currentLine += 2;
  currentLine = renderHeatmap_(sheet, currentLine, collectAnalysis, "代收 - 每小時【提單筆數】分佈", "筆數");
  currentLine += 2;
  currentLine = renderHeatmap_(sheet, currentLine, payoutAnalysis,  "代付 - 每小時【提單筆數】分佈", "筆數");
  currentLine += 2;
  currentLine = renderHeatmap_(sheet, currentLine, collectAnalysis, "代收 - 每小時【提單金額】分佈", "金額");
  currentLine += 2;
  currentLine = renderHeatmap_(sheet, currentLine, payoutAnalysis,  "代付 - 每小時【提單金額】分佈", "金額");
  currentLine += 2;
  currentLine = renderPercentageHeatmap_(sheet, currentLine, collectAnalysis, "代收 - 幣別內 PSP 提單佔比 (%)");
  currentLine += 2;
  currentLine = renderPercentageHeatmap_(sheet, currentLine, payoutAnalysis,  "代付 - 幣別內 PSP 提單佔比 (%)");

  sheet.setColumnWidth(1,80); sheet.setColumnWidth(2,150);
  for (let col=3;col<=27;col++) sheet.setColumnWidth(col,60);
  SpreadsheetApp.flush();
}

/**
 * 5 維度評分表
 * 注意：成功率=0 時，超時分強制歸 0（避免無單可超時卻得滿分的邏輯錯誤）
 */
function renderScoringTable_(sheet, startRow, data, bizType, timeoutMap) {
  if (!data) return startRow;
  sheet.getRange(startRow,1).setValue(`${bizType} - PSP 五維度效能評等報告`).setFontWeight("bold").setFontSize(12);
  const headers = [["幣別","PSP名稱","時效","成功率","超時率","總金額","總筆數","時效分","成功分","超時分","金額分","筆數分","總分","評等"]];
  sheet.getRange(startRow+1,1,1,14).setValues(headers).setBackground("#444444").setFontColor("white").setFontWeight("bold");

  const rows = [];
  const currTotals = {};
  Object.keys(data).forEach(curr => {
    currTotals[curr] = {amt:0,cnt:0};
    Object.values(data[curr]).forEach(pspData => {
      for(let h=0;h<24;h++){currTotals[curr].amt+=pspData[h].totalAmount;currTotals[curr].cnt+=pspData[h].total;}
    });
  });

  Object.keys(data).sort().forEach(curr => {
    const originalCurrency = curr.split("-")[0];
    const limitMin = bizType==="代收"
      ? (timeoutMap?.[curr]?.collect||timeoutMap?.[originalCurrency]?.collect||5)
      : (timeoutMap?.[curr]?.payout ||timeoutMap?.[originalCurrency]?.payout ||5);
    const limitSec = limitMin*60;

    Object.keys(data[curr]).sort().forEach(psp => {
      let total=0,success=0,overtime=0,totalSec=0,totalAmt=0;
      for (let h=0;h<24;h++){const s=data[curr][psp][h];total+=s.total;success+=s.success;overtime+=s.overtimeCount;totalSec+=s.totalSec;totalAmt+=s.totalAmount;}
      if (total>0) {
        const avgSec  = success>0?totalSec/success:limitSec;
        const sucRate = success/total;
        const overRate = success>0?overtime/success:0;

        const W = PRO_CONFIG.WEIGHTS;
        const s_time    = Math.max(0, W.TIME*(1-(avgSec/limitSec)));
        const s_success = W.SUCCESS*sucRate;
        const s_over    = sucRate===0?0:W.OVERTIME*(1-overRate);  // ← guard: 成功率=0 → 超時分=0
        const s_amt     = currTotals[curr].amt>0?(totalAmt/currTotals[curr].amt)*W.AMOUNT:0;
        const s_cnt     = currTotals[curr].cnt>0?(total/currTotals[curr].cnt)*W.COUNT:0;
        const scoreNum  = s_time+s_success+s_over+s_amt+s_cnt;

        let rating="D";
        if(scoreNum>=90)rating="A+"; else if(scoreNum>=85)rating="A"; else if(scoreNum>=75)rating="B+";
        else if(scoreNum>=65)rating="B"; else if(scoreNum>=55)rating="C+"; else if(scoreNum>=45)rating="C";
        else if(scoreNum>=35)rating="D+";

        rows.push([curr,psp,formatTime_(avgSec),sucRate,overRate,totalAmt,total,
          s_time.toFixed(1),s_success.toFixed(1),s_over.toFixed(1),s_amt.toFixed(1),s_cnt.toFixed(1),
          scoreNum.toFixed(1),rating]);
      }
    });
  });

  if (rows.length>0) {
    const range = sheet.getRange(startRow+2,1,rows.length,14);
    range.setValues(rows).setHorizontalAlignment("center").setBorder(true,true,true,true,true,true);
    sheet.getRange(startRow+2,3,rows.length,5).setBackground("#e1f5fe");
    sheet.getRange(startRow+2,4,rows.length,2).setNumberFormat("0.00%");
    sheet.getRange(startRow+2,6,rows.length,2).setNumberFormat("#,##0");
    sheet.getRange(startRow+2,8,rows.length,5).setBackground("#f3e5f5");
    sheet.getRange(startRow+2,13,rows.length,2).setBackground("#fff2cc").setFontWeight("bold");
    for(let i=0;i<rows.length-1;i++){
      if(rows[i][0]!==rows[i+1][0])
        sheet.getRange(startRow+2+i,1,1,14).setBorder(null,null,true,null,null,null,"black",SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
    }
  }
  return startRow+rows.length+2;
}

function renderHeatmap_(sheet, startRow, data, title, mode) {
  if (!data) return startRow;
  sheet.getRange(startRow,1).setValue(title).setFontWeight("bold").setFontSize(12);
  const header = ["幣別","PSP名稱","0h","1h","2h","3h","4h","5h","6h","7h","8h","9h","10h","11h","12h","13h","14h","15h","16h","17h","18h","19h","20h","21h","22h","23h",(mode==="金額"?"日總金額":"日總筆數")];
  sheet.getRange(startRow+1,1,1,27).setValues([header]).setBackground(mode==="金額"?"#d1e7dd":"#e8f0fe").setFontWeight("bold");

  const rows = [], currAverages = {};
  Object.keys(data).sort().forEach(curr => {
    let currTotal=0,cellCount=0;
    Object.keys(data[curr]).sort().forEach(psp => {
      const row = [curr,psp];
      let rowSumTotal=0,rowSumSuccess=0;
      for (let h=0;h<24;h++) {
        if (mode==="金額") {
          const val=data[curr][psp][h].totalAmount;
          row.push(val); rowSumTotal+=val;
          if(val>0){currTotal+=val;cellCount++;}
        } else {
          const totalCnt  = data[curr][psp][h].total;
          const successCnt = data[curr][psp][h].success;
          row.push(`${totalCnt} / ${successCnt}`);
          rowSumTotal+=totalCnt; rowSumSuccess+=successCnt;
          if(totalCnt>0){currTotal+=totalCnt;cellCount++;}
        }
      }
      row.push(mode==="金額"?rowSumTotal:`${rowSumTotal} / ${rowSumSuccess}`);
      rows.push(row);
    });
    currAverages[curr] = cellCount>0?currTotal/cellCount:0;
  });

  if (rows.length>0) {
    sheet.getRange(startRow+2,1,rows.length,27).setValues(rows).setHorizontalAlignment("center").setBorder(true,true,true,true,true,true);
    if (mode==="金額") {
      sheet.getRange(startRow+2,3,rows.length,25).setNumberFormat("#,##0");
      sheet.getRange(startRow+2,27,rows.length,1).setBackground("#f3f3f3").setFontWeight("bold").setNumberFormat("#,##0");
    } else {
      sheet.getRange(startRow+2,3,rows.length,25).setNumberFormat("@");
      sheet.getRange(startRow+2,27,rows.length,1).setBackground("#f3f3f3").setFontWeight("bold").setNumberFormat("@");
    }
    for (let r=0;r<rows.length;r++) {
      const avg=currAverages[rows[r][0]], curr=rows[r][0], psp=rows[r][1];
      for (let c=2;c<26;c++) {
        const compareVal = mode==="金額"?data[curr][psp][c-2].totalAmount:data[curr][psp][c-2].total;
        if (compareVal>avg) {
          const cell=sheet.getRange(startRow+2+r,c+1);
          if (mode==="筆數") cell.setBackground(compareVal>avg*1.5?"#ea9999":"#f4cccc");
          else cell.setBackground(compareVal>avg*1.5?"#b7e1cd":"#d1e7dd");
        }
      }
      if(r<rows.length-1&&rows[r][0]!==rows[r+1][0])
        sheet.getRange(startRow+2+r,1,1,27).setBorder(null,null,true,null,null,null,"black",SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
    }
  }
  return startRow+rows.length+2;
}

function renderPercentageHeatmap_(sheet, startRow, data, title) {
  if (!data) return startRow;
  sheet.getRange(startRow,1).setValue(title).setFontWeight("bold").setFontSize(12);
  const header = ["幣別","PSP名稱","0h","1h","2h","3h","4h","5h","6h","7h","8h","9h","10h","11h","12h","13h","14h","15h","16h","17h","18h","19h","20h","21h","22h","23h","全天佔比"];
  sheet.getRange(startRow+1,1,1,27).setValues([header]).setBackground("#fff2cc").setFontWeight("bold");

  const rows = [];
  Object.keys(data).sort().forEach(curr => {
    const psps = Object.keys(data[curr]).sort();
    const hourlyTotal = Array(24).fill(0);
    let currencyTotalAllDay = 0;
    psps.forEach(p=>{for(let h=0;h<24;h++){const val=data[curr][p][h].total;hourlyTotal[h]+=val;currencyTotalAllDay+=val;}});
    psps.forEach(psp=>{
      const row=[curr,psp];
      let pspTotalAllDay=0;
      for(let h=0;h<24;h++){const val=data[curr][psp][h].total;pspTotalAllDay+=val;row.push(hourlyTotal[h]>0?val/hourlyTotal[h]:0);}
      row.push(currencyTotalAllDay>0?pspTotalAllDay/currencyTotalAllDay:0);
      rows.push(row);
    });
  });

  if (rows.length>0) {
    sheet.getRange(startRow+2,1,rows.length,27).setValues(rows).setNumberFormat("0.0%").setHorizontalAlignment("center").setBorder(true,true,true,true,true,true);
    sheet.getRange(startRow+2,27,rows.length,1).setBackground("#f3f3f3").setFontWeight("bold");
    for(let r=0;r<rows.length;r++){
      for(let c=2;c<26;c++){
        const pVal=rows[r][c], cell=sheet.getRange(startRow+2+r,c+1);
        if(pVal>=1.0) cell.setBackground("#ffe599").setFontWeight("bold").setFontColor("#b45f06");
        else if(pVal>=0.8) cell.setBackground("#ffe599");
      }
      if(r<rows.length-1&&rows[r][0]!==rows[r+1][0])
        sheet.getRange(startRow+2+r,1,1,27).setBorder(null,null,true,null,null,null,"black",SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
    }
  }
  return startRow+rows.length+2;
}

// ============================================================
// HISTORY — 月度歷史績效
// ============================================================

function recordPspHistory(dateStr, collectData, payoutData, timeoutMap) {
  let ssId = HISTORY_CONFIG.TARGET_SS_ID;
  if (ssId.includes("docs.google.com")) {
    const matches = ssId.match(/\/d\/([a-zA-Z0-9-_]+)/);
    if (matches&&matches[1]) ssId=matches[1];
  }
  const targetSS = SpreadsheetApp.openById(ssId);
  const monthStr = dateStr.substring(0,7);
  let sheet = targetSS.getSheetByName(monthStr) || targetSS.insertSheet(monthStr);

  if (sheet.getLastColumn()<4) {
    sheet.getRange(1,1,1,4).setValues([["幣別","PSP名稱","平均分數","績效評級"]]).setBackground("#444444").setFontColor("white").setFontWeight("bold");
    sheet.setFrozenColumns(4);
    sheet.setFrozenRows(1);
  }

  const combinedScores = calculateCombinedScores_(collectData, payoutData, timeoutMap);
  let dateCol = findOrInsertDateColumn_(sheet, dateStr);

  combinedScores.forEach(item => {
    let rowIdx = findOrInsertPspRow_(sheet, item.curr, item.psp);
    const cell = sheet.getRange(rowIdx, dateCol);
    cell.setValue(item.score).setHorizontalAlignment("center");
    if (item.score>=85) cell.setBackground("#d9f9ad");
    else if (item.score<60) cell.setBackground("#f4cccc");
    else cell.setBackground(null);
  });
  SpreadsheetApp.flush();
}

function calculateCombinedScores_(collectData, payoutData, timeoutMap) {
  const scoreMap = {};
  const process = (data, bizType) => {
    if (!data) return;
    Object.keys(data).forEach(curr => {
      let totalAmt=0,totalCnt=0;
      Object.values(data[curr]).forEach(p=>{for(let h=0;h<24;h++){totalAmt+=p[h].totalAmount;totalCnt+=p[h].total;}});
      Object.keys(data[curr]).forEach(psp=>{
        let t=0,s=0,o=0,ts=0,ta=0;
        for(let h=0;h<24;h++){const d=data[curr][psp][h];t+=d.total;s+=d.success;o+=d.overtimeCount;ts+=d.totalSec;ta+=d.totalAmount;}
        if(t>0){
          const limitSec=(bizType==="代收"?(timeoutMap?.[curr]?.collect||5):(timeoutMap?.[curr]?.payout||5))*60;
          const avgSec=s>0?ts/s:limitSec;
          const W=HISTORY_CONFIG.WEIGHTS;
          const score=(Math.max(0,W.TIME*(1-(avgSec/limitSec))))+(W.SUCCESS*(s/t))+(W.OVERTIME*(1-(s>0?o/s:0)))+(totalAmt>0?(ta/totalAmt)*W.AMOUNT:0)+(totalCnt>0?(t/totalCnt)*W.COUNT:0);
          const key=`${curr}|${psp}`;
          if(!scoreMap[key]) scoreMap[key]={payin:[],payout:[]};
          scoreMap[key][bizType==="代收"?"payin":"payout"].push(score);
        }
      });
    });
  };
  process(collectData,"代收");
  process(payoutData,"代付");
  return Object.keys(scoreMap).map(key=>{
    const [curr,psp]=key.split("|");
    const vals=scoreMap[key];
    const finalScore=(vals.payin.length>0&&vals.payout.length>0)?(vals.payin[0]+vals.payout[0])/2:(vals.payin[0]||vals.payout[0]);
    return{curr,psp,score:parseFloat(finalScore.toFixed(1))};
  });
}

function findOrInsertDateColumn_(sheet, dateStr) {
  const lastCol = sheet.getLastColumn();
  if (lastCol<5) { sheet.getRange(1,5).setValue(dateStr).setBackground("#eeeeee").setFontWeight("bold"); return 5; }
  const headers = sheet.getRange(1,5,1,lastCol-4).getDisplayValues()[0];
  for (let i=0;i<headers.length;i++) {
    if (headers[i]===dateStr) return i+5;
    if (headers[i]>dateStr) { sheet.insertColumnBefore(i+5); sheet.getRange(1,i+5).setValue(dateStr).setBackground("#eeeeee").setFontWeight("bold"); return i+5; }
  }
  sheet.getRange(1,lastCol+1).setValue(dateStr).setBackground("#eeeeee").setFontWeight("bold");
  return lastCol+1;
}

function findOrInsertPspRow_(sheet, curr, psp) {
  const lastRow = sheet.getLastRow();
  const pspData = lastRow>=2?sheet.getRange(2,1,lastRow-1,2).getValues():[];
  for (let i=0;i<pspData.length;i++) if(pspData[i][0]===curr&&pspData[i][1]===psp) return i+2;
  sheet.appendRow([curr,psp]);
  if (lastRow>=2) { sheet.getRange(2,1,sheet.getLastRow()-1,sheet.getLastColumn()).sort([{column:1,ascending:true},{column:2,ascending:true}]); return findOrInsertPspRow_(sheet,curr,psp); }
  return sheet.getLastRow();
}

// ============================================================
// AUTOMATION PIPELINE
// ============================================================

function installTrigger() {
  ScriptApp.getProjectTriggers().filter(t=>t.getHandlerFunction()==="pollDriveFolder").forEach(t=>ScriptApp.deleteTrigger(t));
  ScriptApp.newTrigger("pollDriveFolder").timeBased().everyMinutes(5).create();
  Logger.log("✅ 觸發器安裝成功！每 5 分鐘自動掃描 Drive 資料夾。");
}

function pollDriveFolder() {
  const today = _todayStr_();
  const props = PropertiesService.getScriptProperties();
  if (props.getProperty(`DONE_${today}`)) return;

  const folder     = DriveApp.getFolderById(CFG.WATCH_FOLDER_ID);
  const todayStart = new Date(); todayStart.setHours(0,0,0,0);

  let uploadFileId   = props.getProperty(`CSV_UP_${today}`);
  let downloadFileId = props.getProperty(`CSV_DN_${today}`);

  const files = folder.getFiles();
  while (files.hasNext()) {
    const file = files.next();
    if (file.getLastUpdated()<todayStart) continue;
    if (!file.getName().toLowerCase().endsWith(".csv")) continue;
    const name=file.getName();
    if (name.includes(CFG.FILE_KEYWORDS.UPLOAD)&&!uploadFileId)   { uploadFileId=file.getId();   props.setProperty(`CSV_UP_${today}`,uploadFileId);   _tg_(`📥 偵測到代收 CSV：${name}`); }
    if (name.includes(CFG.FILE_KEYWORDS.DOWNLOAD)&&!downloadFileId){ downloadFileId=file.getId(); props.setProperty(`CSV_DN_${today}`,downloadFileId); _tg_(`📤 偵測到代付 CSV：${name}`); }
  }

  if (uploadFileId&&downloadFileId) {
    props.deleteProperty(`CSV_UP_${today}`); props.deleteProperty(`CSV_DN_${today}`);
    props.deleteProperty(`WAIT_NOTIFIED_${today}`); props.deleteProperty(`FIRST_SEEN_${today}`);
    const uploadFileName = DriveApp.getFileById(uploadFileId).getName();
    const csvDateStr = uploadFileName.substring(0,10);
    const dateStr = isNaN(new Date(csvDateStr).getTime())?today:csvDateStr;
    _tg_(`✅ 兩份 CSV 齊備，開始自動分析...`);
    runStep_A_ImportUpload_(dateStr, uploadFileId, downloadFileId);
    return;
  }

  if ((uploadFileId||downloadFileId)&&!props.getProperty(`WAIT_NOTIFIED_${today}`)) {
    if (!props.getProperty(`FIRST_SEEN_${today}`)) { props.setProperty(`FIRST_SEEN_${today}`,new Date().toISOString()); }
    else {
      const waited=(new Date()-new Date(props.getProperty(`FIRST_SEEN_${today}`)))/60000;
      if (waited>=CFG.WAIT_TIMEOUT_MINUTES) {
        const missing=[]; if(!uploadFileId)missing.push("代收 CSV"); if(!downloadFileId)missing.push("代付 CSV");
        _tg_(`⚠️ *報表缺檔提醒*\n📅 ${today}\n❌ 仍缺：${missing.join("、")}`);
        props.setProperty(`WAIT_NOTIFIED_${today}`,"true");
      }
    }
  }
}

function runStep_A_ImportUpload_(dateStr, uploadCsvId, downloadCsvId) {
  const props=PropertiesService.getScriptProperties();
  _tg_(`🚀 *${dateStr} OPS 日報開始*\n📊 Step 1/4：匯入代收 CSV...`);
  props.setProperty("PENDING_DATE",dateStr); props.setProperty(`DL_CSV_${dateStr}`,downloadCsvId); props.setProperty(`STATUS_${dateStr}`,"STEP_A");
  const uploadTabName=_importCsvAsSheet_(CFG.UPLOAD_SS_ID, uploadCsvId);
  props.setProperty(`UPLOAD_TAB_${dateStr}`,uploadTabName);
  _installOneShotTrigger_("runStep_B_ImportDownload_",1);
}
function runStep_B_ImportDownload_() {
  _deleteOneShotTrigger_("runStep_B_ImportDownload_");
  const props=PropertiesService.getScriptProperties(), dateStr=props.getProperty("PENDING_DATE");
  _tg_(`📊 Step 2/4：匯入代付 CSV...`);
  const downloadTabName=_importCsvAsSheet_(CFG.DOWNLOAD_SS_ID, props.getProperty(`DL_CSV_${dateStr}`));
  props.setProperty(`DOWNLOAD_TAB_${dateStr}`,downloadTabName);
  _installOneShotTrigger_("runStep_C_Transfer_",1);
}
function runStep_C_Transfer_() {
  _deleteOneShotTrigger_("runStep_C_Transfer_");
  const props=PropertiesService.getScriptProperties(), dateStr=props.getProperty("PENDING_DATE");
  _tg_(`📥 Step 3/4：執行上分 / 回分分析...`);
  const uploadSs=SpreadsheetApp.openById(CFG.UPLOAD_SS_ID), downloadSs=SpreadsheetApp.openById(CFG.DOWNLOAD_SS_ID);
  _runUploadTransfer_(uploadSs);
  _runDownloadTransfer_(downloadSs);
  _installOneShotTrigger_("runStep_D_PSP_",1);
}
function runStep_D_PSP_() {
  _deleteOneShotTrigger_("runStep_D_PSP_");
  const props=PropertiesService.getScriptProperties(), dateStr=props.getProperty("PENDING_DATE");
  _tg_(`📈 Step 4/4：PSP 綜合分析...`);
  _runPSPMaster_(dateStr);
  _sendDailySummary_(dateStr);
  props.setProperty(`DONE_${dateStr}`,"true"); props.setProperty(`STATUS_${dateStr}`,"SUCCESS");
  props.deleteProperty("PENDING_DATE"); props.deleteProperty(`DL_CSV_${dateStr}`);
  props.deleteProperty(`UPLOAD_TAB_${dateStr}`); props.deleteProperty(`DOWNLOAD_TAB_${dateStr}`);
  _tg_(`✅ *${dateStr} OPS 日報完成！* ⏱ ${_nowStr_()}`);
}

/**
 * CSV → Google Sheet 新分頁
 * 代付專用前處理：刪 S 欄（超时处理订单）+ 新增「處理時間」欄（HH:MM:SS）
 */
function _importCsvAsSheet_(spreadsheetId, csvFileId) {
  const ss=SpreadsheetApp.openById(spreadsheetId), csvFile=DriveApp.getFileById(csvFileId);
  const tabName=csvFile.getName().replace(/\.csv$/i,"");
  const isDownload=tabName.includes("代付");

  const existing=ss.getSheetByName(tabName); if(existing) ss.deleteSheet(existing);

  const raw=csvFile.getBlob().getDataAsString("UTF-8").replace(/^﻿/,"");
  const lines=raw.split(/\r?\n/).filter(l=>l.trim());
  let data=lines.map(line=>_parseCsvLine_(line));

  if (isDownload&&data.length>1) {
    const header=data[0];
    const sColIdx   =header.findIndex(h=>h.includes("超时处理订单"));
    const applyIdx  =header.findIndex(h=>h.includes("申请时间"));
    const doneIdx   =header.findIndex(h=>h.includes("完成时间"));

    if (sColIdx!==-1) data=data.map(row=>{const r=[...row];r.splice(sColIdx,1);return r;});

    data[0].push("處理時間");
    for (let i=1;i<data.length;i++) {
      const applyStr=(data[i][applyIdx]||"").toString().substring(0,19).replace(/\//g,"-").replace(" ","T");
      const doneStr =(data[i][doneIdx] ||"").toString().substring(0,19).replace(/\//g,"-").replace(" ","T");
      if (!applyStr||!doneStr){data[i].push("");continue;}
      const diff=Math.floor((new Date(doneStr)-new Date(applyStr))/1000);
      if (isNaN(diff)||diff<=0){data[i].push("");continue;}
      data[i].push(`${String(Math.floor(diff/3600)).padStart(2,"0")}:${String(Math.floor((diff%3600)/60)).padStart(2,"0")}:${String(diff%60).padStart(2,"0")}`);
    }
  }

  const maxCols=Math.max(...data.map(r=>r.length));
  const padded=data.map(r=>{while(r.length<maxCols)r.push("");return r;});
  const newSheet=ss.insertSheet(tabName);
  const BATCH=500;
  for(let start=0;start<padded.length;start+=BATCH){
    const chunk=padded.slice(start,start+BATCH);
    newSheet.getRange(start+1,1,chunk.length,maxCols).setValues(chunk);
    SpreadsheetApp.flush();
  }

  const analysisSheet=ss.getSheetByName("Analysis");
  if(analysisSheet) analysisSheet.getRange("A1").setValue(tabName);
  return tabName;
}

function _parseCsvLine_(line) {
  const result=[]; let cur="",inQuote=false;
  for(let i=0;i<line.length;i++){
    const ch=line[i];
    if(ch==='"'){if(inQuote&&line[i+1]==='"'){cur+='"';i++;}else inQuote=!inQuote;}
    else if(ch===","&&!inQuote){result.push(cur);cur="";}
    else cur+=ch;
  }
  result.push(cur);
  return result;
}

function _runUploadTransfer_(ss) {
  const COL={MERCHANT:2,APPLY:3,TYPE:4,PSP:5,CURRENCY:10,AMOUNT:12,DONE:17,STATUS:18};
  const CRYPTO=new Set(["USDT","USDC","BTC","ETH","BNB","TRX","DOGE"]);
  const CHANNEL_MAP={INR:{"INR-Indiapay":["JHPAY","TPAY"],DEFAULT:"INR-Worldpay"},PHP:{"PHP-Gcash":["NEXPAYGCASH","SUNNYPAY-PHP-GCASH"],DEFAULT:"PHP-Maya"}};

  const analysisSheet=ss.getSheetByName("Analysis"); if(!analysisSheet) return;
  const rawTabName=analysisSheet.getRange("A1").getValue().toString().trim();
  const rawSheet=ss.getSheetByName(rawTabName); if(!rawSheet) return;
  const dateStr=rawTabName.substring(0,10);
  const rawData=rawSheet.getDataRange().getValues();

  const agg={};
  for(let i=1;i<rawData.length;i++){
    const row=rawData[i];
    const type=(row[COL.TYPE]||"").toString().trim(),curr=(row[COL.CURRENCY]||"").toString().trim();
    const psp=(row[COL.PSP]||"").toString().trim().toUpperCase(),merchant=(row[COL.MERCHANT]||"").toString().trim();
    const status=(row[COL.STATUS]||"").toString().trim(),amount=parseFloat((row[COL.AMOUNT]||"0").toString().replace(/,/g,""))||0;
    if(type!=="上分"||!curr||CRYPTO.has(curr)||!merchant) continue;

    let sheetName=curr;
    if(CHANNEL_MAP[curr]){const mapping=CHANNEL_MAP[curr];sheetName=Object.entries(mapping).find(([k,psps])=>k!=="DEFAULT"&&psps.some(p=>psp.includes(p)))?.[0]||mapping.DEFAULT;}

    if(!agg[sheetName]) agg[sheetName]={};
    if(!agg[sheetName][merchant]) agg[sheetName][merchant]={total:0,success:0,totalSec:0};
    const a=agg[sheetName][merchant];
    a.total++;
    if(status==="完成"){a.success++;const diff=(new Date(row[COL.DONE])-new Date(row[COL.APPLY]))/1000;if(diff>0)a.totalSec+=diff;}
  }

  Object.entries(agg).forEach(([sheetName,merchants])=>{
    const sheet=ss.getSheetByName(sheetName); if(!sheet) return;
    const dateColIdx=_ensureDateCol_(sheet,dateStr,"upload");
    const sheetData=sheet.getDataRange().getValues();
    const updates=[];
    Object.entries(merchants).forEach(([merchant,a])=>{
      const rowIdx=sheetData.slice(2).findIndex(r=>(r[0]||"").toString().toLowerCase()===merchant.toLowerCase());
      if(rowIdx===-1) return;
      updates.push({row:rowIdx+3,data:[a.total,a.success,a.total>0?a.success/a.total:0]});
    });
    updates.forEach(u=>sheet.getRange(u.row,dateColIdx+1,1,u.data.length).setValues([u.data]));
    if(updates.length>0) SpreadsheetApp.flush();
  });
}

function _runDownloadTransfer_(ss) {
  const COL={MERCHANT:2,APPLY:4,TYPE:5,PSP:6,CURRENCY:11,AMOUNT:12,DONE:17,STATUS:18};
  const CRYPTO=new Set(["USDT","USDC","BTC","ETH","BNB","TRX","DOGE"]);

  const analysisSheet=ss.getSheetByName("Analysis"); if(!analysisSheet) return;
  const rawTabName=analysisSheet.getRange("A1").getValue().toString().trim();
  const rawSheet=ss.getSheetByName(rawTabName); if(!rawSheet) return;
  const dateStr=rawTabName.substring(0,10);
  const rawData=rawSheet.getDataRange().getValues();
  const lastColIdx=rawData[0].length-1;

  const agg={};
  for(let i=1;i<rawData.length;i++){
    const row=rawData[i];
    const type=(row[COL.TYPE]||"").toString().trim(),curr=(row[COL.CURRENCY]||"").toString().trim();
    const merchant=(row[COL.MERCHANT]||"").toString().trim(),status=(row[COL.STATUS]||"").toString().trim();
    const amount=parseFloat((row[COL.AMOUNT]||"0").toString().replace(/,/g,""))||0;
    if(type!=="回分"||!curr||CRYPTO.has(curr)||!merchant) continue;

    if(!agg[curr]) agg[curr]={};
    if(!agg[curr][merchant]) agg[curr][merchant]={total:0,success:0,totalAmt:0,successAmt:0,totalSec:0,over10:0};
    const a=agg[curr][merchant];
    a.total++; a.totalAmt+=amount;
    if(status==="完成"){
      a.success++; a.successAmt+=amount;
      const timeStr=(row[lastColIdx]||"").toString().trim();
      if(timeStr&&timeStr.includes(":")){
        const parts=timeStr.split(":");
        const sec=parseInt(parts[0]||0)*3600+parseInt(parts[1]||0)*60+parseInt(parts[2]||0);
        a.totalSec+=sec; if(sec>600) a.over10++;
      }
    }
  }

  Object.entries(agg).forEach(([currency,merchants])=>{
    const sheet=ss.getSheetByName(currency); if(!sheet) return;
    const dateColIdx=_ensureDateCol_(sheet,dateStr,"download");
    const sheetData=sheet.getDataRange().getDisplayValues();
    const updates=[];
    Object.entries(merchants).forEach(([merchant,a])=>{
      const rowIdx=sheetData.slice(2).findIndex(r=>(r[0]||"").toString().toLowerCase()===merchant.toLowerCase());
      if(rowIdx===-1) return;
      const avgSec=a.success>0?a.totalSec/a.success:0;
      updates.push({row:rowIdx+3,data:[a.totalAmt,a.successAmt,
        `${String(Math.floor(avgSec/3600)).padStart(2,"0")}:${String(Math.floor((avgSec%3600)/60)).padStart(2,"0")}:${String(Math.floor(avgSec%60)).padStart(2,"0")}`,
        a.over10,a.success>0?a.over10/a.success:0]});
    });
    updates.forEach(u=>sheet.getRange(u.row,dateColIdx+1,1,u.data.length).setValues([u.data]));
    if(updates.length>0) SpreadsheetApp.flush();
  });
}

function _ensureDateCol_(sheet, dateStr, type) {
  const COLS_PER_DAY=type==="upload"?3:5;
  const SUBTITLES=type==="upload"?["總筆數","成功筆數","成功率"]:["總金額","交易金額","均時","超10分鐘筆數","超10分鐘%"];
  const lastCol=sheet.getLastColumn();
  const headerVals=lastCol>0?sheet.getRange(1,1,1,lastCol).getValues()[0]:[];
  const headerDisp=lastCol>0?sheet.getRange(1,1,1,lastCol).getDisplayValues()[0]:[];

  let dateColIdx=headerVals.findIndex((h,i)=>{
    if(i%COLS_PER_DAY!==0) return false;
    if(h instanceof Date) return Utilities.formatDate(h,Session.getScriptTimeZone(),"yyyy-MM-dd")===dateStr;
    return false;
  });
  if(dateColIdx===-1) dateColIdx=headerDisp.findIndex((h,i)=>i%COLS_PER_DAY===0&&h.trim().substring(0,10)===dateStr);
  if(dateColIdx!==-1) return dateColIdx;

  const newCol1based=lastCol+1;
  sheet.getRange(1,newCol1based).setValue(new Date(dateStr)).setNumberFormat("yyyy-MM-dd");
  sheet.getRange(2,newCol1based,1,SUBTITLES.length).setValues([SUBTITLES]);
  return newCol1based-1;
}

function _runPSPMaster_(dateStr) {
  const ss=SpreadsheetApp.getActiveSpreadsheet();
  const dashboard=ss.getSheetByName(CFG.DASHBOARD_SHEET_NAME); if(!dashboard) return;
  dashboard.getRange(CFG.DATE_CELL).setValue(new Date(dateStr));
  dashboard.getRange("B2").clearContent();
  const lr=dashboard.getLastRow();
  if(lr>=3) dashboard.getRange(3,8,lr,9).setBackground(null).clearNote();

  const timeoutMap={};
  dashboard.getRange(CFG.TIMEOUT_SETTING_RANGE).getValues().forEach(row=>{
    if(row[0]) timeoutMap[String(row[0]).trim()]={collect:parseFloat(row[1])||5,payout:parseFloat(row[2])||5};
  });

  const newCurrs=[];
  const collectAnalysis=processBusinessData(dateStr,CFG.PARENT_FOLDER_ID,"代收","上分訂單分析",timeoutMap,newCurrs);
  const payoutAnalysis =processBusinessData(dateStr,CFG.PARENT_FOLDER_ID,"代付","回分訂單分析",timeoutMap,newCurrs);
  updateDashboardSummary(dashboard,collectAnalysis,payoutAnalysis);
  if(newCurrs.length>0) syncCurrencyToDashboard(dashboard,newCurrs);
  if(typeof generateAdvancedAnalysis==="function") generateAdvancedAnalysis(dateStr,collectAnalysis,payoutAnalysis,timeoutMap);
  if(typeof recordPspHistory==="function") recordPspHistory(dateStr,collectAnalysis,payoutAnalysis,timeoutMap);
  SpreadsheetApp.flush();
}

function _sendDailySummary_(dateStr) {
  const dashboard=SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CFG.DASHBOARD_SHEET_NAME);
  const lr=dashboard.getLastRow(); if(lr<3){_tg_(`📊 ${dateStr} 無資料`);return;}
  const data=dashboard.getRange(3,8,lr-2,10).getValues();
  const cLines=[],pLines=[];
  data.forEach(row=>{
    if(!row[0]) return;
    const [type,curr,psp,total,success,,,,overtime]=row;
    const rate=total>0?((success/total)*100).toFixed(1):"0.0";
    const line=`  ${curr}/${psp}：${success}/${total}筆(${rate}%) 超時:${overtime}`;
    if(type==="代收") cLines.push(line); else pLines.push(line);
  });
  const lines=[`📊 *${dateStr} OPS 日報摘要*`];
  if(cLines.length){lines.push("\n*代收*");lines.push(...cLines.slice(0,15));}
  if(pLines.length){lines.push("\n*代付*");lines.push(...pLines.slice(0,15));}
  _tg_(lines.join("\n"));
}

// ============================================================
// Shared Helpers
// ============================================================
function getHourByStringSplit_(value) { if(!value)return 0; const match=value.toString().match(/(\d{1,2}):\d{2}:\d{2}/); return match?parseInt(match[1]):0; }
function parseDurationToSeconds_(val) { if(!val)return 0; let s=val.toString().trim().replace(/,/g,""); if(s.startsWith("-"))return 0; if(s.includes("."))s=s.split(".")[0]; const parts=s.split(":"); if(parts.length===3)return parseInt(parts[0])*3600+parseInt(parts[1])*60+parseInt(parts[2]); if(parts.length===2)return parseInt(parts[0])*60+parseInt(parts[1]); return isNaN(parseInt(s))?0:parseInt(s); }
function formatTime_(totalSeconds) { if(totalSeconds<0)totalSeconds=0; return `${String(Math.floor(totalSeconds/3600)).padStart(2,"0")}:${String(Math.floor((totalSeconds%3600)/60)).padStart(2,"0")}:${String(Math.floor(totalSeconds%60)).padStart(2,"0")}`; }
function syncCurrencyToDashboard(dashboard, newCurrs) { const cValues=dashboard.getRange("C:C").getValues(); let cLastRow=0; for(let i=cValues.length-1;i>=0;i--){if(cValues[i]&&cValues[i][0]!==""){cLastRow=i+1;break;}} if(cLastRow<1)cLastRow=1; dashboard.getRange(cLastRow+1,3,newCurrs.length,3).setValues(newCurrs.map(c=>[c,5,5])).setBackground("#fff2cc").setBorder(true,true,true,true,true,true); }
function _todayStr_() { return Utilities.formatDate(new Date(),CFG.TIMEZONE,"yyyy-MM-dd"); }
function _nowStr_()   { return Utilities.formatDate(new Date(),CFG.TIMEZONE,"HH:mm:ss"); }
function _tg_(message) { if(!CFG.TG.BOT_TOKEN){Logger.log(`[TG] ${message}`);return;} try{UrlFetchApp.fetch(`https://api.telegram.org/bot${CFG.TG.BOT_TOKEN}/sendMessage`,{method:"post",contentType:"application/json",payload:JSON.stringify({chat_id:CFG.TG.CHAT_ID,text:message,parse_mode:"Markdown"})});}catch(e){Logger.log(`TG 發送失敗: ${e}`);} }
function _installOneShotTrigger_(funcName,minutes) { ScriptApp.newTrigger(funcName).timeBased().after(minutes*60*1000).create(); }
function _deleteOneShotTrigger_(funcName) { ScriptApp.getProjectTriggers().filter(t=>t.getHandlerFunction()===funcName).forEach(t=>ScriptApp.deleteTrigger(t)); }
