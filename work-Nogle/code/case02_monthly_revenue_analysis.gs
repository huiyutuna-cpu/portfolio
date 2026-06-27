// ============================================================
// Case 02 — Monthly Merchant Volume & Revenue Analysis System
// Tools: Google Apps Script · Google Sheets · Telegram Bot API
// ============================================================

// CONFIG — replace with actual values in your own deployment
const TELEGRAM_BOT_TOKEN = 'CONFIG_TELEGRAM_BOT_TOKEN';
const TELEGRAM_API = 'https://api.telegram.org/bot' + TELEGRAM_BOT_TOKEN + '/';
const ACCOUNT_NOTIFY_CHATS = ['CONFIG_TELEGRAM_CHAT_ID'];

// ============================================================
// SYSTEM 1 — Merchant Volume Diff Report
// Compares two monthly Sheets; outputs missing / new / stable merchants
// ============================================================

function compareMerchantsByCurrencyFourSheets() {
  const mainSS = SpreadsheetApp.getActiveSpreadsheet();
  const outSheet = mainSS.getActiveSheet();

  const lastRow = outSheet.getLastRow();
  if (lastRow > 3) {
    outSheet.getRange(6, 2, lastRow - 3, 14).clearContent();
  }

  const urls = [
    { old: outSheet.getRange("B2").getValue(), new: outSheet.getRange("B3").getValue() },
    { old: outSheet.getRange("C2").getValue(), new: outSheet.getRange("C3").getValue() }
  ];

  const spreadsheets = urls.map(u => {
    const urlOld = String(u.old || "").trim();
    const urlNew = String(u.new || "").trim();
    if (!urlOld || !urlNew) throw new Error("請在 B2/B3 與 C2/C3 填入試算表連結或 ID。");
    const idOld = extractSpreadsheetId(urlOld);
    const idNew = extractSpreadsheetId(urlNew);
    if (!idOld || !idNew) throw new Error("無法解析試算表 ID，請檢查 B2/B3 與 C2/C3。");
    return { old: SpreadsheetApp.openById(idOld), new: SpreadsheetApp.openById(idNew) };
  });

  function compareTwoSheets(ssOld, ssNew, startRow = 3) {
    const result = { leftSide: [], rightSide: [], bothMonths: [] };
    const excludeSheets = ["商戶PSP統計-代收", "商戶PSP統計-代付", "PSP 總覽報表", "Merchant 總覽報表",
      "當天所有幣別狀況", "工作表17", "跑量公式", "INR-Bestpay", "Bestpay跑量"];

    const currenciesOld = ssOld.getSheets().map(s => s.getName()).filter(n => !excludeSheets.includes(n));
    const currenciesNew = ssNew.getSheets().map(s => s.getName()).filter(n => !excludeSheets.includes(n));
    const commonCurrencies = currenciesOld.filter(c => currenciesNew.includes(c));
    const onlyOld = currenciesOld.filter(c => !currenciesNew.includes(c));
    const onlyNew = currenciesNew.filter(c => !currenciesOld.includes(c));

    function readMerchants(sht) {
      const last = sht.getLastRow();
      const numRows = last >= startRow ? (last - startRow + 1) : 0;
      if (numRows <= 0) return [];
      return Array.from(new Set(
        sht.getRange(startRow, 1, numRows, 1).getValues()
          .flat()
          .map(v => String(v || "").trim())
          .filter(v => v && !v.toLowerCase().includes("test"))
      )).sort((a, b) => a.localeCompare(b, 'zh-Hant'));
    }

    commonCurrencies.sort((a, b) => a.localeCompare(b, 'zh-Hant')).forEach(cur => {
      const shtOld = ssOld.getSheetByName(cur);
      const shtNew = ssNew.getSheetByName(cur);
      if (!shtOld || !shtNew) return;
      const oldMerchants = readMerchants(shtOld);
      const newMerchants = readMerchants(shtNew);
      const setOld = new Set(oldMerchants);
      const setNew = new Set(newMerchants);

      const missing = oldMerchants.filter(m => !setNew.has(m)).sort((a, b) => a.localeCompare(b, 'zh-Hant'));
      if (missing.length > 0) {
        result.leftSide.push([`*** 幣別: ${cur} ***`, ""]);
        missing.forEach(m => result.leftSide.push([m, cur]));
        result.leftSide.push(["", ""]);
      }

      const added = newMerchants.filter(m => !setOld.has(m)).sort((a, b) => a.localeCompare(b, 'zh-Hant'));
      if (added.length > 0) {
        result.rightSide.push([`*** 幣別: ${cur} ***`, ""]);
        added.forEach(m => result.rightSide.push([m, cur]));
        result.rightSide.push(["", ""]);
      }

      const both = oldMerchants.filter(m => setNew.has(m)).sort((a, b) => a.localeCompare(b, 'zh-Hant'));
      if (both.length > 0) {
        result.bothMonths.push([`*** 幣別: ${cur} ***`, ""]);
        both.forEach(m => result.bothMonths.push([m, cur]));
        result.bothMonths.push(["", ""]);
      }
    });

    onlyOld.sort((a, b) => a.localeCompare(b, 'zh-Hant')).forEach(cur => {
      const shtOld = ssOld.getSheetByName(cur);
      if (!shtOld) return;
      const oldMerchants = readMerchants(shtOld);
      result.leftSide.push([`*** 幣別: ${cur} ***`, ""]);
      if (oldMerchants.length === 0) result.leftSide.push(["（無商戶資料）", cur]);
      else oldMerchants.forEach(m => result.leftSide.push([m, cur]));
      result.leftSide.push(["", ""]);
    });

    onlyNew.sort((a, b) => a.localeCompare(b, 'zh-Hant')).forEach(cur => {
      const shtNew = ssNew.getSheetByName(cur);
      if (!shtNew) return;
      const newMerchants = readMerchants(shtNew);
      result.rightSide.push([`*** 幣別: ${cur} ***`, ""]);
      if (newMerchants.length === 0) result.rightSide.push(["（無商戶資料）", cur]);
      else newMerchants.forEach(m => result.rightSide.push([m, cur]));
      result.rightSide.push(["", ""]);
    });

    return result;
  }

  const allResults = spreadsheets.map(s => compareTwoSheets(s.old, s.new));

  const finalLeft = [["🟠 本月缺失商戶", ""]];
  const finalRight = [["🟢 本月新增商戶", ""]];
  const finalBoth = [["🔵 兩個月都有的商戶", ""]];

  allResults.forEach(r => {
    finalLeft.push(...r.leftSide.slice(1));
    finalRight.push(...r.rightSide.slice(1));
    finalBoth.push(...r.bothMonths.slice(1));
  });

  outSheet.getRange(6, 2, finalLeft.length, 2).setValues(finalLeft.length > 0 ? finalLeft : [["沒有🟠 本月缺失商戶差異", ""]]);
  outSheet.getRange(6, 5, finalRight.length, 2).setValues(finalRight.length > 0 ? finalRight : [["沒有🟢 本月新增商戶差異", ""]]);
  outSheet.getRange(6, 13, finalBoth.length, 2).setValues(finalBoth.length > 0 ? finalBoth : [["沒有🔵 兩個月都有商戶", ""]]);
}

function extractSpreadsheetId(str) {
  if (!str) return null;
  const maybeId = str.match(/[-\w]{25,}/);
  return maybeId ? maybeId[0] : null;
}

// ============================================================
// SYSTEM 1b — Archive diff result to history sheet
// ============================================================

function appendCompareResultTo2025Sheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sourceSheet = ss.getActiveSheet();
  let targetSheet = ss.getSheetByName("2025");
  if (!targetSheet) targetSheet = ss.insertSheet("2025");

  const leftStartRow = 4;
  const leftStartCol = 2;
  const lastRow = sourceSheet.getLastRow();

  let leftLastRow = leftStartRow - 1;
  for (let r = lastRow; r >= leftStartRow; r--) {
    if (sourceSheet.getRange(r, leftStartCol).getValue() || sourceSheet.getRange(r, leftStartCol + 1).getValue()) {
      leftLastRow = r; break;
    }
  }
  const leftNumRows = leftLastRow >= leftStartRow ? (leftLastRow - leftStartRow + 1) : 0;
  const leftData = leftNumRows > 0 ? sourceSheet.getRange(leftStartRow, leftStartCol, leftNumRows, 2).getValues() : [];

  const rightStartCol = 5;
  let rightLastRow = leftStartRow - 1;
  for (let r = lastRow; r >= leftStartRow; r--) {
    if (sourceSheet.getRange(r, rightStartCol).getValue() || sourceSheet.getRange(r, rightStartCol + 1).getValue()) {
      rightLastRow = r; break;
    }
  }
  const rightNumRows = rightLastRow >= leftStartRow ? (rightLastRow - leftStartRow + 1) : 0;
  const rightData = rightNumRows > 0 ? sourceSheet.getRange(leftStartRow, rightStartCol, rightNumRows, 2).getValues() : [];

  const leftFiltered = leftData.filter(row => {
    const m = row[0].toString().trim();
    return m && !m.startsWith("***") && m !== "🟠 本月缺失商戶" && m !== "";
  });
  const rightFiltered = rightData.filter(row => {
    const m = row[0].toString().trim();
    return m && !m.startsWith("***") && m !== "🟢 本月新增商戶" && m !== "" && m !== "🟢 本月新增商戶（整個幣別新增）";
  });

  const maxLen = Math.max(leftFiltered.length, rightFiltered.length);
  const combined = [];
  for (let i = 0; i < maxLen; i++) {
    const leftRow = leftFiltered[i] || ["", ""];
    const rightRow = rightFiltered[i] || ["", ""];
    combined.push([leftRow[0], leftRow[1], rightRow[0], rightRow[1]]);
  }

  let targetLastRow = targetSheet.getLastRow();
  let appendStartRow = targetLastRow >= 1 ? targetLastRow + 3 : 1;
  targetSheet.getRange(appendStartRow, 1, 1, 4).setValues([["🟠 本月缺失商戶", "", "🟢 本月新增商戶", ""]]);
  if (combined.length > 0) {
    targetSheet.getRange(appendStartRow + 1, 1, combined.length, 4).setValues(combined);
  }
}

// ============================================================
// SYSTEM 1c — Send merchant diff to Telegram
// ============================================================

function sendMerchantDiffReportToTelegram() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sourceSheet = ss.getSheetByName("商戶分析主表");
  if (!sourceSheet) { SpreadsheetApp.getUi().alert("找不到名為 '商戶分析主表' 的工作表"); return; }

  const monthStartRaw = sourceSheet.getRange("D2").getValue();
  const monthEndRaw = sourceSheet.getRange("D3").getValue();
  const formatMonthForDisplay = d => {
    if (d instanceof Date) return Utilities.formatDate(d, Session.getScriptTimeZone(), "M") + "月";
    return String(d).replace(/\D/g, '') + "月";
  };
  const monthStart = formatMonthForDisplay(monthStartRaw);
  const monthEnd = formatMonthForDisplay(monthEndRaw);

  const missingData = sourceSheet.getRange("B4:C1000").getValues().filter(r => r[0] && r[1]);
  const newData = sourceSheet.getRange("E4:F1000").getValues().filter(r => r[0] && r[1]);

  let msg = `📊【HTPAY｜商戶異動報告】\n\n`;
  msg += `📆 報告月份：${monthStart} 和 ${monthEnd} 差異\n\n`;
  msg += `🟠 本月缺失商戶\n`;
  msg += missingData.length === 0 ? `（無本月缺失商戶）\n` : missingData.map(([merchant, currency]) => `${merchant} ${currency}`).join("\n");
  msg += `\n\n`;
  msg += `🟢 本月新增商戶\n`;
  msg += newData.length === 0 ? `（無本月新增商戶）` : newData.map(([merchant, currency]) => `${merchant} ${currency}`).join("\n");

  sourceSheet.getRange("I7").setValue(msg).setWrap(true);
  sendToTG(msg);
}

function sendToTG(message) {
  if (!message || message.trim() === "") return;
  const safeMsg = message.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;");
  ACCOUNT_NOTIFY_CHATS.forEach(chatId => {
    try {
      UrlFetchApp.fetch(TELEGRAM_API + 'sendMessage', {
        method: 'post',
        payload: { chat_id: chatId, text: safeMsg, parse_mode: 'HTML' },
        muteHttpExceptions: true
      });
    } catch (e) { Logger.log("TG傳送失敗: " + e.message); }
  });
}

// ============================================================
// SYSTEM 1d — Merchant volume diff with color coding + TG
// ============================================================

function compareMerchantVolumeDiffWithColorAndStatus() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const mainSheet = ss.getSheetByName("商戶分析主表");

  const sheetPairs = [
    { old: mainSheet.getRange("B2").getValue(), new: mainSheet.getRange("B3").getValue(), label: "B2→B3" },
    { old: mainSheet.getRange("C2").getValue(), new: mainSheet.getRange("C3").getValue(), label: "C2→C3" }
  ];

  const monthOld = mainSheet.getRange("D2").getValue();
  const monthNew = mainSheet.getRange("D3").getValue();
  const outSheetName = `${monthOld}月 vs ${monthNew}月 商戶月度交易分析`;

  const existingSheet = ss.getSheetByName(outSheetName);
  if (existingSheet) ss.deleteSheet(existingSheet);
  const outSheet = ss.insertSheet(outSheetName);
  outSheet.clear();

  function getSpreadsheet(idOrUrl) {
    const id = (idOrUrl.match(/[-\w]{25,}/) || [])[0];
    if (!id) throw new Error("無法解析試算表ID: " + idOrUrl);
    return SpreadsheetApp.openById(id);
  }

  function parseNumber(val) {
    const n = Number(String(val || "").replace(/[,％%]/g, "").trim());
    return isNaN(n) ? 0 : n;
  }

  function getGrowthColor(valStr) {
    if (!valStr) return null;
    if (valStr.includes("▲")) return "#006400";
    if (valStr.includes("▼")) return "#8B0000";
    const num = parseFloat(valStr.replace(/[^\d.-]/g, ''));
    if (isNaN(num)) return null;
    if (num > 0) return "green";
    if (num < 0) return "red";
    return null;
  }

  function sumSheetData(spreadsheet) {
    const result = {};
    const typesToCheck = ["上分", "回分"];
    const englishNameRegex = /^[A-Za-z0-9 _-]+$/;

    spreadsheet.getSheets().forEach(sheet => {
      const sheetName = sheet.getName();
      if (!englishNameRegex.test(sheetName)) return;
      const currency = (sheet.getRange("A1").getValue() || sheetName).toString().trim();
      const data = sheet.getDataRange().getValues();
      if (data.length < 2) return;

      const dateRow = data[0];
      const validCols = [];
      for (let col = 3; col < dateRow.length; col += 7) {
        if (dateRow[col] !== "") validCols.push(col);
      }

      for (let r = 1; r < data.length; r++) {
        const row = data[r];
        const merchant = (row[0] || "").toString().trim();
        const type = (row[2] || "").toString().trim();
        if (!merchant || !typesToCheck.includes(type) || /test/i.test(merchant)) continue;
        if (!result[merchant]) result[merchant] = {};
        if (!result[merchant][currency]) result[merchant][currency] = {};
        if (!result[merchant][currency][type]) result[merchant][currency][type] = { total: 0, success: 0, fail: 0, amount: 0 };
        let subtotal = result[merchant][currency][type];
        validCols.forEach(col => {
          subtotal.total += parseNumber(row[col]);
          subtotal.success += parseNumber(row[col + 1]);
          subtotal.fail += parseNumber(row[col + 2]);
          subtotal.amount += parseNumber(row[col + 5]);
        });
      }
    });

    for (const m in result) {
      for (const c in result[m]) {
        for (const t in result[m][c]) {
          const obj = result[m][c][t];
          obj.successRate = obj.total > 0 ? obj.success / obj.total : 0;
        }
      }
    }
    return result;
  }

  function calcDiff(oldData, newData) {
    const diffRows = [];
    const allMerchants = new Set([...Object.keys(oldData), ...Object.keys(newData)]);
    allMerchants.forEach(merchant => {
      const currencies = new Set([...Object.keys(oldData[merchant] || {}), ...Object.keys(newData[merchant] || {})]);
      currencies.forEach(currency => {
        const types = new Set([
          ...Object.keys((oldData[merchant] || {})[currency] || {}),
          ...Object.keys((newData[merchant] || {})[currency] || {})
        ]);
        types.forEach(type => {
          const oldVal = ((oldData[merchant] || {})[currency] || {})[type] || { total: 0, success: 0, fail: 0, successRate: 0, amount: 0 };
          const newVal = ((newData[merchant] || {})[currency] || {})[type] || { total: 0, success: 0, fail: 0, successRate: 0, amount: 0 };

          function growth(oldNum, newNum) {
            if (oldNum === 0 && newNum > 0) return "▲ 新商戶";
            if (oldNum > 0 && newNum === 0) return "▼ 本月未跑";
            if (oldNum === 0 && newNum === 0) return "-";
            let val = (((newNum - oldNum) / oldNum) * 100).toFixed(2) + "%";
            if (newNum - oldNum > 0) val += " ▲"; else if (newNum - oldNum < 0) val += " ▼";
            return val;
          }

          let status = "";
          if (oldVal.total === 0 && newVal.total > 0) status = "▲ 新商戶";
          else if (oldVal.total > 0 && newVal.total === 0) status = "▼ 本月未跑";

          diffRows.push({
            merchant, currency, type,
            totalOld: oldVal.total, totalNew: newVal.total, totalDiff: newVal.total - oldVal.total,
            successOld: oldVal.success, successNew: newVal.success, successDiff: newVal.success - oldVal.success,
            failOld: oldVal.fail, failNew: newVal.fail, failDiff: newVal.fail - oldVal.fail,
            successRateOld: oldVal.successRate, successRateNew: newVal.successRate,
            amountOld: oldVal.amount, amountNew: newVal.amount, amountDiff: newVal.amount - oldVal.amount,
            totalGrowth: growth(oldVal.total, newVal.total),
            successGrowth: growth(oldVal.success, newVal.success),
            failGrowth: growth(oldVal.fail, newVal.fail),
            successRateGrowth: growth(oldVal.successRate, newVal.successRate),
            amountGrowth: growth(oldVal.amount, newVal.amount),
            status
          });
        });
      });
    });
    return diffRows;
  }

  let outputRow = 1;
  const header = [
    "商戶", "幣別", "類型",
    "總筆數(舊)", "總筆數(新)", "總筆數差異",
    "成功筆數(舊)", "成功筆數(新)", "成功筆數差異",
    "失敗筆數(舊)", "失敗筆數(新)", "失敗筆數差異",
    "成功率(舊)", "成功率(新)",
    "成功金額(舊)", "成功金額(新)", "成功金額差異",
    "總筆數變化%", "成功筆數變化%", "失敗筆數變化%", "成功率變化%", "成功金額變化%",
    "狀態"
  ];
  outSheet.getRange(outputRow, 2, 1, header.length).setValues([header]);
  outputRow++;

  const over150List = [];
  const below60List = [];

  sheetPairs.forEach(pair => {
    const oldSS = getSpreadsheet(pair.old);
    const newSS = getSpreadsheet(pair.new);
    const oldData = sumSheetData(oldSS);
    const newData = sumSheetData(newSS);
    const diffs = calcDiff(oldData, newData);

    outSheet.getRange(outputRow, 2).setValue("比較來源: " + pair.label);
    outputRow++;

    diffs.forEach(row => {
      const values = [
        row.merchant, row.currency, row.type,
        row.totalOld, row.totalNew, row.totalDiff,
        row.successOld, row.successNew, row.successDiff,
        row.failOld, row.failNew, row.failDiff,
        (row.successRateOld * 100).toFixed(2) + "%", (row.successRateNew * 100).toFixed(2) + "%",
        row.amountOld, row.amountNew, row.amountDiff,
        row.totalGrowth, row.successGrowth, row.failGrowth, row.successRateGrowth, row.amountGrowth,
        row.status
      ];
      outSheet.getRange(outputRow, 2, 1, values.length).setValues([values]);

      [[6, row.totalDiff], [9, row.successDiff], [12, row.failDiff], [17, row.amountDiff]].forEach(([col, num]) => {
        const cell = outSheet.getRange(outputRow, col + 1);
        if (num > 0) cell.setFontColor("green");
        else if (num < 0) cell.setFontColor("red");
      });

      [18, 19, 20, 21, 22].forEach((col, index) => {
        const color = getGrowthColor([row.totalGrowth, row.successGrowth, row.failGrowth, row.successRateGrowth, row.amountGrowth][index]);
        if (color) outSheet.getRange(outputRow, col + 1).setFontColor(color);
      });

      const statusCell = outSheet.getRange(outputRow, 23);
      if (row.status === "▲ 新商戶") statusCell.setFontColor("#006400");
      else if (row.status === "▼ 本月未跑") statusCell.setFontColor("#8B0000");

      const amountGrowthVal = parseFloat(row.amountGrowth.replace(/[^\d.-]/g, ''));
      if (amountGrowthVal > 150) {
        outSheet.getRange(outputRow, 2, 1, 23).setBackground("#fff2cc");
        over150List.push({ merchant: row.merchant, currency: row.currency, type: row.type, growth: row.amountGrowth });
      } else if (amountGrowthVal < -60) {
        outSheet.getRange(outputRow, 2, 1, 23).setBackground("#ead1dc");
        below60List.push({ merchant: row.merchant, currency: row.currency, type: row.type, growth: row.amountGrowth });
      }
      outputRow++;
    });
  });

  SpreadsheetApp.flush();

  function buildTGMessage(sourceSheetName) {
    let msg = `#HTpay商戶月跑量分析\n來源表：${sourceSheetName}\n`;
    const growthToNumber = g => { const n = parseFloat(String(g || "").replace(/[^\d.-]/g, "")); return isNaN(n) ? -Infinity : n; };
    const over150Sorted = [...over150List].sort((a, b) => growthToNumber(b.growth) - growthToNumber(a.growth));
    const below60Sorted = [...below60List].sort((a, b) => growthToNumber(a.growth) - growthToNumber(b.growth));
    msg += "\n跑量超過 150%\n";
    if (over150Sorted.length === 0) msg += "無商戶\n";
    else over150Sorted.forEach((r, i) => { msg += `${i + 1}. ${r.merchant} - ${r.currency} - ${r.type} - ${r.growth}\n`; });
    msg += "\n跑量下降超過 60%\n";
    if (below60Sorted.length === 0) msg += "無商戶\n";
    else below60Sorted.forEach((r, i) => { msg += `${i + 1}. ${r.merchant} - ${r.currency} - ${r.type} - ${r.growth}\n`; });
    return msg.trim() || "無資料可傳送";
  }

  sendToTG(buildTGMessage(outSheetName));
}

// ============================================================
// SYSTEM 2 — Monthly Revenue Analysis (Payin / Payout cross-month)
// ============================================================

function generateRevenueAnalysis() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const indexSheet = ss.getSheetByName("Analytical Index");
  if (!indexSheet) throw new Error('找不到工作表：Analytical Index');

  const dataFileRaw = String(indexSheet.getRange("B1").getValue()).trim();
  const dataFileId = extractSpreadsheetId_(dataFileRaw);
  if (!dataFileId) throw new Error("B1 讀不到有效的 Spreadsheet ID。");

  const payinPrevName = String(indexSheet.getRange("B3").getValue()).trim();
  const payinCurrName = String(indexSheet.getRange("C3").getValue()).trim();
  const payoutPrevName = String(indexSheet.getRange("B4").getValue()).trim();
  const payoutCurrName = String(indexSheet.getRange("C4").getValue()).trim();

  const dataSS = SpreadsheetApp.openById(dataFileId);
  const payinPrevSheet = dataSS.getSheetByName(payinPrevName);
  const payinCurrSheet = dataSS.getSheetByName(payinCurrName);
  const payoutPrevSheet = dataSS.getSheetByName(payoutPrevName);
  const payoutCurrSheet = dataSS.getSheetByName(payoutCurrName);

  if (!payinPrevSheet) throw new Error(`找不到工作表：${payinPrevName}`);
  if (!payinCurrSheet) throw new Error(`找不到工作表：${payinCurrName}`);
  if (!payoutPrevSheet) throw new Error(`找不到工作表：${payoutPrevName}`);
  if (!payoutCurrSheet) throw new Error(`找不到工作表：${payoutCurrName}`);

  const prevMonth = extractMonth_(payinPrevName);
  const currMonth = extractMonth_(payinCurrName);
  const newSheetName = `${prevMonth}-${currMonth}月收益分析`;

  const existedSheet = ss.getSheetByName(newSheetName);
  if (existedSheet) { ss.toast(`已存在工作表「${newSheetName}」，本次不建立也不覆蓋。`, "提醒", 6); return; }

  const payinPrevData = buildSheetSummary_(payinPrevSheet, "payin");
  const payinCurrData = buildSheetSummary_(payinCurrSheet, "payin");
  const payoutPrevData = buildSheetSummary_(payoutPrevSheet, "payout");
  const payoutCurrData = buildSheetSummary_(payoutCurrSheet, "payout");

  const allCurrencies = Array.from(new Set([
    ...payinPrevData.currencies, ...payinCurrData.currencies,
    ...payoutPrevData.currencies, ...payoutCurrData.currencies
  ])).sort();

  if (allCurrencies.length === 0) throw new Error("來源資料中找不到任何幣別。");

  const resultSheet = ss.insertSheet(newSheetName);
  const header = [
    "幣別", "類型",
    `${prevMonth}月總筆數`, `${currMonth}月總筆數`, "總筆數-變化",
    `${prevMonth}月成功筆數`, `${currMonth}月成功筆數`, "成功筆數-變化",
    `${prevMonth}月整體跑量`, `${currMonth}月整體跑量`, "整體跑量-變化",
    `${prevMonth}月整體收益`, `${currMonth}月整體收益`, "整體收益-變化",
    "備註"
  ];
  resultSheet.getRange(1, 1, 1, header.length).setValues([header]);

  const rows = [];
  allCurrencies.forEach(currency => {
    const payinPrev = payinPrevData.byCurrency[currency] || emptyMetric_();
    const payinCurr = payinCurrData.byCurrency[currency] || emptyMetric_();
    const payoutPrev = payoutPrevData.byCurrency[currency] || emptyMetric_();
    const payoutCurr = payoutCurrData.byCurrency[currency] || emptyMetric_();

    rows.push([currency, "上分", payinPrev.depositCount, payinCurr.depositCount, "",
      payinPrev.depositSuccessCount, payinCurr.depositSuccessCount, "",
      payinPrev.depositAmount, payinCurr.depositAmount, "",
      payinPrev.depositRevenue, payinCurr.depositRevenue, "", ""]);
    rows.push(["", "回分", payoutPrev.payoutCount, payoutCurr.payoutCount, "",
      payoutPrev.payoutSuccessCount, payoutCurr.payoutSuccessCount, "",
      payoutPrev.payoutAmount, payoutCurr.payoutAmount, "",
      payoutPrev.payoutRevenue, payoutCurr.payoutRevenue, "", ""]);
    rows.push(["", "提現", payinPrev.withdrawCount, payinCurr.withdrawCount, "",
      payinPrev.withdrawSuccessCount, payinCurr.withdrawSuccessCount, "",
      payinPrev.withdrawAmount, payinCurr.withdrawAmount, "",
      payinPrev.withdrawRevenue, payinCurr.withdrawRevenue, "", ""]);
  });

  if (rows.length > 0) {
    resultSheet.getRange(2, 1, rows.length, 15).setValues(rows);
    for (let r = 2; r <= rows.length + 1; r++) {
      resultSheet.getRange(r, 5).setFormula(`=D${r}-C${r}`);
      resultSheet.getRange(r, 8).setFormula(`=G${r}-F${r}`);
      resultSheet.getRange(r, 11).setFormula(`=J${r}-I${r}`);
      resultSheet.getRange(r, 14).setFormula(`=M${r}-L${r}`);
    }
    resultSheet.getRange(2, 3, rows.length, 6).setNumberFormat("#,##0");
    resultSheet.getRange(2, 9, rows.length, 3).setNumberFormat("#,##0.00");
    resultSheet.getRange(2, 12, rows.length, 3).setNumberFormat("#,##0.00");
  }

  resultSheet.setFrozenRows(1);
  resultSheet.autoResizeColumns(1, 15);
  ss.toast(`已建立「${newSheetName}」`, "完成", 4);
}

function buildSheetSummary_(sheet, mode) {
  const values = sheet.getDataRange().getValues();
  if (!values || values.length < 2) return { currencies: [], byCurrency: {} };
  const headers = values[0].map(h => String(h).trim());
  const currencyIdx = headers.indexOf("币别");
  if (currencyIdx === -1) throw new Error(`找不到欄位「币别」(sheet: ${sheet.getName()})`);
  const colIndex = {};
  headers.forEach((h, i) => { colIndex[h] = i; });
  const byCurrency = {};

  for (let i = 1; i < values.length; i++) {
    const row = values[i];
    const currency = String(row[currencyIdx] || "").trim();
    if (!currency) continue;
    if (!byCurrency[currency]) byCurrency[currency] = emptyMetric_();
    const metric = byCurrency[currency];
    if (mode === "payin") {
      metric.depositCount += getNum_(row, colIndex, "代收总笔数");
      metric.depositAmount += getNum_(row, colIndex, "代收总金额");
      metric.depositSuccessCount += getOptionalNum_(row, colIndex, ["代收成功笔数", "成功代收笔数"]);
      metric.depositRevenue += getNum_(row, colIndex, "收益");
      metric.withdrawCount += getNum_(row, colIndex, "提现总笔数");
      metric.withdrawAmount += getNum_(row, colIndex, "提现总金额");
      metric.withdrawSuccessCount += getOptionalNum_(row, colIndex, ["提现成功笔数", "成功提现笔数"]);
      metric.withdrawRevenue += getNum_(row, colIndex, "提现总手续费");
    }
    if (mode === "payout") {
      metric.payoutCount += getNum_(row, colIndex, "代付总笔数");
      metric.payoutAmount += getNum_(row, colIndex, "代付总金额");
      metric.payoutSuccessCount += getOptionalNum_(row, colIndex, ["代付成功笔数", "成功代付笔数"]);
      metric.payoutRevenue += getNum_(row, colIndex, "收益");
    }
  }

  Object.keys(byCurrency).forEach(currency => {
    const m = byCurrency[currency];
    if (m.depositSuccessCount === 0 && m.depositCount !== 0) m.depositSuccessCount = m.depositCount;
    if (m.withdrawSuccessCount === 0 && m.withdrawCount !== 0) m.withdrawSuccessCount = m.withdrawCount;
    if (m.payoutSuccessCount === 0 && m.payoutCount !== 0) m.payoutSuccessCount = m.payoutCount;
  });

  return { currencies: Object.keys(byCurrency), byCurrency };
}

function emptyMetric_() {
  return { depositCount: 0, depositSuccessCount: 0, depositAmount: 0, depositRevenue: 0,
    payoutCount: 0, payoutSuccessCount: 0, payoutAmount: 0, payoutRevenue: 0,
    withdrawCount: 0, withdrawSuccessCount: 0, withdrawAmount: 0, withdrawRevenue: 0 };
}

function getNum_(row, colIndex, headerName) {
  if (!(headerName in colIndex)) return 0;
  return Number(row[colIndex[headerName]] || 0);
}

function getOptionalNum_(row, colIndex, headerNames) {
  for (var i = 0; i < headerNames.length; i++) {
    if (headerNames[i] in colIndex) return Number(row[colIndex[headerNames[i]]] || 0);
  }
  return 0;
}

function extractSpreadsheetId_(input) {
  if (/^[a-zA-Z0-9-_]{20,}$/.test(input) && !input.includes("/")) return input;
  const m = input.match(/\/d\/([a-zA-Z0-9-_]+)/);
  return m ? m[1] : "";
}

function extractMonth_(sheetName) {
  const m = String(sheetName).match(/(?:\d{4}[\/-])(\d{1,2})/);
  if (!m) throw new Error(`無法從工作表名稱解析月份：${sheetName}`);
  return String(Number(m[1]));
}

// ============================================================
// SYSTEM 3 — Revenue Ratio Report (per-merchant rev/amount ratio)
// ============================================================

function generateRevenueRatioReport() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var targetSheet = ss.getSheetByName("收益比");
  if (!targetSheet) { SpreadsheetApp.getUi().alert("找不到名為 '收益比' 的工作表！"); return; }

  var currentMonthStr = targetSheet.getRange("B1").getDisplayValue().toString().trim().replace("-", "/");
  var previousMonthStr = targetSheet.getRange("B2").getDisplayValue().toString().trim().replace("-", "/");
  if (!currentMonthStr || !previousMonthStr) { SpreadsheetApp.getUi().alert("請確保 B1 填入新月份、B2 填入舊月份（格式如：2026/05）！"); return; }

  var curParts = currentMonthStr.split("/");
  var prevParts = previousMonthStr.split("/");
  if (curParts.length < 2 || prevParts.length < 2) { SpreadsheetApp.getUi().alert("月份格式不正確"); return; }

  var curMonthName = parseInt(curParts[1], 10) + "月";
  var prevMonthName = parseInt(prevParts[1], 10) + "月";

  var lastColumn = targetSheet.getLastColumn();
  var lastRow = targetSheet.getLastRow();
  if (lastColumn >= 4 && lastRow > 0) {
    targetSheet.getRange(1, 4, lastRow, lastColumn - 3).clear();
    targetSheet.getRange(1, 4, lastRow, lastColumn - 3).clearFormat();
  }

  // CONFIG_DATA_SHEET_ID — replace with your actual Data spreadsheet ID
  var dataSheetId = "CONFIG_DATA_SHEET_ID";
  var dataSpreadsheet;
  try { dataSpreadsheet = SpreadsheetApp.openById(dataSheetId); }
  catch (e) { SpreadsheetApp.getUi().alert("無法開啟 Data 工作簿，請確認 ID 是否正確。"); return; }

  var curPayinSheet = dataSpreadsheet.getSheetByName(currentMonthStr + "-Payin");
  var curPayoutSheet = dataSpreadsheet.getSheetByName(currentMonthStr + "-Payout");
  var prevPayinSheet = dataSpreadsheet.getSheetByName(previousMonthStr + "-Payin");
  var prevPayoutSheet = dataSpreadsheet.getSheetByName(previousMonthStr + "-Payout");

  if (!curPayinSheet || !curPayoutSheet || !prevPayinSheet || !prevPayoutSheet) {
    SpreadsheetApp.getUi().alert("Data 工作簿中缺少比對所需的標籤頁！"); return;
  }

  var categorizedData = {};
  function processSheetData(sheet, typeKey, isCurrentMonth) {
    var values = sheet.getDataRange().getValues();
    for (var i = 1; i < values.length; i++) {
      var row = values[i];
      var merchant = row[3] ? row[3].toString().trim() : "";
      var currencyCol = (typeKey === "上分") ? 5 : 4;
      var currency = row[currencyCol] ? row[currencyCol].toString().trim() : "";
      if (!merchant || !currency) continue;
      var amt = parseFloat(row[6]) || 0;
      var revCol = (typeKey === "上分") ? 15 : 8;
      var rev = parseFloat(row[revCol]) || 0;
      if (!categorizedData[currency]) categorizedData[currency] = { "上分": {}, "回分": {} };
      if (!categorizedData[currency][typeKey][merchant]) categorizedData[currency][typeKey][merchant] = { curAmt: 0, prevAmt: 0, curRev: 0, prevRev: 0 };
      if (isCurrentMonth) { categorizedData[currency][typeKey][merchant].curAmt += amt; categorizedData[currency][typeKey][merchant].curRev += rev; }
      else { categorizedData[currency][typeKey][merchant].prevAmt += amt; categorizedData[currency][typeKey][merchant].prevRev += rev; }
    }
  }

  processSheetData(curPayinSheet, "上分", true);
  processSheetData(curPayoutSheet, "回分", true);
  processSheetData(prevPayinSheet, "上分", false);
  processSheetData(prevPayoutSheet, "回分", false);

  var currentCol = 4;
  var headers = [
    "商戶名稱",
    curMonthName + " 提單成功金額", prevMonthName + " 提單成功金額", "提單成功金額 變化",
    curMonthName + " 收益", prevMonthName + " 收益", "收益 變化",
    curMonthName + " 收益比", prevMonthName + " 收益比", "收益比 變化 (pp)"
  ];

  for (var currency in categorizedData) {
    var currentRow = 1;
    targetSheet.getRange(currentRow, currentCol).setValue(currency).setFontWeight("bold").setFontSize(14);
    currentRow++;
    targetSheet.getRange(currentRow, currentCol, 1, headers.length).setValues([headers])
      .setFontWeight("bold").setBackground("#E0E0E0")
      .setBorder(true, true, true, true, true, true, "#BDBDBD", SpreadsheetApp.BorderStyle.SOLID);
    currentRow++;

    ["上分", "回分"].forEach(subType => {
      var merchantMap = categorizedData[currency][subType];
      var merchantNames = Object.keys(merchantMap);
      if (merchantNames.length === 0) return;
      var startRowIndex = currentRow;

      merchantNames.forEach(mName => {
        var item = merchantMap[mName];
        var amountChange = item.prevAmt !== 0 ? (item.curAmt - item.prevAmt) / item.prevAmt : 0;
        var revChange = item.prevRev !== 0 ? (item.curRev - item.prevRev) / item.prevRev : 0;
        var curRatio = item.curAmt !== 0 ? item.curRev / item.curAmt : 0;
        var prevRatio = item.prevAmt !== 0 ? item.prevRev / item.prevAmt : 0;
        var ratioChange = curRatio - prevRatio;

        targetSheet.getRange(currentRow, currentCol, 1, headers.length).setValues([[
          mName, item.curAmt, item.prevAmt, amountChange,
          item.curRev, item.prevRev, revChange,
          curRatio, prevRatio, ratioChange
        ]]);

        setCellColorAndFormat(targetSheet.getRange(currentRow, currentCol + 3), amountChange, "0.00%");
        setCellColorAndFormat(targetSheet.getRange(currentRow, currentCol + 6), revChange, "0.00%");
        setCellColorAndFormat(targetSheet.getRange(currentRow, currentCol + 9), ratioChange, "0.00%");
        currentRow++;
      });

      var numRows = currentRow - startRowIndex;
      if (numRows > 0) {
        targetSheet.getRange(startRowIndex, currentCol + 1, numRows, 2).setNumberFormat("#,##0.00");
        targetSheet.getRange(startRowIndex, currentCol + 4, numRows, 2).setNumberFormat("#,##0.00");
        targetSheet.getRange(startRowIndex, currentCol + 7, numRows, 2).setNumberFormat("0.00%");
        targetSheet.getRange(startRowIndex, currentCol, numRows, headers.length)
          .setBorder(true, true, true, true, true, true, "#E0E0E0", SpreadsheetApp.BorderStyle.SOLID);
      }

      if (subType === "上分" && Object.keys(categorizedData[currency]["回分"]).length > 0) currentRow++;
    });

    currentCol += headers.length + 1;
  }

  SpreadsheetApp.getUi().alert("【動態對比 + 漲跌著色】報告已產生完畢！");
}

function setCellColorAndFormat(range, value, numberFormat) {
  if (value > 0) {
    range.setBackground("#E8F5E9");
    range.setFontColor("#2E7D32");
    range.setNumberFormat(numberFormat + " '▲'");
  } else if (value < 0) {
    range.setBackground("#FFEBEE");
    range.setFontColor("#C62828");
    range.setNumberFormat(numberFormat + " '▼'");
  } else {
    range.setNumberFormat(numberFormat);
  }
}
