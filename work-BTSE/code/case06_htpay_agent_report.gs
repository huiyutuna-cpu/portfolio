/**
 * Case 06 — HTPay Monthly Agent Commission Report Generator
 *
 * One button → generates one tab per agent, one horizontal block per currency.
 * Each block: Pay In section → Pay Out section → Total row.
 * Also generates 總計 (all data by Month+Currency) and 代理總計 (agent fees summary).
 *
 * Source sheets required: PayIn-Data, Payout-Data
 * Chinese-name agents are excluded from individual tabs but included in 總計.
 */

function generateAgentReports() {
  const ss = SpreadsheetApp.getActive();
  const { payinRows, payoutRows } = readSourceRowsForAgentSheets_();

  const agents = new Set();
  payinRows.forEach(r => agents.add(r.agentName));
  payoutRows.forEach(r => agents.add(r.agentName));
  let agentList = Array.from(agents).filter(Boolean).sort();

  if (agentList.length === 0) {
    buildSummarySheetFromSource_();
    buildAgentTotalSheet_(payinRows, payoutRows);
    SpreadsheetApp.getUi().alert("沒有可產生的代理工作表（可能全部是中文代理），但已更新「總計」與「代理總計」。");
    return;
  }

  const payinByAgent  = groupBy_(payinRows,  r => r.agentName);
  const payoutByAgent = groupBy_(payoutRows, r => r.agentName);

  const blockWidth = 5; // Date/Month · accountId · total_count · total_amount · total_agentFee
  const gap = 2;

  agentList.forEach(agentName => {
    const agentPayin  = payinByAgent.get(agentName)  || [];
    const agentPayout = payoutByAgent.get(agentName) || [];

    const currencies = new Set();
    agentPayin.forEach(r  => currencies.add(r.currency));
    agentPayout.forEach(r => currencies.add(r.currency));
    const currencyList = Array.from(currencies).filter(Boolean).sort();
    if (!currencyList.length) return;

    const sheetName = sanitizeSheetName_(agentName);
    let sh = ss.getSheetByName(sheetName);
    if (!sh) sh = ss.insertSheet(sheetName);
    sh.clear();

    currencyList.forEach((ccy, idx) => {
      const col0 = 1 + idx * (blockWidth + gap);
      writeCurrencyBlock_(sh, 1, col0, ccy,
        aggregateRowsByDateAccount_(agentPayin.filter(r => r.currency === ccy)),
        aggregateRowsByDateAccount_(agentPayout.filter(r => r.currency === ccy))
      );
      for (let c = col0; c < col0 + blockWidth; c++) sh.autoResizeColumn(c);
    });
  });

  buildSummarySheetFromSource_();
  buildAgentTotalSheet_(payinRows, payoutRows);

  SpreadsheetApp.getUi().alert(`完成：已產生/更新 ${agentList.length} 個代理工作表 + 總計 + 代理總計。`);
}

// ==================== Data reading ====================

/**
 * Reads PayIn-Data and Payout-Data.
 * Excludes rows where agentName contains Chinese characters (internal accounts).
 */
function readSourceRowsForAgentSheets_() {
  const ss = SpreadsheetApp.getActive();
  const payinSh  = ss.getSheetByName("PayIn-Data");
  const payoutSh = ss.getSheetByName("Payout-Data");
  if (!payinSh || !payoutSh) throw new Error("找不到工作表：PayIn-Data 或 Payout-Data");

  const payinVals  = payinSh.getDataRange().getValues();
  const payoutVals = payoutSh.getDataRange().getValues();

  const payinRows = [];
  for (let i = 1; i < payinVals.length; i++) {
    const row = payinVals[i];
    const agentName = normalizeText_(row[11] || ""); // col L
    const currency  = normalizeText_(row[5]  || ""); // col F
    if (!agentName || !currency || hasChinese_(agentName)) continue;

    payinRows.push({
      date:           formatYearMonth_(row[0]),   // col A → YYYY-MM
      accountId:      row[3],                     // col D
      total_count:    toNumber_(row[4]),           // col E
      total_amount:   toNumber_(row[6]),           // col G
      total_agentFee: toNumber_(row[12]),          // col M
      agentName, currency
    });
  }

  const payoutRows = [];
  for (let i = 1; i < payoutVals.length; i++) {
    const row = payoutVals[i];
    const agentName = normalizeText_(row[9]  || ""); // col J
    const currency  = normalizeText_(row[4]  || ""); // col E
    if (!agentName || !currency || hasChinese_(agentName)) continue;

    payoutRows.push({
      date:           formatYearMonth_(row[0]),   // col A → YYYY-MM
      accountId:      row[3],                     // col D
      total_count:    toNumber_(row[5]),           // col F
      total_amount:   toNumber_(row[6]),           // col G
      total_agentFee: toNumber_(row[10]),          // col K
      agentName, currency
    });
  }

  return { payinRows, payoutRows };
}

// ==================== Sheet writers ====================

function writeCurrencyBlock_(sh, startRow, startCol, currency, payinAgg, payoutAgg) {
  const headers = ["Date/Month", "accountId", "total_count", "total_amount", "total_agentFee"];

  sh.getRange(startRow,     startCol).setValue("HTPay 代理佣金报表").setFontWeight("bold");
  sh.getRange(startRow + 1, startCol).setValue(currency).setFontWeight("bold");

  let r = startRow + 2;
  sh.getRange(r++, startCol).setValue("Pay In").setFontWeight("bold");
  sh.getRange(r++, startCol, 1, headers.length).setValues([headers]).setFontWeight("bold");
  const payinRes = writeDataTable_(sh, r, startCol, payinAgg, "Payin Total");
  r = payinRes.nextRow + 2;

  sh.getRange(r++, startCol).setValue("Pay Out").setFontWeight("bold");
  sh.getRange(r++, startCol, 1, headers.length).setValues([headers]).setFontWeight("bold");
  const payoutRes = writeDataTable_(sh, r, startCol, payoutAgg, "Payout Total");
  r = payoutRes.nextRow + 1;

  // Total = Payin Total + Payout Total (cross-references via formula)
  sh.getRange(r, startCol).setValue("Total").setFontWeight("bold");
  [[2, "#,##0"], [3, "#,##0.00"], [4, "#,##0.00"]].forEach(([offset, fmt]) => {
    const col = colA1_(startCol + offset);
    sh.getRange(r, startCol + offset)
      .setFormula(`=${col}${payinRes.totalRow}+${col}${payoutRes.totalRow}`)
      .setNumberFormat(fmt).setFontWeight("bold");
  });
}

function writeDataTable_(sh, startRow, startCol, rows, totalLabel) {
  const n = rows.length;
  if (n > 0) {
    sh.getRange(startRow, startCol, n, 5).setValues(
      rows.map(r => [r.date, r.accountId, r.total_count, r.total_amount, r.total_agentFee])
    );
    sh.getRange(startRow, startCol + 2, n, 1).setNumberFormat("#,##0");
    sh.getRange(startRow, startCol + 3, n, 2).setNumberFormat("#,##0.00");
  }

  const totalRow = startRow + Math.max(n, 1);
  sh.getRange(totalRow, startCol).setValue(totalLabel).setFontWeight("bold");

  if (n > 0) {
    [[2, "#,##0"], [3, "#,##0.00"], [4, "#,##0.00"]].forEach(([offset, fmt]) => {
      const col = colA1_(startCol + offset);
      sh.getRange(totalRow, startCol + offset)
        .setFormula(`=SUM(${col}${startRow}:${col}${startRow + n - 1})`)
        .setNumberFormat(fmt).setFontWeight("bold");
    });
  } else {
    sh.getRange(totalRow, startCol + 2, 1, 3).setValues([[0, 0, 0]]);
    sh.getRange(totalRow, startCol + 2).setNumberFormat("#,##0").setFontWeight("bold");
    sh.getRange(totalRow, startCol + 3, 1, 2).setNumberFormat("#,##0.00").setFontWeight("bold");
  }

  return { nextRow: totalRow + 1, totalRow };
}

// ==================== Summary sheets ====================

/**
 * 總計: reads ALL rows from source (including Chinese-name agents) for Finance reconciliation.
 */
function buildSummarySheetFromSource_() {
  const ss = SpreadsheetApp.getActive();
  const payinSh  = ss.getSheetByName("PayIn-Data");
  const payoutSh = ss.getSheetByName("Payout-Data");
  if (!payinSh || !payoutSh) throw new Error("找不到工作表：PayIn-Data 或 Payout-Data");

  const payinAgg  = aggregateSourceByMonthCurrency_(payinSh.getDataRange().getValues(),  "PAYIN");
  const payoutAgg = aggregateSourceByMonthCurrency_(payoutSh.getDataRange().getValues(), "PAYOUT");

  let sh = ss.getSheetByName("總計") || ss.insertSheet("總計");
  sh.clear();

  const headers = ["Date/Month", "Currency", "total_count", "total_amount", "total_agentFee"];
  let r = 1;

  sh.getRange(r++, 1).setValue("Pay In").setFontWeight("bold");
  sh.getRange(r++, 1, 1, headers.length).setValues([headers]).setFontWeight("bold");
  if (payinAgg.length > 0) {
    sh.getRange(r, 1, payinAgg.length, 5).setValues(payinAgg.map(x => [x.date, x.currency, x.total_count, x.total_amount, x.total_agentFee]));
    sh.getRange(r, 3, payinAgg.length, 1).setNumberFormat("#,##0");
    sh.getRange(r, 4, payinAgg.length, 2).setNumberFormat("#,##0.00");
    r += payinAgg.length;
  }

  r += 2;
  sh.getRange(r++, 1).setValue("Pay Out").setFontWeight("bold");
  sh.getRange(r++, 1, 1, headers.length).setValues([headers]).setFontWeight("bold");
  if (payoutAgg.length > 0) {
    sh.getRange(r, 1, payoutAgg.length, 5).setValues(payoutAgg.map(x => [x.date, x.currency, x.total_count, x.total_amount, x.total_agentFee]));
    sh.getRange(r, 3, payoutAgg.length, 1).setNumberFormat("#,##0");
    sh.getRange(r, 4, payoutAgg.length, 2).setNumberFormat("#,##0.00");
  }

  for (let c = 1; c <= 5; c++) sh.autoResizeColumn(c);
}

function buildAgentTotalSheet_(payinRows, payoutRows) {
  let sh = SpreadsheetApp.getActive().getSheetByName("代理總計") || SpreadsheetApp.getActive().insertSheet("代理總計");
  sh.clear();

  const agg = aggregateByMonthAgentCurrency_(payinRows, payoutRows);
  const headers = ["Date/Month", "Agent", "", "Agent Fee", "Currency"];
  let r = 1;
  sh.getRange(r++, 1).setValue("代理總計").setFontWeight("bold");
  sh.getRange(r++, 1, 1, headers.length).setValues([headers]).setFontWeight("bold");
  if (agg.length > 0) {
    sh.getRange(r, 1, agg.length, 5).setValues(agg.map(x => [x.date, x.agentName, "", x.agentFee, x.currency]));
    sh.getRange(r, 4, agg.length, 1).setNumberFormat("#,##0.00");
  }
  for (let c = 1; c <= 5; c++) sh.autoResizeColumn(c);
}

// ==================== Aggregation helpers ====================

function aggregateRowsByDateAccount_(rows) {
  const m = new Map();
  rows.forEach(r => {
    const k = `${String(r.date)}||${String(r.accountId)}`;
    if (!m.has(k)) { m.set(k, { ...r }); }
    else {
      const t = m.get(k);
      t.total_count    += r.total_count;
      t.total_amount   += r.total_amount;
      t.total_agentFee += r.total_agentFee;
    }
  });
  return Array.from(m.values()).sort((a, b) => {
    const d = String(a.date).localeCompare(String(b.date));
    return d !== 0 ? d : String(a.accountId).localeCompare(String(b.accountId));
  });
}

function aggregateSourceByMonthCurrency_(values, type) {
  const m = new Map();
  for (let i = 1; i < values.length; i++) {
    const row = values[i];
    const date = formatYearMonth_(row[0]);
    const currency = normalizeText_((type === "PAYIN" ? row[5] : row[4]) || "");
    if (!date || !currency) continue;

    const count  = toNumber_(type === "PAYIN" ? row[4] : row[5]);
    const amount = toNumber_(row[6]);
    const fee    = toNumber_(type === "PAYIN" ? row[12] : row[10]);

    const key = `${date}||${currency}`;
    if (!m.has(key)) m.set(key, { date, currency, total_count: 0, total_amount: 0, total_agentFee: 0 });
    const t = m.get(key);
    t.total_count += count; t.total_amount += amount; t.total_agentFee += fee;
  }
  return Array.from(m.values()).sort((a, b) => {
    const d = a.date.localeCompare(b.date);
    return d !== 0 ? d : a.currency.localeCompare(b.currency);
  });
}

function aggregateByMonthAgentCurrency_(payinRows, payoutRows) {
  const m = new Map();
  [...payinRows, ...payoutRows].forEach(r => {
    const k = `${r.date}||${r.agentName}||${r.currency}`;
    if (!m.has(k)) m.set(k, { date: r.date, agentName: r.agentName, currency: r.currency, agentFee: 0 });
    m.get(k).agentFee += Number(r.total_agentFee || 0);
  });
  return Array.from(m.values()).sort((a, b) => {
    const d = a.date.localeCompare(b.date);
    if (d !== 0) return d;
    const ag = a.agentName.localeCompare(b.agentName);
    return ag !== 0 ? ag : a.currency.localeCompare(b.currency);
  });
}

// ==================== Utilities ====================

function groupBy_(arr, keyFn) {
  const m = new Map();
  arr.forEach(x => { const k = keyFn(x); if (!m.has(k)) m.set(k, []); m.get(k).push(x); });
  return m;
}

function toNumber_(v) {
  if (v === null || v === "" || v === undefined) return 0;
  const n = Number(v); return isNaN(n) ? 0 : n;
}

function normalizeText_(s) {
  return String(s).replace(/ /g, " ").replace(/　/g, " ").replace(/\s+/g, " ").trim();
}

function sanitizeSheetName_(name) {
  return String(name).replace(/[\[\]\:\*\?\/\\]/g, " ").trim().slice(0, 99) || "Agent";
}

// Returns A1-style column letter for 1-based column number
function colA1_(col) {
  let s = "", n = col;
  while (n > 0) { const r = (n - 1) % 26; s = String.fromCharCode(65 + r) + s; n = Math.floor((n - 1) / 26); }
  return s;
}

function hasChinese_(s) {
  return /[㐀-䶿一-鿿]/.test(String(s || ""));
}

function formatYearMonth_(v) {
  if (!v && v !== 0) return "";
  if (Object.prototype.toString.call(v) === "[object Date]" && !isNaN(v.getTime())) {
    return Utilities.formatDate(v, Session.getScriptTimeZone(), "yyyy-MM");
  }
  const s = String(v).trim();
  const m = s.match(/^(\d{4})[-\/](\d{1,2})/);
  return m ? `${m[1]}-${("0" + m[2]).slice(-2)}` : s;
}
