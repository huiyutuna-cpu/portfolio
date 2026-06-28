/**
 * Case 06 — BestPay Monthly Agent Commission Report Generator
 *
 * Reads pivot tables (Payin-DD / Payout-DD) + merchant-to-agent mapping (代理名單).
 * Generates one tab per agent: Payin block (left) | Payout block (right).
 * Full month calendar always shown — missing days filled with 0.
 * INR-specific: Payout fee base = Amount ÷ 1.04 (Net Subtotal column).
 */

const SHEET_PAYIN     = "Payin-DD";
const SHEET_PAYOUT    = "Payout-DD";
const SHEET_AGENT_MAP = "代理名單";
const PAYOUT_DIVISOR  = 1.04; // BestPay INR: amounts include GST (×1.04); fee is calculated on net

function runGenerate() { generateAgentSheets(); }

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu("代理報表")
    .addItem("產生/更新代理工作表", "runGenerate")
    .addToUi();
}

function generateAgentSheets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();

  const payinSheet  = ss.getSheetByName(SHEET_PAYIN);
  const payoutSheet = ss.getSheetByName(SHEET_PAYOUT);
  const mapSheet    = ss.getSheetByName(SHEET_AGENT_MAP);
  if (!payinSheet || !payoutSheet || !mapSheet)
    throw new Error(`找不到必要工作表：${SHEET_PAYIN}, ${SHEET_PAYOUT}, ${SHEET_AGENT_MAP}`);

  const mapping = loadAgentMapping_(mapSheet);

  // Graceful degradation: if one side is missing headers, prompt and continue with the other
  let payinPivot = { rows: [] }, payoutPivot = { rows: [] };

  for (const [sheetRef, pivotRef, label] of [
    [payinSheet, "payinPivot", "Payin-DD"],
    [payoutSheet, "payoutPivot", "Payout-DD"]
  ]) {
    try {
      if (label === "Payin-DD")  payinPivot  = readPivot_(sheetRef);
      else                       payoutPivot = readPivot_(sheetRef);
    } catch (e) {
      if (String(e.message).includes("找不到 Date/Transaction Date 表頭列")) {
        const r = ui.alert(`${label} 讀取失敗`, `${label} 找不到表頭列。\n\n按「確認」忽略並繼續。`, ui.ButtonSet.OK_CANCEL);
        if (r !== ui.Button.OK) throw e;
      } else { throw e; }
    }
  }

  const monthSourceRows = payinPivot.rows.length ? payinPivot.rows
    : payoutPivot.rows.length ? payoutPivot.rows : [];

  if (!monthSourceRows.length) {
    ui.alert("沒有可用日期資料", "Payin-DD 與 Payout-DD 都沒有可用的日期資料。", ui.ButtonSet.OK);
    return;
  }

  const { year, month0 } = pickMainMonth_(monthSourceRows);
  const monthDates = buildMonthDates_(year, month0);

  const agents = new Set(Object.values(mapping.merchantToAgent).filter(Boolean));
  agents.forEach(agentName => {
    upsertAgentSheet_(ss, agentName, payinPivot, payoutPivot, mapping, monthDates);
  });
}

// ==================== Mapping loader ====================

function loadAgentMapping_(sheet) {
  const values = sheet.getDataRange().getValues();
  if (values.length < 2) throw new Error("代理名單沒有資料");

  const header      = values[0].map(v => String(v).trim());
  const idxMerchant = header.findIndex(h => h === "商戶名");
  const idxAgent    = header.findIndex(h => h === "代理名");
  const idxPayin    = header.findIndex(h => h === "代理代收手續費(%)");
  const idxPayout   = header.findIndex(h => h === "代理代付手續費(%)");

  if ([idxMerchant, idxAgent, idxPayin, idxPayout].some(i => i < 0))
    throw new Error("代理名單欄位名稱不符合（商戶名 / 代理名 / 代理代收手續費(%) / 代理代付手續費(%)）");

  const merchantToAgent = {}, merchantRatePayin = {}, merchantRatePayout = {};
  for (let r = 1; r < values.length; r++) {
    const row      = values[r];
    const merchant = String(row[idxMerchant] || "").trim();
    const agent    = String(row[idxAgent]    || "").trim();
    if (!merchant || !agent) continue;

    const key = normalizeKey_(merchant);
    merchantToAgent[key] = agent;
    const pi = parseRatePercent_(row[idxPayin]);
    const po = parseRatePercent_(row[idxPayout]);
    if (pi !== null) merchantRatePayin[key]  = pi;
    if (po !== null) merchantRatePayout[key] = po;
  }
  return { merchantToAgent, merchantRatePayin, merchantRatePayout };
}

// ==================== Pivot reader ====================

/**
 * Finds header row (scanning first 20 rows for "date" or "transaction date"),
 * reads date rows until hitting "總和", maps merchant columns to keyed values.
 */
function readPivot_(sheet) {
  const all = sheet.getDataRange().getValues();
  const headerRowIndex = findHeaderRowIndex_(all);
  if (headerRowIndex < 0) throw new Error(`${sheet.getName()} 找不到 Date/Transaction Date 表頭列`);

  const headers = all[headerRowIndex].map(v => String(v).trim());
  const rows = [];

  for (let r = headerRowIndex + 1; r < all.length; r++) {
    const first = String(all[r][0] || "").trim();
    if (!first) continue;
    if (first === "總和") break;

    const dateObj = parseAnyDate_(all[r][0]);
    if (!dateObj) continue;

    const valuesByMerchantKey = {};
    for (let c = 1; c < headers.length; c++) {
      const merchant = String(headers[c] || "").trim();
      if (!merchant || merchant === "總和") continue;
      const key = normalizeKey_(merchant);
      const num = parseNumber_(all[r][c]);
      if (num !== 0) valuesByMerchantKey[key] = (valuesByMerchantKey[key] || 0) + num;
    }
    rows.push({ date: dateObj, valuesByMerchantKey });
  }
  return { rows };
}

// ==================== Agent sheet writer ====================

function upsertAgentSheet_(ss, agentName, payinPivot, payoutPivot, mapping, monthDates) {
  let sh = ss.getSheetByName(agentName) || ss.insertSheet(agentName);
  sh.clear();

  const merchantsForAgent = Object.entries(mapping.merchantToAgent)
    .filter(([, a]) => a === agentName).map(([k]) => k).sort();

  writeSection_({ sheet: sh, title: "Payin",  startRow: 1, startCol: 1,
    pivot: payinPivot,  merchantsForAgent, rateMap: mapping.merchantRatePayin,
    feeBaseDivisor: 1, showNetSubtotal: false, dateList: monthDates });

  const payinBlockWidth = merchantsForAgent.length + 3; // Date + merchants + Subtotal + Fee
  writeSection_({ sheet: sh, title: "Payout", startRow: 1, startCol: payinBlockWidth + 3,
    pivot: payoutPivot, merchantsForAgent, rateMap: mapping.merchantRatePayout,
    feeBaseDivisor: PAYOUT_DIVISOR, showNetSubtotal: true, dateList: monthDates });

  sh.autoResizeColumns(1, Math.min(payinBlockWidth + merchantsForAgent.length + 7, 60));
}

function writeSection_({ sheet, title, startRow, startCol, pivot, merchantsForAgent,
                          rateMap, feeBaseDivisor = 1, showNetSubtotal = false, dateList = [] }) {
  const tz = Session.getScriptTimeZone();
  sheet.getRange(startRow++, startCol).setValue(title).setFontWeight("bold");

  const header = showNetSubtotal
    ? ["Date", ...merchantsForAgent, "Subtotal", "Net Subtotal", "Fee"]
    : ["Date", ...merchantsForAgent, "Subtotal", "Fee"];
  sheet.getRange(startRow++, startCol, 1, header.length).setValues([header]).setFontWeight("bold");

  const rowMap = new Map();
  for (const r of pivot.rows) rowMap.set(dateKey_(r.date, tz), r);

  // Full calendar: every day in the month, zero-fill if no data
  const out = dateList.map(d => {
    const r    = rowMap.get(dateKey_(d, tz));
    const vals = r ? r.valuesByMerchantKey : {};
    const row  = [formatDate_(d, tz)];
    let subtotal = 0, netSubtotal = 0, fee = 0;

    for (const mk of merchantsForAgent) {
      const amt  = vals[mk] || 0;
      const base = amt / feeBaseDivisor;
      row.push(amt);
      subtotal    += amt;
      netSubtotal += base;
      const rate = Object.prototype.hasOwnProperty.call(rateMap, mk) ? rateMap[mk] : null;
      if (rate !== null) fee += base * (rate / 100);
    }

    showNetSubtotal ? row.push(subtotal, netSubtotal, fee) : row.push(subtotal, fee);
    return row;
  });

  if (out.length > 0) {
    sheet.getRange(startRow, startCol, out.length, header.length).setValues(out);
    sheet.getRange(startRow, startCol, out.length, 1).setNumberFormat("yyyy/MM/dd");
    sheet.getRange(startRow, startCol + 1, out.length, header.length - 1).setNumberFormat("₹ #,##0.00");
  }

  // Total row
  const totalRow  = startRow + out.length;
  const subtotalC = startCol + header.indexOf("Subtotal");
  const feeC      = startCol + header.indexOf("Fee");
  sheet.getRange(totalRow, startCol).setValue("Total").setFontWeight("bold");

  if (out.length > 0) {
    sheet.getRange(totalRow, subtotalC)
      .setFormulaR1C1(`=SUM(R${startRow}C${subtotalC}:R${totalRow-1}C${subtotalC})`)
      .setNumberFormat("₹ #,##0.00").setFontWeight("bold");

    if (showNetSubtotal) {
      const netC = startCol + header.indexOf("Net Subtotal");
      sheet.getRange(totalRow, netC)
        .setFormulaR1C1(`=SUM(R${startRow}C${netC}:R${totalRow-1}C${netC})`)
        .setNumberFormat("₹ #,##0.00").setFontWeight("bold");
    }

    sheet.getRange(totalRow, feeC)
      .setFormulaR1C1(`=SUM(R${startRow}C${feeC}:R${totalRow-1}C${feeC})`)
      .setNumberFormat("₹ #,##0.00").setFontWeight("bold");
  }
}

// ==================== Utilities ====================

function pickMainMonth_(rows) {
  const counts = new Map();
  for (const r of rows) {
    const k = `${r.date.getFullYear()}-${String(r.date.getMonth() + 1).padStart(2, "0")}`;
    counts.set(k, (counts.get(k) || 0) + 1);
  }
  let bestKey = null, bestCnt = -1;
  for (const [k, c] of counts) if (c > bestCnt) { bestCnt = c; bestKey = k; }
  const [yy, mm] = bestKey.split("-").map(Number);
  return { year: yy, month0: mm - 1 };
}

function buildMonthDates_(year, month0) {
  const lastDay = new Date(year, month0 + 1, 0).getDate();
  return Array.from({ length: lastDay }, (_, i) => new Date(year, month0, i + 1));
}

function findHeaderRowIndex_(grid) {
  const maxScan = Math.min(20, grid.length);
  for (let r = 0; r < maxScan; r++) {
    const row = grid[r].map(v => String(v).trim().toLowerCase());
    if (row.includes("date") || row.includes("transaction date")) return r;
  }
  return -1;
}

function dateKey_(d, tz) { return Utilities.formatDate(d, tz, "yyyy-MM-dd"); }
function formatDate_(d, tz) { return Utilities.formatDate(d, tz, "yyyy/MM/dd"); }
function normalizeKey_(s) { return String(s || "").trim().toLowerCase(); }

function parseNumber_(v) {
  if (v === null || v === undefined || v === "") return 0;
  if (typeof v === "number") return v;
  const n = Number(String(v).replace(/,/g, "").replace(/[^\d.-]/g, "").trim());
  return isNaN(n) ? 0 : n;
}

function parseRatePercent_(v) {
  if (v === null || v === undefined || v === "") return null;
  if (typeof v === "number") return v;
  const n = Number(String(v).trim());
  return isNaN(n) ? null : n;
}

function parseAnyDate_(v) {
  if (!v) return null;
  if (Object.prototype.toString.call(v) === "[object Date]" && !isNaN(v.getTime())) return v;
  const s = String(v).trim();
  let m = s.match(/^(\d{4})-(\d{2})-(\d{2})$/);
  if (m) return new Date(+m[1], +m[2] - 1, +m[3]);
  m = s.match(/^(\d{2})-(\d{2})-(\d{4})$/);
  if (m) return new Date(+m[3], +m[2] - 1, +m[1]);
  const d = new Date(s);
  return isNaN(d.getTime()) ? null : d;
}
