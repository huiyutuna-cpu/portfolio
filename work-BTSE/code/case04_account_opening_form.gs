/***** ====== CONFIG ====== *****/
const THIRD_BOOK_URL = 'CONFIG_THIRD_BOOK_URL';
const STATION_SHEET_NAME = '商戶＆代理';
const STATION_COL_INDEX = 1;
const TIMEZONE = 'Asia/Taipei';

const TELEGRAM_BOT_TOKEN = 'CONFIG_BOT_TOKEN';
const NOTIFY_CHAT_IDS = ['CONFIG_CHAT_ID_1', 'CONFIG_CHAT_ID_2'];

/***** ====== WEB APP ENTRY ====== *****/
function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index');
}

/***** ====== UTILITIES ====== *****/
function normalizeToArray_(v) {
  if (Array.isArray(v)) return v;
  if (!v) return [];
  const s = String(v).trim();
  if (!s) return [];
  return s.split(/[\n,]/).map(x => x.trim()).filter(Boolean);
}

function cleanNumber_(raw) {
  if (raw === null || raw === undefined) return '';
  const n = Number(String(raw).replace(/[^\d.\-]/g, ''));
  return isFinite(n) ? n : '';
}

function prettyCurrencyLabel_(raw) {
  const s = String(raw || '').trim();
  if (!s) return '';
  const key = s.toLowerCase();
  if (key === 'crypto' || key === 'cy') return 'Crypto';
  if (/^[a-z]{3}$/i.test(s)) return s.toUpperCase();
  return s;
}

function findFirstEmptyABRow_(sheet) {
  const lastRow = sheet.getLastRow();
  if (lastRow === 0) return 1;
  const data = sheet.getRange(1, 1, lastRow, 2).getValues();
  for (let i = 0; i < data.length; i++) {
    if (!data[i][0] && !data[i][1]) return i + 1;
  }
  return lastRow + 1;
}

function notifyGroups_(text) {
  const token = TELEGRAM_BOT_TOKEN;
  if (!token || token === 'CONFIG_BOT_TOKEN') return;
  const url = `https://api.telegram.org/bot${token}/sendMessage`;
  NOTIFY_CHAT_IDS.forEach(id => {
    try {
      UrlFetchApp.fetch(url, {
        method: 'post',
        contentType: 'application/json',
        payload: JSON.stringify({ chat_id: id, text, disable_web_page_preview: true })
      });
    } catch (e) {
      Logger.log(`[Telegram Error] ${e}`);
    }
  });
}

/***** ====== STATION ID GENERATION ====== *****/
// Prefix: CY for crypto-only, 2-char fiat code for single fiat, MA for multi-fiat
function generateStationId_(data) {
  const prefix = decidePrefix_(data);
  const nextNum = findNextNumberForPrefix_(prefix);
  const number4 = String(nextNum).padStart(4, '0');
  return prefix + number4;
}

function decidePrefix_(data) {
  const ALIAS_MAP = { 'RMB': 'CNY', 'CN¥': 'CNY', 'YUAN': 'CNY', 'INR': 'INR', 'RUPEE': 'INR' };
  const CRYPTO_SET = new Set(['CRYPTO','CY','USDT','USDC','BTC','ETH','TRX']);

  const tokens = normalizeToArray_(data['總開了哪些幣別'])
    .map(s => String(s || '').trim().toUpperCase())
    .map(s => ALIAS_MAP[s] || s)
    .map(t => t.replace(/[^A-Z]/g, ''))
    .filter(t => t.length >= 2);

  const allSet  = new Set(tokens);
  const fiatSet = new Set([...allSet].filter(t => !CRYPTO_SET.has(t)));

  if (fiatSet.size === 0) return 'CY';
  if (fiatSet.size === 1) return [...fiatSet][0].slice(0, 2);
  return 'MA';
}

function findNextNumberForPrefix_(prefix) {
  try {
    const ss = SpreadsheetApp.openByUrl(THIRD_BOOK_URL);
    const sheet = ss.getSheetByName(STATION_SHEET_NAME);
    const vals = sheet.getRange(1, STATION_COL_INDEX, sheet.getLastRow(), 1).getValues();
    const re = new RegExp('^' + prefix + '(\\d{4})$');
    let max = 0;
    vals.forEach(r => {
      const m = String(r[0] || '').match(re);
      if (m) max = Math.max(max, parseInt(m[1], 10));
    });
    return max + 1;
  } catch (err) {
    Logger.log(err);
    return 1;
  }
}

/***** ====== MAIN ENTRY ====== *****/
function submitData(data) {
  if (!data || Object.keys(data).length === 0) throw new Error("提交資料為空！");

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const merchantName = String(data["名称"] || "未知名称").trim();

  // Duplicate guard: abort if tab already exists (prevents F5 resubmission)
  const isNewAcct = (data['accountType'] !== 'old_add_currency');
  if (isNewAcct && merchantName && merchantName !== "未知名称" && ss.getSheetByName(merchantName)) {
    Logger.log(`[Duplicate blocked] ${merchantName} already exists`);
    return "此商戶已完成開戶，請勿重複發送。";
  }

  if (isNewAcct) {
    try { data['分站号'] = generateStationId_(data); }
    catch (e) { data['分站号'] = 'CY0001'; }
  } else {
    if (!data['分站号'] || !String(data['分站号']).trim()) {
      throw new Error("舊商戶新增新幣別必須填寫『分站号 (Station ID)』！");
    }
    data['分站号'] = String(data['分站号']).trim().toUpperCase();
  }

  let sheet = ss.getSheetByName(merchantName);
  if (sheet) ss.deleteSheet(sheet);
  sheet = ss.insertSheet(merchantName);

  // ====== Write merchant tab ======
  let rowIndex = 1;
  const labelMap = {
    "分站号": "Station ID", "名称": "Name", "联络人": "Contact Person",
    "连络电话": "Phone Number", "Email": "Email",
    "开通业务_上分": "Using Services_Pay", "开通业务_回分": "Using Services_Payout",
    "總開了哪些幣別": "Which currencies are currently active",
    "回分币别[]": "Payout Currency", "回分费率[]": "Payout Rate",
    "回分服务费元/笔[]": "Payout Charge (Dollar per Transaction)",
    "回分代理费率[]": "Service Rate of Agent", "回分支付大项[]": "Payment Type",
    "提现币别[]": "Withdraw Currency", "提现费率[]": "Withdraw Rate",
    "提现手续费[]": "Withdraw Charge", "提现区间下限[]": "Withdraw Range Min (Dollar)",
    "提现区间上限[]": "Withdraw Range Max (Dollar)",
    "上分币别[]": "Pay Currency", "上分支付大项[]": "Payment Type",
    "服务费率[]": "Service Rate (%)", "服务费点数元[]": "Service Fee Scores-Dollar per Transaction",
    "上分代理费率[]": "Service Rate of Agent (%)",
    "所属代理": "Affiliated agent", "填写代理名字": "Enter the agent name",
    "访问API": "API IP white list", "后台登入": "Login IP white list",
    "下发IP": "Withdraw IP white list"
  };

  // Block 1: Basic info
  (function writeBasicInfo_() {
    const fields = ["分站号","名称","联络人","连络电话","Email"];
    let r = 1;
    fields.forEach(f => {
      sheet.getRange(r, 1).setValue(`${f} (${labelMap[f] || ''})`);
      const val = data[f] == null ? "" : String(data[f]);
      if (f === "连络电话") {
        const cell = sheet.getRange(r, 2);
        cell.setNumberFormat('@');
        cell.setValue(val ? ("'" + val) : "");
      } else {
        sheet.getRange(r, 2).setValue(val);
      }
      r++;
    });
    rowIndex = r + 1;
  })();

  // Block 2: Services + currencies
  ["开通业务_上分","开通业务_回分","總開了哪些幣別"].forEach(f => {
    sheet.getRange(rowIndex, 1).setValue(`${f} (${labelMap[f] || ''})`);
    normalizeToArray_(data[f]).forEach((v,i) => sheet.getRange(rowIndex, 2+i).setValue(v));
    rowIndex++;
  });
  rowIndex++;

  // Block 3–5: Payout / Withdraw / Payin (multi-column per currency)
  [
    ["回分币别[]","回分费率[]","回分服务费元/笔[]","回分代理费率[]","回分支付大项[]"],
    ["提现币别[]","提现费率[]","提现手续费[]","提现区间下限[]","提现区间上限[]"],
    ["上分币别[]","上分支付大项[]","服务费率[]","服务费点数元[]","上分代理费率[]"]
  ].forEach(fields => {
    fields.forEach(f => {
      sheet.getRange(rowIndex, 1).setValue(`${f.replace("[]","")} (${labelMap[f] || ''})`);
      normalizeToArray_(data[f]).forEach((v,i) => sheet.getRange(rowIndex, 2+i).setValue(v));
      rowIndex++;
    });
    rowIndex++;
  });

  // Block 6: Agent + IP whitelists
  ["所属代理","填写代理名字","访问API","后台登入","下发IP"].forEach(f => {
    sheet.getRange(rowIndex, 1).setValue(`${f} (${labelMap[f] || ''})`);
    sheet.getRange(rowIndex, 2).setValue(normalizeToArray_(data[f]).join("\n"));
    rowIndex++;
  });

  // Sync to 商戶＆代理 management sheet
  writeToStationSheet_(data);

  // Telegram notification to CS group
  const alertNotice = "\n\n⚠️ Reminder: This is production account info. If the merchant needs to test, please ensure a test account is provided.";
  const msg = isNewAcct
    ? `${data['分站号']} - ${merchantName} is now ready to create account, please proceed via HTpay_Selfservice bot.${alertNotice}`
    : `${data['分站号']} - ${merchantName} (Existing Merchant - New Currency) is now ready, please proceed via HTpay_Selfservice bot.${alertNotice}`;
  notifyGroups_(msg);
}

/***** ====== WRITE TO 商戶＆代理 SHEET ====== *****/
function writeToStationSheet_(data) {
  try {
    const ss = SpreadsheetApp.openByUrl(THIRD_BOOK_URL);
    const sheet = ss.getSheetByName(STATION_SHEET_NAME);
    const selectedCurrencies = normalizeToArray_(data['總開了哪些幣別']);
    const currenciesToWrite = selectedCurrencies.length ? selectedCurrencies : [''];

    const stationId  = String(data['分站号'] || '');
    const name       = String(data['名称']   || '');
    const agentName  = String(data['填写代理名字'] || data['所属代理'] || '').trim();
    const todayStr   = Utilities.formatDate(new Date(), TIMEZONE, 'yyyy/MM/dd');
    const topupOn    = normalizeToArray_(data['上分支付大项[]']).length > 0 || !!data['开通业务_上分'];
    const payoutOn   = normalizeToArray_(data['回分支付大项[]']).length > 0 || !!data['开通业务_回分'];

    // Rates stored as decimals (e.g. 3‰ → 0.003) for direct use in Sheets formulas
    const agentRatePct  = (() => { const a = normalizeToArray_(data['上分代理费率[]']); const b = normalizeToArray_(data['回分代理费率[]']); const r = a.length ? a[0] : (b.length ? b[0] : ''); const n = cleanNumber_(r); return n === '' ? '' : n / 10; })();
    const agentRateDec  = agentRatePct  === '' ? '' : agentRatePct  / 100;
    const topupRateDec  = (() => { if (!topupOn) return ''; const arr = normalizeToArray_(data['服务费率[]']); if (!arr.length) return ''; const n = cleanNumber_(arr[0]); return n === '' ? '' : n / 100; })();
    const payoutRateDec = (() => { if (!payoutOn) return ''; const arr = normalizeToArray_(data['回分费率[]']); if (!arr.length) return ''; const n = cleanNumber_(arr[0]); return n === '' ? '' : n / 100; })();

    currenciesToWrite.forEach(currRaw => {
      const row = findFirstEmptyABRow_(sheet);
      sheet.getRange(row, 1).setValue(stationId);
      sheet.getRange(row, 2).setValue(name);
      sheet.getRange(row, 4).setValue(prettyCurrencyLabel_(currRaw));
      sheet.getRange(row, 6).setValue(todayStr);
      sheet.getRange(row, 8).setValue(agentName);
      sheet.getRange(row, 9).setValue(agentRateDec).setNumberFormat('0.00%');
      sheet.getRange(row,14).setValue(topupRateDec).setNumberFormat('0.00%');
      sheet.getRange(row,15).setValue(topupOn  ? 'ON' : 'OFF');
      sheet.getRange(row,17).setValue(payoutRateDec).setNumberFormat('0.00%');
      sheet.getRange(row,18).setValue(payoutOn ? 'ON' : 'OFF');
      sheet.getRange(row,21).setValue('On');
      sheet.getRange(row,22).setValue('待開戶');
    });
  } catch (err) {
    Logger.log('[ERROR] writeToStationSheet_: ' + err);
  }
}

/***** ====== CURRENCY DROPDOWN (called from frontend) ====== *****/
function getUniqueCurrencies() {
  const ss = SpreadsheetApp.openByUrl(THIRD_BOOK_URL);
  const sheet = ss.getSheetByName(STATION_SHEET_NAME);
  const values = sheet.getRange("D3:D" + sheet.getLastRow()).getValues().flat();

  const excludeSet = new Set(['大額充U','小額充U','BTC,CRYPTO,ETH,INR,USDT','BRL','ALL','-'].map(s => s.toLowerCase()));
  const seen = new Set();
  const result = [];

  values.forEach(v => {
    if (!v) return;
    String(v).split(',').forEach(s => {
      const raw = s.trim();
      if (!raw || raw.toLowerCase().includes('(in progress)') || excludeSet.has(raw.toLowerCase())) return;
      const canon = raw.toUpperCase();
      if (!seen.has(canon)) { seen.add(canon); result.push(canon); }
    });
  });

  result.sort();
  return result;
}
