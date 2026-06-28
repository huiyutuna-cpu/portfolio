/**
 * Case 05 — Telegram Bot: Account Opening Flow
 *
 * CS uses /openacct command (or menu button) to walk through merchant account creation.
 * Bot reads from the Account Book spreadsheet (one tab per merchant).
 * Auto-detects new vs. existing merchant from sheet data.
 * Guides through fixed 4 steps + dynamic checklist + final merchant notification.
 */

const ACCOUNT_BOOK_ID = 'CONFIG_ACCOUNT_BOOK_ID';
const ACCOUNT_NOTIFY_CHATS = ['CONFIG_NOTIFY_CHAT_ID_1', 'CONFIG_NOTIFY_CHAT_ID_2'];

// ==================== 商戶&代理 open-date stamp ====================
/**
 * After account creation, stamp today's date into column V of all matching rows
 * (matched by Station ID in col A AND merchant name in col B).
 */
function updateMerchantOpenDateInThirdSheetAll_(subId, subName) {
  try {
    if (!subId || !subName) return 0;
    const ss = SpreadsheetApp.openById(THIRD_BOOK_ID); // CONFIG_THIRD_BOOK_ID
    const sh = ss.getSheetByName(SHEET_MERCHANT);
    if (!sh) return 0;

    const last = sh.getLastRow();
    const ab = sh.getRange(1, 1, last, 2).getDisplayValues();
    const Akey = toHalfWidth_(String(subId).trim());
    const Bkey = toHalfWidth_(String(subName).trim());

    const hitRows = [];
    for (let i = 0; i < last; i++) {
      if (toHalfWidth_(String(ab[i][0]||'').trim()) === Akey &&
          toHalfWidth_(String(ab[i][1]||'').trim()) === Bkey) {
        hitRows.push(i + 1);
      }
    }
    if (!hitRows.length) return 0;

    const today = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy/MM/dd');
    sh.getRangeList(hitRows.map(r => 'V' + r)).setValue(today);
    return hitRows.length;
  } catch (e) {
    Logger.log('[updateMerchantOpenDateInThirdSheetAll_] ' + e);
    return 0;
  }
}

// ==================== State machine ====================
function startAccountOpenFlow_(userId) {
  PropertiesService.getUserProperties().setProperty(
    'state:' + userId,
    JSON.stringify({ flow: 'acct_open', step: 'ask_name', data: {} })
  );
}
function clearAccountOpenState_(userId) {
  PropertiesService.getUserProperties().deleteProperty('state:' + userId);
}
function getAccountOpenState_(userId) {
  const raw = PropertiesService.getUserProperties().getProperty('state:' + userId);
  try { return raw ? JSON.parse(raw) : null; } catch(_) { return null; }
}
function setAccountOpenState_(userId, st) {
  PropertiesService.getUserProperties().setProperty('state:' + userId, JSON.stringify(st));
}

// ==================== Sheet lookup ====================
function findAccountSheetsByName_(keyword) {
  const ss = SpreadsheetApp.openById(ACCOUNT_BOOK_ID);
  const kw = toHalfWidth_(String(keyword||'').trim()).toLowerCase();
  if (!kw) return [];
  return ss.getSheets()
    .filter(sh => toHalfWidth_(sh.getName().trim()).toLowerCase() === kw)
    .map(sh => ({ sheetName: sh.getName(), gid: sh.getSheetId() }));
}

// ==================== Checklist builder ====================
/**
 * Reads the merchant account sheet and produces an array of checklist strings.
 * Sections: basic info (A1-A9) → payout per currency → withdraw per currency
 *           → payin per column → misc (A29+)
 */
function buildChecklistFromOpenSheet_(sheetName) {
  const ss = SpreadsheetApp.openById(ACCOUNT_BOOK_ID);
  const sh = ss.getSheetByName(sheetName);
  if (!sh) throw new Error('Sheet not found: ' + sheetName);

  const lastRow = Math.max(sh.getLastRow(), 40);
  const lastCol = Math.max(sh.getLastColumn(), 7);
  const labelsA = sh.getRange(1, 1, lastRow, 1).getDisplayValues().map(r => String(r[0]||'').trim());

  const findRow_ = regex => { for (let i=0;i<labelsA.length;i++) if (regex.test(labelsA[i])) return i; return -1; };
  const items = [];

  // Basic info: rows 1-8 (col B)
  const bVals = sh.getRange(1, 2, Math.min(8, lastRow), 1).getDisplayValues();
  for (let r=1; r<=Math.min(8,lastRow); r++) {
    const lab = labelsA[r-1]; if (!lab) continue;
    items.push(`${lab}: ${(bVals[r-1]&&bVals[r-1][0]) ? String(bVals[r-1][0]).trim() : '（未填）'}`);
  }
  // Row 9: merge B:G
  if (lastRow >= 9 && labelsA[8]) {
    const row9 = sh.getRange(9, 2, 1, Math.max(lastCol-1,1)).getDisplayValues()[0];
    items.push(`${labelsA[8]}: ${row9.map(v=>String(v||'').trim()).filter(Boolean).join(', ') || '（未填）'}`);
  }

  // Payout & Withdraw: per currency column
  const poCurR  = findRow_(/回分.*幣別|Payout.*Currency/i);
  const poRateR = findRow_(/回分.*費率|Payout.*Rate/i);
  const poSrvR  = findRow_(/回分服務費.*筆|Service fee.*per/i);
  const poAgR   = findRow_(/回分代理費率|Agent.*Rate/i);
  const poPmR   = findRow_(/回分支付大項|Payout.*Methods/i);
  const wdCurR  = findRow_(/提現.*幣別|Withdraw.*Currency/i);
  const wdRateR = findRow_(/提現.*費率|Withdraw.*Rate/i);
  const wdFeeR  = findRow_(/提現手續費|Withdraw.*Fee/i);
  const wdMinR  = findRow_(/提現區間下限/i);
  const wdMaxR  = findRow_(/提現區間上限/i);

  for (let c=2; c<=lastCol; c++) {
    const col = sh.getRange(1, c, lastRow, 1).getDisplayValues().map(r=>String(r[0]||'').trim());
    if (poCurR >= 0 && col[poCurR]) {
      items.push([
        `${labelsA[poCurR]||'回分幣別'}: ${col[poCurR]}`,
        `${labelsA[poRateR]||'回分費率'}: ${col[poRateR]||'（未填）'}`,
        `${labelsA[poSrvR]||'回分服務費'}: ${col[poSrvR]||'（未填）'}`,
        `${labelsA[poAgR]||'回分代理費率'}: ${col[poAgR]||'（未填）'}`,
        `${labelsA[poPmR]||'回分支付大項'}: ${col[poPmR]||'（未填）'}`
      ].join('\n'));
    }
    if (wdCurR >= 0 && col[wdCurR]) {
      items.push([
        `${labelsA[wdCurR]||'提現幣別'}: ${col[wdCurR]}`,
        `${labelsA[wdRateR]||'提現費率'}: ${col[wdRateR]||'（未填）'}`,
        `${labelsA[wdFeeR]||'提現手續費'}: ${col[wdFeeR]||'（未填）'}`,
        `${labelsA[wdMinR]||'下限'}: ${col[wdMinR]||'（未填）'}`,
        `${labelsA[wdMaxR]||'上限'}: ${col[wdMaxR]||'（未填）'}`
      ].join('\n'));
    }
  }

  // Payin: per column
  const tpCurR  = findRow_(/上分.*幣別|Top.?up.*Currency/i);
  const tpPayR  = findRow_(/上分支付大項|Top.?up.*Methods/i);
  const tpRateR = findRow_(/服務費率.*上分|Service.*Rate.*Top/i);
  const tpAmtR  = findRow_(/服務費.*金額|Amount.*Top/i);
  const tpAgR   = findRow_(/上分代理費率|Agent.*Rate.*Top/i);

  if (tpPayR >= 0) {
    const payNames = sh.getRange(tpPayR+1, 2, 1, lastCol-1).getDisplayValues()[0].map(v=>String(v||'').trim());
    payNames.forEach((pay, c) => {
      if (!pay) return;
      const ci = 2 + c;
      items.push([
        `${labelsA[tpCurR]||'上分幣別'}: ${sh.getRange(tpCurR+1,ci).getDisplayValue()}`,
        `${labelsA[tpPayR]||'上分支付大項'}: ${pay}`,
        `${labelsA[tpRateR]||'服務費率'}: ${sh.getRange(tpRateR+1,ci).getDisplayValue()||'（未填）'}`,
        `${labelsA[tpAmtR]||'服務費金額'}: ${sh.getRange(tpAmtR+1,ci).getDisplayValue()||'（未填）'}`,
        `${labelsA[tpAgR]||'代理費率'}: ${sh.getRange(tpAgR+1,ci).getDisplayValue()||'（未填）'}`
      ].join('\n'));
    });
  }

  // Misc rows A29+
  if (lastRow >= 29) {
    for (let r=29; r<=lastRow; r++) {
      const lab = labelsA[r-1]; if (!lab) continue;
      for (let c=2; c<=Math.min(lastCol,5); c++) {
        const val = sh.getRange(r,c).getDisplayValue().trim();
        if (val) items.push(`${lab}: ${val}`);
      }
    }
  }

  return items.filter(Boolean);
}

// ==================== Checklist stepper ====================
function sendChecklistItem_(chatId, msgId, st) {
  const items = st.data.items || [];
  const idx   = st.data.idx || 0;
  const total = items.length;

  if (idx >= total) {
    const tabName  = st.data.sheetName || '';
    const merId    = st.data.merchantId || '';
    const merName  = st.data.merchantName || tabName;
    const head     = merId ? `${merId} - ${tabName}` : tabName;

    const finalMsg = st.data.isOldMerchant
      ? `${head} Merchant account (Existing Merchant - New Currency) is now created, please check for assignment and routing.`
      : `${head} Merchant account is now created, please check for assignment and routing.`;

    // Notify CS groups
    ACCOUNT_NOTIFY_CHATS.forEach(gid => tgSendMessage(gid, finalMsg));

    // Stamp open date in 商戶&代理
    let stampedCount = 0;
    try { stampedCount = updateMerchantOpenDateInThirdSheetAll_(merId, merName); } catch(_) {}
    const suffix = stampedCount > 0
      ? `\n（已於《商戶＆代理》V 欄標記今日，共 ${stampedCount} 列）`
      : `\n（提醒：未能在《商戶＆代理》標記日期，請人工確認）`;

    st.step = 'notify_lang';
    setAccountOpenState_(st.userId, st);

    const prompt = '✅ All items completed.\n' + finalMsg + suffix + '\n\n請選擇要發給商戶的通知語言：';
    const kb = { inline_keyboard: [[
      { text: 'ZH-CN', callback_data: 'open:lang:zh' },
      { text: 'EN',    callback_data: 'open:lang:en' }
    ],[
      { text: '↩︎ Back to Menu', callback_data: 'menu:root' }
    ]]};

    if (msgId) tgEditMessageText(chatId, msgId, prompt, { reply_markup: JSON.stringify(kb) });
    else       tgSendMessage(chatId, prompt, { reply_markup: JSON.stringify(kb) });
    return;
  }

  const text = `Step ${idx+1}/${total}\n\n請確認並完成以下設定，完成後點下一項：\n\n${items[idx]}`;
  const rows = [];
  if (idx > 0) rows.push([{ text: '⬅︎ 上一項', callback_data: 'open:prev' }]);
  rows.push([{ text: '✅ 已完成，下一項', callback_data: 'open:next' }]);
  rows.push([{ text: '↩︎ 取消', callback_data: 'open:cancel' }]);
  const kb = { inline_keyboard: rows };

  if (msgId) {
    tgEditMessageText(chatId, msgId, text, { reply_markup: JSON.stringify(kb) });
  } else {
    const res = tgSendMessage(chatId, text, { reply_markup: JSON.stringify(kb) });
    try {
      const mid = res && res.result && res.result.message_id;
      if (mid) { st.data.lastMsgId = mid; setAccountOpenState_(st.userId, st); }
    } catch(_) {}
  }
}

// ==================== Navigation handlers ====================
function proceedAccountOpenNext_(userId, chatId, msgId) {
  const st = getAccountOpenState_(userId);
  if (!st || st.flow !== 'acct_open') {
    tgEditMessageText(chatId, msgId, 'Session expired. Please /menu → 🧾 Account Opening', {}); return;
  }

  const curIdx = st.data.idx || 0;

  // Step 4 (index 3) = API key step
  // New merchant: require 64-char key; existing merchant: auto-skip
  if (curIdx === 3) {
    if (st.data.isOldMerchant) {
      st.data.apiKey = "舊商戶 無需填寫";
    } else if (!st.data.apiKey) {
      st.step = 'await_api_key';
      setAccountOpenState_(userId, st);
      const ask = '請貼上剛產生的 API key（僅此一次會保存）：';
      const kb = { inline_keyboard: [[{ text: '↩︎ 取消', callback_data: 'open:cancel' }]] };
      if (msgId) tgEditMessageText(chatId, msgId, ask, { reply_markup: JSON.stringify(kb) });
      else       tgSendMessage(chatId, ask, { reply_markup: JSON.stringify(kb) });
      return;
    }
  }

  st.step = 'confirming';
  st.data.idx = curIdx + 1;
  setAccountOpenState_(userId, st);
  sendChecklistItem_(chatId, msgId, st);
}

function proceedAccountOpenPrev_(userId, chatId, msgId) {
  const st = getAccountOpenState_(userId);
  if (!st || st.flow !== 'acct_open') {
    tgEditMessageText(chatId, msgId, 'Session expired. Please /menu → 🧾 Account Opening', {}); return;
  }
  const cur = st.data.idx || 0;
  st.data.idx = cur > 0 ? cur - 1 : 0;
  setAccountOpenState_(userId, st);
  sendChecklistItem_(chatId, msgId, st);
}

// ==================== Sheet selection → checklist init ====================
/**
 * confirmAccountOpenPick_: reads selected sheet, auto-detects merchant type,
 * prepends 4 fixed steps, then starts checklist.
 */
function confirmAccountOpenPick_(userId, chatId, msgId, pickIdx) {
  const st = getAccountOpenState_(userId);
  if (!st || st.flow !== 'acct_open' || st.step !== 'choose_tab') {
    (msgId ? tgEditMessageText : tgSendMessage)(chatId, msgId||chatId, 'Session expired. Please /menu → 🧾 Account Opening', {});
    return;
  }

  const sel = (st.data.matches || [])[pickIdx];
  if (!sel) {
    st.step = 'ask_name'; st.data = {}; setAccountOpenState_(userId, st);
    (msgId ? tgEditMessageText(chatId, msgId, 'Invalid selection. Please enter merchant name again:', {})
           : tgSendMessage(chatId, 'Invalid selection. Please enter merchant name again:'));
    return;
  }

  let merchantId = '', merchantName = sel.sheetName, isOldMerchant = false;
  try {
    const s2 = SpreadsheetApp.openById(ACCOUNT_BOOK_ID).getSheetByName(sel.sheetName);
    merchantId   = (s2 && s2.getRange('B1').getDisplayValue()) || '';
    merchantName = (s2 && s2.getRange('B2').getDisplayValue()) || merchantName;
    // Detect existing merchant: frontend injects "舊商戶無須填寫" into B3 for old merchants
    const contact = (s2 && s2.getRange('B3').getDisplayValue()) || '';
    isOldMerchant = contact.indexOf("舊商戶無須填寫") !== -1;
  } catch(_) {}

  let items = [];
  try { items = buildChecklistFromOpenSheet_(sel.sheetName); }
  catch(e) {
    Logger.log('[buildChecklist] ' + e);
    (msgId ? tgEditMessageText(chatId, msgId, 'Failed to read the sheet. Please try another name.', {})
           : tgSendMessage(chatId, 'Failed to read the sheet. Please try another name.'));
    st.step = 'ask_name'; st.data = {}; setAccountOpenState_(userId, st);
    return;
  }

  // Fixed 4 steps vary by merchant type
  const pwdHint    = isOldMerchant ? "舊商戶 無需填寫" : "12345";
  const apiKeyHint = isOldMerchant ? "舊商戶 無需填寫" : "（請產生後貼上）";
  const fixedSteps = [
    `登入帐号 (Login Account): ${merchantName}`,
    `登入密码 (Password): ${pwdHint}`,
    `验证密码 (Confirmed password): ${pwdHint}`,
    `API 密钥 (API key): ${apiKeyHint}`
  ];
  items = fixedSteps.concat(items);

  st.step = 'confirming';
  st.userId = userId;
  st.data = { sheetName: sel.sheetName, items, idx: 0, merchantId, merchantName, apiKey: '', isOldMerchant };
  setAccountOpenState_(userId, st);
  sendChecklistItem_(chatId, null, st);
}

// ==================== Text input router ====================
function handleAccountOpenConversation_(m, userId, chatId, text) {
  const st = getAccountOpenState_(userId);
  if (!st || st.flow !== 'acct_open') return false;

  if (st.step === 'ask_name') {
    const name = String(text||'').trim();
    if (!name) { tgSendMessage(chatId, 'Please enter merchant name:'); return true; }

    const matches = findAccountSheetsByName_(name);
    if (!matches.length) { tgSendMessage(chatId, 'No sheet matched. Please enter another keyword:'); return true; }

    st.step = 'choose_tab';
    st.data = { matches };
    setAccountOpenState_(userId, st);

    if (matches.length === 1) { confirmAccountOpenPick_(userId, chatId, null, 0); return true; }

    const kb = { inline_keyboard: [] };
    matches.forEach((it, i) => kb.inline_keyboard.push([{ text: it.sheetName, callback_data: 'open:pick:' + i }]));
    kb.inline_keyboard.push([{ text: '↩︎ 取消', callback_data: 'open:cancel' }]);
    tgSendMessage(chatId, 'Multiple sheets matched, please pick one:', { reply_markup: JSON.stringify(kb) });
    return true;
  }

  // API key input: validate exactly 64 characters
  if (st.step === 'await_api_key') {
    const key = String(text||'').trim();
    if (!key) { tgSendMessage(chatId, 'API key 不能為空，請重新貼上（需 64 字元）：'); return true; }
    if (key.length !== 64) {
      tgSendMessage(chatId, `API key 長度不符（目前 ${key.length} 字元），請重新貼上正確的 64 字元 key：`);
      return true;
    }
    st.data.apiKey = key;
    st.step = 'confirming';
    st.data.idx = (st.data.idx || 0) + 1;
    setAccountOpenState_(userId, st);
    sendChecklistItem_(chatId, null, st);
    return true;
  }

  return true;
}

// ==================== Final merchant notification ====================
/**
 * Generates bilingual (ZH-CN / EN) merchant onboarding notice.
 * Content differs for new merchants (includes password + API key)
 * vs. existing merchants (no credential change).
 */
function finalizeWithMerchantNotice_(userId, chatId, msgId, lang) {
  const st = getAccountOpenState_(userId);
  if (!st || st.flow !== 'acct_open' || st.step !== 'notify_lang') {
    tgEditMessageText(chatId, msgId, 'Session expired. Please /menu → 🧾 Account Opening', {}); return;
  }

  const merId   = st.data.merchantId || '';
  const merName = st.data.merchantName || (st.data.sheetName || '');
  const apiKey  = st.data.isOldMerchant ? "（延用原本密鑰，無須更換）" : (st.data.apiKey || '(未記錄)');
  const url     = `https://CONFIG_BACKEND_DOMAIN/${merId || ''}`;

  let note = '';
  if (lang === 'zh') {
    note = st.data.isOldMerchant
      ? `商户名：${merName}\n已成功添加新币别！\n如果有任何需要协助的都欢迎跟我们说。\n原本的分站号及金钥維持不變，即可直接开始使用新幣別。\n\n商户后台：\n${url}\n\n分站号 (opmhtid)：\n${merId}\n\n分站名称：\n${merName}\n\nAPI 密钥：\n${apiKey}`
      : `商户名：${merName}\n已开户完成\n如果有任何需要协助的都欢迎跟我们说。\n只需更换测试账号的分站号及金钥，即可开始使用。\n\n商户后台：\n${url}\n\n分站号 (opmhtid)：\n${merId}\n\n分站名称：\n${merName}\n\n密码：\n12345\n\n<b>登入后请马上修改密码</b>\n\nAPI 密钥：\n${apiKey}`;
  } else {
    note = st.data.isOldMerchant
      ? `Merchant ID: ${merName}\nNew Currency has been successfully added!\nIf there is any assistance you need, please feel free to talk to us!\nYour existing Station ID and API key remain unchanged, you can start using the new currency immediately.\n\nMerchant Backend URL：\n${url}\n\nID (opmhtid)：\n${merId}\n\nMerchant Name：\n${merName}\n\nAPI key：\n${apiKey}`
      : `Merchant ID: ${merName}\nAccount is now created.\nIf there is any assistance you need, please feel free to talk to us!\nYou only need to replace the station ID and API key of the test account to start using.\n\nMerchant Backend URL：\n${url}\n\nStation ID (opmhtid)：\n${merId}\n\nMerchant Name：\n${merName}\n\nPassword：\n12345\n\n<b>Please change your password immediately after logging in.</b>\n\nAPI key：\n${apiKey}`;
  }

  const payload = {
    reply_markup: JSON.stringify({ inline_keyboard: [[{ text: '↩︎ Back to Menu', callback_data: 'menu:root' }]] }),
    parse_mode: 'HTML'
  };
  if (msgId) tgEditMessageText(chatId, msgId, note, payload);
  else       tgSendMessage(chatId, note, payload);

  clearAccountOpenState_(userId);
}
