/**
 * Case 07 — Crypto Wallet Transaction Auto-Matching System
 *
 * For each unprocessed row in a Wallet Data sheet, extracts a 64-char TxHash
 * or 36-char UUID from configured target columns and looks it up across multiple
 * platform export sheets (BestPay / HTPay).
 * INR OTC fallback: matches by Amount + Date when no standard hash is present.
 *
 * Trigger: onEdit (auto, 10-row threshold) or manualForceRun() (button).
 * Unmatched rows → "未配到" (red) + email alert.
 */

const CONFIG = [
  {
    targetSheetName: "Bestpayops Wallet Data",
    targetHashCols: ["O", "R"],
    sources: [
      { name: "bestpayops Data",       hashCols: ["K"], valCols: ["E", "C"], type: "EC"     },
      { name: "HTPay報表_USDT充值",     hashCols: ["M", "N"], valCols: ["A"], type: "SINGLE" }
    ]
  },
  {
    targetSheetName: "Bestpayotc Wallet Data",
    targetHashCols: ["O"],
    sources: [
      { name: "bestpayotc Data",        hashCols: ["K"], valCols: ["E", "C"], type: "EC"     },
      { name: "HTPay報表_USDT下發",      hashCols: ["M", "N"], valCols: ["C"], type: "SINGLE" }
    ]
  },
  {
    targetSheetName: "bestpayinrotc Wallet Data",
    targetHashCols: ["O", "R"],
    sources: [
      { name: "bestpayinrotc Data",     hashCols: ["K"], valCols: ["E", "C"], type: "EC"     },
      { name: "bestpayinrotc TXID",     hashCols: ["I"], valCols: ["A"],      type: "SINGLE" }
    ]
  }
];

const RECEIVER_EMAIL = "CONFIG_RECEIVER_EMAIL";

// ==================== Triggers ====================

function onEdit(e) {
  if (!e || !e.range) return;
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = e.range.getSheet();
  const conf  = CONFIG.find(c => c.targetSheetName === sheet.getName());
  if (!conf) return;

  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return;

  // Auto-run only when 10+ consecutive rows have col B filled but col A empty
  const rangeAB    = sheet.getRange(1, 1, lastRow, 2).getValues();
  let emptyACount  = 0;
  let shouldRun    = false;
  for (let i = 1; i < rangeAB.length; i++) {
    if (rangeAB[i][1] !== "" && rangeAB[i][0] === "") {
      emptyACount++;
    } else {
      emptyACount = 0;
    }
    if (emptyACount >= 10) { shouldRun = true; break; }
  }

  if (shouldRun) {
    ss.toast("正在執行多重條件比對 (金額/日期/字串)...", "⚙️ 運行中", -1);
    try {
      executeMappingEngine(false);
      ss.toast("處理完畢！格式已統一為 Arial 12。", "✅ 完成", 5);
    } catch (err) {
      ss.toast("錯誤：" + err.message, "❌ 錯誤", 10);
    }
  }
}

function manualForceRun() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  ss.toast("全面校對模式啟動...", "🚀 執行中", -1);
  try {
    executeMappingEngine(true);
    ss.toast("校對完成！", "✅ 完成", 5);
  } catch (err) {
    ss.toast("校對錯誤：" + err.message, "❌ 錯誤", 10);
  }
}

// ==================== Core engine ====================

function executeMappingEngine(isForceAll) {
  const ss        = SpreadsheetApp.getActiveSpreadsheet();
  const errorLogs = [];

  // Load all source sheets into memory once — avoids repeated API reads in the inner loop
  const sourceCaches = {};
  CONFIG.forEach(conf => {
    conf.sources.forEach(src => {
      if (!sourceCaches[src.name]) {
        const sSheet = ss.getSheetByName(src.name);
        sourceCaches[src.name] = sSheet ? sSheet.getDataRange().getValues() : null;
      }
    });
  });

  CONFIG.forEach(conf => {
    const targetSheet = ss.getSheetByName(conf.targetSheetName);
    if (!targetSheet) return;

    const lastRow   = targetSheet.getLastRow();
    const targetData = targetSheet.getDataRange().getValues();

    for (let i = 1; i < lastRow; i++) {
      const rowNum = i + 1;
      const colA   = targetData[i][colToIdx("A")];
      const colB   = targetData[i][colToIdx("B")];
      const colP   = targetData[i][colToIdx("P")];

      if (isForceAll) {
        if (colB === "") continue;
      } else {
        if (colA !== "" || colB === "") continue;  // skip already-matched rows
      }

      const cellA = targetSheet.getRange(rowNum, colToIdx("A") + 1);
      cellA.setFontFamily("Arial").setFontSize(12);

      if (String(colP).trim().toLowerCase() === "cancel") {
        cellA.setValue("Cancel").setBackground("#f3f3f3").setFontColor("black");
        continue;
      }

      let foundResult = null;

      // --- Strategy 1: INR OTC — amount + date fallback ---
      if (conf.targetSheetName === "bestpayinrotc Wallet Data") {
        const colK   = String(targetData[i][colToIdx("K")]).trim();
        const colM   = String(targetData[i][colToIdx("M")]).trim();
        const colG   = targetData[i][colToIdx("G")];   // amount
        const colS   = targetData[i][colToIdx("S")];   // date

        if (colK === "IN" && colM === "CONFIG_MERCHANT_NAME") {
          const txidData = sourceCaches["bestpayinrotc TXID"];
          if (txidData) {
            for (let k = 0; k < txidData.length; k++) {
              const amtMatch  = Number(colG) === Number(txidData[k][colToIdx("H")]);
              const dateMatch = formatDate(colS) === formatDate(txidData[k][colToIdx("B")]);
              if (amtMatch && dateMatch && colS !== "" && txidData[k][colToIdx("B")] !== "") {
                foundResult = txidData[k][colToIdx("A")];
                break;
              }
            }
          }
        }
      }

      // --- Strategy 2: Hash / UUID extraction & lookup ---
      if (!foundResult) {
        const targetHashes = [];
        conf.targetHashCols.forEach(letter => {
          const raw   = targetData[i][colToIdx(letter)];
          const match = String(raw).match(/[0-9a-f]{64}|[0-9a-f-]{36}/i);
          if (match) targetHashes.push(match[0].toLowerCase().replace(/-/g, ""));
        });

        if (targetHashes.length > 0) {
          outer:
          for (const src of conf.sources) {
            const sData = sourceCaches[src.name];
            if (!sData) continue;
            for (let j = 0; j < sData.length; j++) {
              for (const hLetter of src.hashCols) {
                const sRaw   = sData[j][colToIdx(hLetter)];
                const sMatch = String(sRaw).match(/[0-9a-f]{64}|[0-9a-f-]{36}/i);
                if (!sMatch) continue;
                const sHash = sMatch[0].toLowerCase().replace(/-/g, "");
                if (!targetHashes.includes(sHash)) continue;

                if (src.type === "EC") {
                  foundResult = `${sData[j][colToIdx("E")]} ${sData[j][colToIdx("C")]}`;
                } else {
                  let val = String(sData[j][colToIdx(src.valCols[0])]);
                  // Append account tag when specific account identifier present in col M
                  if (src.name === "HTPay報表_USDT充值") {
                    if (String(sData[j][colToIdx("M")]).toLowerCase().includes("config_account_id")) {
                      val += " CONFIG_ACCOUNT_ID";
                    }
                  }
                  foundResult = val;
                }
                break outer;
              }
            }
          }
        }
      }

      // Write result
      if (foundResult) {
        cellA.setValue(foundResult).setBackground("#f3f3f3").setFontColor("black");
      } else {
        cellA.setValue("未配到").setBackground("#ffcccc").setFontColor("red");
        errorLogs.push(`表: ${conf.targetSheetName}, 第 ${rowNum} 列`);
      }
    }
  });

  if (errorLogs.length > 0 && !isForceAll) {
    sendErrorEmail(errorLogs);
  }
}

// ==================== Utilities ====================

function formatDate(dateObj) {
  if (!dateObj || isNaN(Date.parse(dateObj))) return "";
  const d = new Date(dateObj);
  return `${d.getFullYear()}-${("0" + (d.getMonth() + 1)).slice(-2)}-${("0" + d.getDate()).slice(-2)}`;
}

function colToIdx(letter) {
  let col = 0;
  for (let i = 0; i < letter.length; i++) {
    col += (letter.charCodeAt(i) - 64) * Math.pow(26, letter.length - i - 1);
  }
  return col - 1;
}

function sendErrorEmail(logs) {
  try {
    MailApp.sendEmail(
      RECEIVER_EMAIL,
      "【系統通知】Hash/多重條件比對失敗提醒",
      "以下資料未匹配成功：\n\n" + logs.join("\n")
    );
  } catch (e) {
    console.log("Email 失敗: " + e.message);
  }
}
