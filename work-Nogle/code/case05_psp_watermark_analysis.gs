/**
 * Case 05 — PSP Monthly Score & Water Level Analysis System
 *
 * Two responsibilities:
 * 1. onOpen: auto-calculate avg scores + grades in YYYY-MM monthly sheets, redraw currency borders
 * 2. runWatermarkAnalysisButton: given a date range, identify core PSPs, fetch volumes
 *    from Drive, and output 4-tier water level thresholds with grade-based multipliers
 */

const WATERMARK_CONFIG = {
  PARENT_FOLDER_ID: "CONFIG_PARENT_FOLDER_ID"
};

// ==================== onOpen: auto-run on spreadsheet open ====================

function onOpen(e) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  ss.getSheets().forEach(sheet => {
    const name = sheet.getName();
    // Only process monthly tabs named YYYY-MM
    if (name.length === 7 && name.includes("-")) {
      calculateScoresAndGrades(sheet);
      autoRenderCurrencyBorders(sheet);
    }
  });
}

// ==================== Module 1: Score + grade calculation ====================

/**
 * Reads daily score columns (E+) and writes:
 * - Col C: average score (to 1 decimal)
 * - Col D: grade string (A+ … D)
 */
function calculateScoresAndGrades(sheet) {
  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();
  if (lastRow < 2) return;

  sheet.getRange("C1").setValue("平均分數").setBackground("#444444").setFontColor("white").setFontWeight("bold").setHorizontalAlignment("center");
  sheet.getRange("D1").setValue("績效評級").setBackground("#444444").setFontColor("white").setFontWeight("bold").setHorizontalAlignment("center");

  if (lastCol < 5) {
    sheet.getRange(2, 3, lastRow - 1, 2).clearContent();
    return;
  }

  const dataRange = sheet.getRange(2, 5, lastRow - 1, lastCol - 4).getValues();
  const cValues = [];
  const dValues = [];

  dataRange.forEach(rowData => {
    let sum = 0, count = 0;
    rowData.forEach(val => {
      if (val !== "" && typeof val === "number" && !isNaN(val)) { sum += val; count++; }
    });

    if (count > 0) {
      const avg = parseFloat((sum / count).toFixed(1));
      cValues.push([avg]);
      let grade = "D";
      if      (avg > 90) grade = "A+";
      else if (avg > 80) grade = "A";
      else if (avg > 70) grade = "B+";
      else if (avg > 60) grade = "B";
      else if (avg > 50) grade = "C+";
      else if (avg > 40) grade = "C";
      else if (avg > 20) grade = "D+";
      dValues.push([grade]);
    } else {
      cValues.push([""]);
      dValues.push([""]);
    }
  });

  sheet.getRange(2, 3, cValues.length, 1).setValues(cValues).setHorizontalAlignment("center").setNumberFormat("0.0");
  sheet.getRange(2, 4, dValues.length, 1).setValues(dValues).setHorizontalAlignment("center").setFontWeight("bold");
}

// ==================== Module 2: Currency border rendering ====================

/**
 * Resets ALL borders to thin grey first (clears stale borders from previous runs),
 * then draws a medium bottom border on each currency group's last row.
 */
function autoRenderCurrencyBorders(sheet) {
  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();
  if (lastRow < 2) return;

  // Full reset — prevents ghost borders when row counts change between months
  sheet.getRange(2, 1, lastRow - 1, lastCol).setBorder(
    true, true, true, true, true, true,
    "#ccc", SpreadsheetApp.BorderStyle.SOLID
  );

  const currencies = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
  for (let i = 0; i < currencies.length; i++) {
    const rowNum = i + 2;
    const isLast = (rowNum === lastRow);
    const isDiff = !isLast && (currencies[i][0] !== currencies[i + 1][0]);
    if (isDiff || isLast) {
      sheet.getRange(rowNum, 1, 1, lastCol).setBorder(
        null, null, true, null, null, null,
        "#444444", SpreadsheetApp.BorderStyle.SOLID_MEDIUM
      );
    }
  }
}

// ==================== Module 3: Main analysis entry (button) ====================

function runWatermarkAnalysisButton() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const dataSheet = ss.getSheetByName("Data");
  if (!dataSheet) {
    SpreadsheetApp.getUi().alert("⚠️ 找不到名為 [ Data ] 的工作表！");
    return;
  }
  executeIndependentWatermarkAnalysis(dataSheet);
}

function executeIndependentWatermarkAnalysis(dataSheet) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const startDateVal = dataSheet.getRange("A2").getValue();
  const endDateVal   = dataSheet.getRange("A3").getValue();
  if (!(startDateVal instanceof Date) || !(endDateVal instanceof Date)) {
    SpreadsheetApp.getUi().alert("⚠️ A2 與 A3 必須填入正確的日期格式！");
    return;
  }

  const startStr = Utilities.formatDate(startDateVal, "GMT+8", "yyyy-MM-dd");
  const endStr   = Utilities.formatDate(endDateVal,   "GMT+8", "yyyy-MM-dd");

  // Collect all months in range
  const targetMonths = [];
  const loopDate = new Date(startDateVal.getTime());
  while (loopDate <= endDateVal) {
    const mStr = Utilities.formatDate(loopDate, "GMT+8", "yyyy-MM-dd").substring(0, 7);
    if (!targetMonths.includes(mStr)) targetMonths.push(mStr);
    loopDate.setDate(loopDate.getDate() + 1);
  }

  // Scan monthly sheets: count active days per Currency_PSP pair
  let totalDaysInRange = 0;
  const merchantActiveDays = {};
  const merchantMeta = {};
  const allUsedCurrencies = new Set();

  targetMonths.forEach(monthStr => {
    const monthSheet = ss.getSheetByName(monthStr);
    if (!monthSheet) return;

    const mValues = monthSheet.getDataRange().getDisplayValues();
    if (mValues.length < 2) return;

    const mHeaders = mValues[0];
    const mBody    = mValues.slice(1);

    // Find columns in date range
    const activeColIndexes = [];
    for (let c = 4; c < mHeaders.length; c++) {
      const d = mHeaders[c].trim();
      if (d >= startStr && d <= endStr) activeColIndexes.push(c);
    }
    totalDaysInRange += activeColIndexes.length;

    mBody.forEach(row => {
      const curr  = row[0].trim();
      const psp   = row[1].trim();
      const avg   = parseFloat(row[2]) || 0;
      const grade = row[3].trim();
      if (!curr || !psp || psp === "None") return;

      allUsedCurrencies.add(curr);
      const key = `${curr}_${psp}`;

      let activeDays = 0;
      activeColIndexes.forEach(ci => {
        const v = row[ci];
        if (v !== "" && !isNaN(v) && parseFloat(v) > 0) activeDays++;
      });

      merchantActiveDays[key] = (merchantActiveDays[key] || 0) + activeDays;
      if (!merchantMeta[key]) {
        merchantMeta[key] = { curr, psp, avgScore: avg, grade };
      } else {
        merchantMeta[key].avgScore = avg;
        merchantMeta[key].grade    = grade;
      }
    });
  });

  if (totalDaysInRange === 0) {
    SpreadsheetApp.getUi().alert("⚠️ 在指定月份表中找不到此區間的數據欄位！");
    return;
  }

  // Core merchant threshold: active on > 1/3 of total days
  const threshold = totalDaysInRange / 3;
  const currencyGroups = {};
  Object.keys(merchantActiveDays).forEach(key => {
    if (merchantActiveDays[key] > threshold) {
      const m = merchantMeta[key];
      if (!currencyGroups[m.curr]) currencyGroups[m.curr] = [];
      currencyGroups[m.curr].push({ psp: m.psp, avgScore: m.avgScore, grade: m.grade });
    }
  });

  ss.toast("正在地毯式掃描並精準加總成功金額...", "🔍 抓取跑量中");
  const volumeMap = fetchCurrencySuccessVolumeAdvanced(startStr, endStr, targetMonths);

  // Clear old data
  const maxRows = dataSheet.getMaxRows();
  if (maxRows >= 2) {
    dataSheet.getRange(2, 6,  maxRows - 1, 4).clearContent();
    dataSheet.getRange(2, 11, maxRows - 1, 8).clearContent();
    dataSheet.getRange(2, 20, maxRows - 1, 4).clearContent().setBackground(null);
  }

  // Headers
  dataSheet.getRange("F1:W1").setValues([[
    "有使用幣別","對應跑量(自動)","核心商戶數","平均跑量(公式)","",
    "商戶幣別","商戶名稱","月份均分","績效評級",
    "HIGH_CRITICAL(100%)","HIGH_WARNING(80%)","LOW_WARNING(40%)","LOW_CRITICAL(15%)","",
    "整數_100%","整數_80%","整數_40%","整數_15%"
  ]]).setBackground("#444444").setFontColor("white").setFontWeight("bold").setHorizontalAlignment("center");

  const sortedCurrencies = Array.from(allUsedCurrencies).sort();

  // Left block: currency summary
  const leftVals = [];
  const leftFormulas = [];
  sortedCurrencies.forEach((curr, idx) => {
    const r = idx + 2;
    leftVals.push([curr, volumeMap[curr] || 0, currencyGroups[curr] ? currencyGroups[curr].length : 0]);
    leftFormulas.push([`=IFERROR(G${r} / H${r} / $A$4, 0)`]);
  });
  if (leftVals.length > 0) {
    dataSheet.getRange(2, 6, leftVals.length, 3).setValues(leftVals).setHorizontalAlignment("center");
    dataSheet.getRange(2, 7, leftVals.length, 1).setNumberFormat("#,##0");
    dataSheet.getRange(2, 9, leftFormulas.length, 1).setFormulas(leftFormulas).setHorizontalAlignment("center").setNumberFormat("#,##0.0");
  }

  // Right block: PSP detail + water level formulas
  const rightRows = [];
  sortedCurrencies.forEach(curr => {
    (currencyGroups[curr] || []).forEach(item => {
      rightRows.push([curr, item.psp, item.avgScore, item.grade]);
    });
  });

  if (rightRows.length > 0) {
    dataSheet.getRange(2, 11, rightRows.length, 4).setValues(rightRows).setHorizontalAlignment("center");
    dataSheet.getRange(2, 13, rightRows.length, 1).setNumberFormat("0.0");

    const rawFormulas     = [];
    const roundedFormulas = [];

    // Grade multipliers based on geometric series (1.2x per grade step, C+ = 1.0 baseline)
    const multiplierIfs = `IFS($N{r}="A+",2.0736,$N{r}="A",1.728,$N{r}="B+",1.44,$N{r}="B",1.2,$N{r}="C+",1.0,$N{r}="C",0.8333,$N{r}="D+",0.6944,TRUE,0.5787)`;

    for (let i = 0; i < rightRows.length; i++) {
      const r   = i + 2;
      const mif = multiplierIfs.replace(/{r}/g, r);
      const base = `IFERROR(VLOOKUP($K${r},$F$2:$I$50,4,FALSE),0)`;
      const vol  = `(${base} * ${mif})`;

      rawFormulas.push([
        `=IF(${base}=0,"",(${vol})*1.0)`,
        `=IF(${base}=0,"",(${vol})*0.8)`,
        `=IF(${base}=0,"",(${vol})*0.4)`,
        `=IF(${base}=0,"",(${vol})*0.15)`
      ]);

      // Significant-figure rounding: ROUND(x, 1 - INT(LOG10(ABS(x))))
      const sfRound = col => `=IF(${col}${r}="","",IF(${col}${r}=0,0,ROUND(${col}${r},1-INT(LOG10(ABS(${col}${r}))))))`;
      roundedFormulas.push([sfRound("O"), sfRound("P"), sfRound("Q"), sfRound("R")]);
    }

    dataSheet.getRange(2, 15, rawFormulas.length, 4).setFormulas(rawFormulas).setHorizontalAlignment("center").setNumberFormat("#,##0");

    const roundedRange = dataSheet.getRange(2, 20, roundedFormulas.length, 4);
    roundedRange.setFormulas(roundedFormulas).setHorizontalAlignment("center").setNumberFormat("#,##0");
    roundedRange.setBackground("#e6f2ff"); // Light blue: these are the values to copy into routing config
  }

  dataSheet.getRange(1, 6, dataSheet.getLastRow(), 18)
    .setBorder(true, true, true, true, true, true, "#ccc", SpreadsheetApp.BorderStyle.SOLID);

  ss.toast("水位線計算完成！請查看 T:W 欄（淺藍色）的整數水位值。", "🚀 完成");
}

// ==================== Module 4: Volume deep-scan ====================

/**
 * Opens Drive folders (PARENT → YYYY-MM → YYYY_MM PSP分析 spreadsheet),
 * then full-grid scans each daily sheet for 幣別: markers and sums 24h of 成功金額.
 */
function fetchCurrencySuccessVolumeAdvanced(startStr, endStr, targetMonths) {
  const volumeMap = {};
  const parentFolder = DriveApp.getFolderById(WATERMARK_CONFIG.PARENT_FOLDER_ID);

  targetMonths.forEach(monthStr => {
    const [year, month]  = monthStr.split("-");
    const subFolderName  = `${year}-${month}`;
    const exactFileName  = `${year}_${month} PSP分析`;

    const subFolders = parentFolder.getFoldersByName(subFolderName);
    if (!subFolders.hasNext()) return;
    const targetFolder = subFolders.next();

    const files = targetFolder.searchFiles(
      `title = '${exactFileName}' and mimeType = 'application/vnd.google-apps.spreadsheet' and trashed = false`
    );
    if (!files.hasNext()) return;

    const pspSS = SpreadsheetApp.openById(files.next().getId());

    pspSS.getSheets().forEach(sheet => {
      const sheetName = sheet.getName().trim();
      // Process only daily 代收分析 sheets within date range
      if (!sheetName.includes("-代收分析")) return;
      const sheetDate = sheetName.substring(0, 10);
      if (sheetDate < startStr || sheetDate > endStr) return;

      const data = sheet.getDataRange().getDisplayValues();

      // Full-grid scan for 幣別: markers
      for (let r = 0; r < data.length; r++) {
        for (let c = 0; c < data[r].length; c++) {
          const cell = data[r][c].trim();
          if (!cell.startsWith("幣別:")) continue;

          const currencyName  = cell.replace("幣別:", "").trim();
          const subHeaderRow  = r + 1;
          if (subHeaderRow >= data.length) continue;

          // Scan right from current column to find 成功金額 within this currency block
          for (let sc = c; sc < Math.min(data[subHeaderRow].length, c + 30); sc++) {
            // Stop if we've crossed into the next currency block
            if (sc > c && data[r][sc] && data[r][sc].trim().startsWith("幣別:")) break;

            if (data[subHeaderRow][sc].trim() === "成功金額") {
              // Sum 24 hourly rows below the sub-header
              let blockSum = 0;
              for (let h = 0; h < 24; h++) {
                const dataRowIdx = r + 2 + h;
                if (dataRowIdx >= data.length) break;
                blockSum += parseFloat(data[dataRowIdx][sc].replace(/,/g, "")) || 0;
              }
              volumeMap[currencyName] = (volumeMap[currencyName] || 0) + blockSum;
              break;
            }
          }
        }
      }
    });
  });

  return volumeMap;
}
