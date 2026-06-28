/**
 * Case 03 — Daily OPS Data Analysis System
 * Module: Unified Long Table — All historical transaction data in one sheet
 *
 * Architecture:
 *   Drive 資料夾（代收/代付 CSV/XLSX/Google Sheet）
 *     → 每 10 分鐘自動掃描
 *     → 解析、聚合（日期/商戶/幣別/PSP/類型/小時）
 *     → 寫入統一長表（整天複寫）
 *     → Email 通知（含新增/覆蓋明細 + 已刪來源檔清單）
 *
 * Output columns:
 *   日期 | 商戶 | 幣別 | PSP | 類型 | 時段 | 小時序
 *   總筆數 | 成功筆數 | 失敗筆數 | 提單金額 | 成功金額
 *   代理費 | 平台費 | 總耗時秒 | 平均耗時
 *
 * Purpose: Replace scattered multi-file architecture with a single source of truth
 *          for future web-based dashboard (multi-role views).
 */

// ★ 使用前需在左側「服務」加入 Drive API (v2)
const CONFIG_LT = {
  SOURCE_FOLDER_ID: "YOUR_SOURCE_FOLDER_ID",       // 每天上傳 CSV/XLSX 的 Drive 資料夾
  LONG_TABLE_SS_ID: "YOUR_LONG_TABLE_SS_ID",        // 統一長表試算表 ID
  LONG_TABLE_SHEET: "長表data",
  COLLECT_KEYWORD:  "代收",
  PAYOUT_KEYWORD:   "代付",
  NOTIFY_EMAIL:     "YOUR_EMAIL",
  TIMEZONE:         "GMT+8",
  HEADERS: {
    merchant:    "分站名称",
    applyTime:   "申请时间",
    type:        "类别",
    psp:         "通道名称",
    currency:    "币别",
    amount:      "金额(元)",
    agentFee:    "代理费(元)",
    platformFee: "通道费(元)",
    finishTime:  "完成时间",
    status:      "订单状态",
    payCategory: "支付大项",
    account:     "帐号"
  },
  STATUS_SUCCESS_KEYWORD: "完成"
};

const LONG_HEADER = [
  "日期", "商戶", "幣別", "PSP", "類型",
  "時段", "小時序",
  "總筆數", "成功筆數", "失敗筆數",
  "提單金額", "成功金額",
  "代理費", "平台費",
  "總耗時秒", "平均耗時"
];

// ============================================================
// 主入口
// ============================================================
function 更新統一長表() {
  const folder = DriveApp.getFolderById(CONFIG_LT.SOURCE_FOLDER_ID);
  const longSS = SpreadsheetApp.openById(CONFIG_LT.LONG_TABLE_SS_ID);
  const sheet  = getOrCreateLongTable_(longSS);

  const entries = scanFolder_(folder);
  if (entries.length===0) { Logger.log("資料夾沒有可處理的訂單檔，安靜結束。"); return; }

  const rowsByDate = {};
  const tempFilesToDelete = [], sourceFilesToDelete = [];

  entries.forEach(entry => {
    const part = aggregateFile_(entry, tempFilesToDelete);
    if (Object.keys(part).length>0) {
      Object.keys(part).forEach(d=>{
        if(!rowsByDate[d]) rowsByDate[d]=[];
        rowsByDate[d].push(...part[d]);
      });
      sourceFilesToDelete.push(entry.file);
    }
  });

  tempFilesToDelete.forEach(id=>{try{DriveApp.getFileById(id).setTrashed(true);}catch(e){}});

  const affectedDates=Object.keys(rowsByDate);
  if (affectedDates.length===0) { Logger.log("沒有解析出任何資料，安靜結束。"); return; }

  const mergedByDate={};
  affectedDates.forEach(d=>mergedByDate[d]=mergeRows_(rowsByDate[d]));

  const existingDates=getExistingDates_(sheet);
  const report=affectedDates.sort().map(d=>({date:d,rowCount:mergedByDate[d].length,isOverwrite:existingDates.has(d)}));

  rewriteLongTable_(sheet, mergedByDate);

  const deletedNames=[];
  sourceFilesToDelete.forEach(f=>{try{deletedNames.push(f.getName());f.setTrashed(true);}catch(e){Logger.log("刪除來源檔失敗:"+e);}});

  sendNotifyEmail_(sheet, report, deletedNames);
  Logger.log("✅ 完成。處理日期:"+report.map(r=>r.date).join(", ")+" | 已刪檔:"+deletedNames.join(", "));
}

// ============================================================
// 資料夾掃描
// ============================================================
function scanFolder_(folder) {
  const out=[];
  const files=folder.getFiles();
  while(files.hasNext()){
    const file=files.next();
    const name=file.getName(), mime=file.getMimeType(), lower=name.toLowerCase();
    const isSheet=(mime===MimeType.GOOGLE_SHEETS);
    const isExcel=(mime===MimeType.MICROSOFT_EXCEL||mime==="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"||lower.endsWith(".xlsx")||lower.endsWith(".xls"));
    const isCsv=(mime===MimeType.CSV||lower.endsWith(".csv"));
    if(!isSheet&&!isExcel&&!isCsv) continue;
    if(!name.includes(CONFIG_LT.COLLECT_KEYWORD)&&!name.includes(CONFIG_LT.PAYOUT_KEYWORD)) continue;
    out.push({file,fileType:isCsv?"csv":(isExcel?"excel":"sheet"),fileName:name});
  }
  return out;
}

// ============================================================
// 單檔聚合（日期從申請時間逐筆抽，GMT+8 顯示文字，不碰 Date 時區）
// ============================================================
function aggregateFile_(entry, tempFilesToDelete) {
  const file=entry.file;
  let values;

  if (entry.fileType==="csv") {
    let csvText=file.getBlob().getDataAsString("UTF-8");
    if(csvText.charCodeAt(0)===0xFEFF) csvText=csvText.slice(1);
    values=Utilities.parseCsv(csvText);
  } else {
    let ss;
    if (entry.fileType==="excel") {
      const blob=file.getBlob();
      const resource={title:"【暫存】"+file.getName(),mimeType:MimeType.GOOGLE_SHEETS};
      const converted=Drive.Files.insert(resource,blob,{convert:true});
      tempFilesToDelete.push(converted.id);
      ss=SpreadsheetApp.openById(converted.id);
    } else { ss=SpreadsheetApp.openById(file.getId()); }
    values=ss.getSheets()[0].getDataRange().getDisplayValues();
  }

  if(!values||values.length<2) return {};

  const headerRow=values[0].map(h=>String(h).replace(/^﻿/,"").trim());
  const col=resolveColumns_(headerRow);
  if(col===null){Logger.log(`⚠️ ${file.getName()} 缺必要表頭，略過`);return{};}

  const map=new Map();
  for(let i=1;i<values.length;i++){
    const row=values[i];
    const merchant=clean_(row[col.merchant]), currency=clean_(row[col.currency]);
    let psp=clean_(row[col.psp]), type=clean_(row[col.type]).replace(/\s/g,"");
    const status=clean_(row[col.status]), applyRaw=row[col.applyTime];
    if(!merchant||!currency||!psp||!type||!status||!applyRaw) continue;

    // 提現細分：下發USDT → 「提现-U」，PSP 從帳號欄取
    if(type==="提现"&&col.payCategory!==-1){
      const payCategory=clean_(row[col.payCategory]);
      if(/下[發发]\s*USDT/i.test(payCategory)){
        type="提现-U";
        if(col.account!==-1){const parts=clean_(row[col.account]).split(/\s*[-－]\s*/);const last=parts.length>1?clean_(parts[parts.length-1]):"";if(last)psp=last.toUpperCase();}
      }
    }

    const dt=parseApplyDateTime_(applyRaw);
    if(!dt){Logger.log(`⚠️ 申請時間解析失敗："${applyRaw}"`);continue;}

    const amount     =toNum_(row[col.amount]);
    const agentFee   =(col.agentFee===-1)?0:toNum_(row[col.agentFee]);
    const platformFee=(col.platformFee===-1)?0:toNum_(row[col.platformFee]);
    const isSuccess  =status.includes(CONFIG_LT.STATUS_SUCCESS_KEYWORD);

    const key=[dt.dateStr,merchant,currency,psp,type,dt.hour].join("||");
    if(!map.has(key)) map.set(key,{total:0,success:0,fail:0,amount:0,successAmount:0,agentFee:0,platformFee:0,totalSec:0});
    const r=map.get(key);
    r.total++; r.amount+=amount; r.agentFee+=agentFee; r.platformFee+=platformFee;
    if(isSuccess){
      r.success++; r.successAmount+=amount;
      const sec=(col.finishTime===-1)?-1:diffSeconds_(row[col.finishTime],applyRaw);
      if(sec>=0) r.totalSec+=sec;
    } else r.fail++;
  }

  const byDate={};
  for(const [key,r] of map.entries()){
    const [dateStr,merchant,currency,psp,type,hourStr]=key.split("||");
    const hour=parseInt(hourStr,10);
    const rowArr=[dateStr,merchant,currency,psp,type,hourLabel_(hour),hour,r.total,r.success,r.fail,Math.round(r.amount),Math.round(r.successAmount),Math.round(r.agentFee),Math.round(r.platformFee),Math.round(r.totalSec),""];
    if(!byDate[dateStr]) byDate[dateStr]=[];
    byDate[dateStr].push(rowArr);
  }
  return byDate;
}

// ============================================================
// 跨檔合併（同 key 數字相加）
// ============================================================
function mergeRows_(rows) {
  const map=new Map();
  rows.forEach(r=>{
    const key=[r[0],r[1],r[2],r[3],r[4],r[6]].join("||");
    if(!map.has(key)){map.set(key,r.slice());}
    else{const e=map.get(key);e[7]+=r[7];e[8]+=r[8];e[9]+=r[9];e[10]+=r[10];e[11]+=r[11];e[12]+=r[12];e[13]+=r[13];e[14]+=r[14];}
  });
  const out=Array.from(map.values());
  out.forEach(r=>{r[15]=r[8]>0?(r[14]/r[8])/86400:0;});  // 平均耗時 → Sheet 時間序列值
  return out;
}

// ============================================================
// 寫入長表（整天刪除 → 整批重寫 → 重排 → 清空白）
// ============================================================
function rewriteLongTable_(sheet, mergedByDate) {
  const lastRow=sheet.getLastRow();
  const overwriteDates=new Set(Object.keys(mergedByDate));
  let body=[];
  if(lastRow>=2){
    body=sheet.getRange(2,1,lastRow-1,LONG_HEADER.length).getValues();
    body=body.filter(r=>{
      const d=r[0] instanceof Date?Utilities.formatDate(r[0],CONFIG_LT.TIMEZONE,"yyyy-MM-dd"):String(r[0]).trim();
      return !overwriteDates.has(d);
    });
  }
  Object.keys(mergedByDate).forEach(d=>body.push(...mergedByDate[d]));
  body.forEach(r=>{const success=Number(r[8])||0;const totalSec=Number(r[14])||0;r[15]=success>0?(totalSec/success)/86400:0;});
  body.sort((a,b)=>{
    if(String(a[0])!==String(b[0])) return String(a[0])<String(b[0])?-1:1;
    if(a[6]!==b[6]) return a[6]-b[6];
    if(String(a[1])!==String(b[1])) return String(a[1])<String(b[1])?-1:1;
    return String(a[3])<String(b[3])?-1:(String(a[3])>String(b[3])?1:0);
  });
  if(sheet.getMaxRows()>1) sheet.getRange(2,1,sheet.getMaxRows()-1,LONG_HEADER.length).clearContent();
  if(body.length>0){
    sheet.getRange(2,1,body.length,LONG_HEADER.length).setValues(body);
    sheet.getRange(2,16,body.length,1).setNumberFormat("[h]:mm:ss");
  }
  const usedRows=body.length+1;
  const extra=sheet.getMaxRows()-usedRows;
  if(extra>0) sheet.deleteRows(usedRows+1,extra);
  SpreadsheetApp.flush();
}

// ============================================================
// Email 通知
// ============================================================
function sendNotifyEmail_(sheet, report, deletedNames) {
  const lastRow=sheet.getLastRow(), totalRows=Math.max(0,lastRow-1);
  let dateRange="—";
  if(totalRows>0){const first=sheet.getRange(2,1).getDisplayValue();const last=sheet.getRange(lastRow,1).getDisplayValue();dateRange=`${first} ～ ${last}`;}

  const added=report.filter(r=>!r.isOverwrite), overwritten=report.filter(r=>r.isOverwrite);
  let html=`<div style="font-family:Arial,sans-serif;font-size:14px;color:#333;"><h2 style="color:#1a73e8;">✅ 統一長表已更新完成</h2>`;
  if(added.length>0){html+=`<p><b>🆕 新增日期</b></p><ul>`;added.forEach(r=>html+=`<li>${r.date}：寫入 ${r.rowCount.toLocaleString()} 列</li>`);html+=`</ul>`;}
  if(overwritten.length>0){html+=`<p><b>♻️ 覆蓋日期</b></p><ul>`;overwritten.forEach(r=>html+=`<li>${r.date}：覆蓋為 ${r.rowCount.toLocaleString()} 列</li>`);html+=`</ul>`;}
  html+=`<hr style="border:none;border-top:1px solid #ddd;"><p><b>📊 長表現況</b></p><ul><li>總列數：${totalRows.toLocaleString()} 列</li><li>涵蓋日期：${dateRange}</li></ul>`;
  if(deletedNames&&deletedNames.length>0) html+=`<p style="background:#e8f5e9;padding:10px;border-radius:6px;">🗑️ 已自動刪除處理完的來源檔：<br>`+deletedNames.map(n=>"• "+n).join("<br>")+`</p>`;
  html+=`<p style="color:#999;font-size:12px;">執行時間：${Utilities.formatDate(new Date(),CONFIG_LT.TIMEZONE,"yyyy-MM-dd HH:mm:ss")}</p></div>`;

  const subject=`【長表更新】${report.map(r=>r.date).join("、")} 已完成（${report.length}天）`;
  MailApp.sendEmail({to:CONFIG_LT.NOTIFY_EMAIL,subject,htmlBody:html});
}

// ============================================================
// 共用工具
// ============================================================
function getOrCreateLongTable_(ss) {
  let sheet=ss.getSheetByName(CONFIG_LT.LONG_TABLE_SHEET);
  if(!sheet){
    sheet=ss.insertSheet(CONFIG_LT.LONG_TABLE_SHEET);
    sheet.getRange(1,1,1,LONG_HEADER.length).setValues([LONG_HEADER]).setFontWeight("bold").setBackground("#333333").setFontColor("white");
    sheet.setFrozenRows(1);
  }
  return sheet;
}
function getExistingDates_(sheet) {
  const set=new Set();
  const lastRow=sheet.getLastRow(); if(lastRow<2) return set;
  sheet.getRange(2,1,lastRow-1,1).getValues().forEach(r=>{
    const v=r[0];
    if(v instanceof Date) set.add(Utilities.formatDate(v,CONFIG_LT.TIMEZONE,"yyyy-MM-dd"));
    else if(v) set.add(String(v).trim());
  });
  return set;
}
function resolveColumns_(headerRow) {
  const H=CONFIG_LT.HEADERS;
  const col={};
  Object.keys(H).forEach(k=>col[k]=headerRow.indexOf(H[k]));
  for(const k of["merchant","applyTime","type","psp","currency","amount","status"]) if(col[k]===-1){Logger.log(`缺必要表頭:${H[k]}`);return null;}
  return col;
}

/**
 * 從申請時間顯示文字解析日期 + 小時
 * 純字串解析，完全不碰 Date 時區，顯示幾點就是幾點（已是 GMT+8 顯示值）
 */
function parseApplyDateTime_(v) {
  const s=String(v).trim();
  const m=s.match(/(\d{4})[\/\-](\d{1,2})[\/\-](\d{1,2})\D+(\d{1,2}):(\d{2})(?::(\d{2}))?/);
  if(m){
    const dateStr=`${m[1]}-${String(+m[2]).padStart(2,"0")}-${String(+m[3]).padStart(2,"0")}`;
    const hour=parseInt(m[4],10);
    if(hour<0||hour>23) return null;
    return{dateStr,hour};
  }
  const m2=s.match(/(\d{4})[\/\-](\d{1,2})[\/\-](\d{1,2})/);
  if(m2) return{dateStr:`${m2[1]}-${String(+m2[2]).padStart(2,"0")}-${String(+m2[3]).padStart(2,"0")}`,hour:0};
  return null;
}
function diffSeconds_(finishV, applyV) {
  const f=parseToEpochSec_(finishV), a=parseToEpochSec_(applyV);
  if(f===null||a===null) return -1;
  return f-a;
}
function parseToEpochSec_(v) {
  const s=String(v).trim();
  const m=s.match(/(\d{4})[\/\-](\d{1,2})[\/\-](\d{1,2})\D+(\d{1,2}):(\d{2})(?::(\d{2}))?/);
  if(!m) return null;
  return Date.UTC(+m[1],+m[2]-1,+m[3],+m[4],+m[5],m[6]?+m[6]:0)/1000;
}
function clean_(v){ return(v==null?"":String(v)).trim(); }
function toNum_(v){ if(v==null||v==="")return 0; const n=parseFloat(String(v).replace(/,/g,"")); return isNaN(n)?0:n; }
function hourLabel_(h){ const hh=String(h).padStart(2,"0"); return `${hh}:00 - ${hh}:59`; }

// ============================================================
// 自動觸發（每 10 分鐘掃描一次）
// ============================================================
function 設定自動掃描() {
  ScriptApp.getProjectTriggers().forEach(t=>{if(t.getHandlerFunction()==="更新統一長表")ScriptApp.deleteTrigger(t);});
  ScriptApp.newTrigger("更新統一長表").timeBased().everyMinutes(10).create();
  Logger.log("✅ 已設定每 10 分鐘自動掃描資料夾");
}
