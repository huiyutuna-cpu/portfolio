# Case 03 — Daily OPS Data Analysis System

> **Role**: Payment Operations Analyst @ BTSE (Bestpay)
> **Tools**: Google Apps Script · Google Sheets · Telegram Bot API · Slack API · Google Drive API
> **Timeline**: Built iteratively → runs daily, auto-triggered; consolidation in progress
> **Scale**: 30+ merchants · 8 currencies (CNY, VND, IDR, INR, PHP, THB, BDT, USDT) · 10+ PSPs
> **GitHub**: [case03_source_analysis.gs](./code/case03_source_analysis.gs) · [case03_psp_analysis_core.gs](./code/case03_psp_analysis_core.gs) · [case03_unified_long_table.gs](./code/case03_unified_long_table.gs)

---

## Problem

No centralized daily analytics pipeline existed for OPS team.

Each day, analysts needed to manually:
- Download CSV exports from two sources (代收 / 代付)
- Import into spreadsheets and cross-reference PSP / merchant data across 8+ currencies
- Identify underperforming PSPs, slow orders, and volume anomalies
- Push reports to Telegram manually

This blocked real-time visibility and delayed issue response.

---

## System Architecture

```
[Drive 資料夾]  ← 每天上傳 代收CSV / 代付CSV
       ↓ (每 5 分鐘自動掃描)
[Source Sheets]
  ├── 上分訂單分析 (代收 raw data)
  └── 回分訂單分析 (代付 raw data)
       ↓ (auto-transfer to per-currency sheets)
[Per-Currency Analysis Sheets]
  ├── 法幣: CNY / VND / IDR / INR / PHP / THB / BDT
  └── 虛擬幣: USDT
       ↓
[PSP 分析 Spreadsheet]  ← 核心系統
  ├── Dashboard (日期設定 / 超時閾值)
  ├── 24h 逐小時分析報表 (代收 / 代付各一張)
  ├── Smart_Insights：PSP 評分表 + 熱點圖 + 佔比圖
  └── PSP 月度歷史績效表 (外部雲端，長期追蹤)
       ↓ [TG / Slack 自動推播]
[統一長表]  ← 整合中 (2026-06 開始)
  └── 所有代收 / 代付資料聚合到單一 Sheet
       ↓ [未來 → 網頁 Dashboard，多角色分頁]
```

---

## What I Built

### Module 1 — Auto Data Ingestion
- `pollDriveFolder()` 每 5 分鐘掃描 Drive 資料夾，偵測代收 / 代付 CSV
- 兩份 CSV 到齊後自動觸發全流程（Step A→B→C→D，逐步排程）
- `_importCsvAsSheet()` 將 CSV 轉入 Google Sheet 新分頁：
  - 支援 UTF-8 BOM 自動去除
  - 代付專用：自動刪除「超时处理订单」S 欄 + 計算「處理時間」欄（完成時間 - 申請時間 → HH:MM:SS）
- 缺檔超過 30 分鐘自動推 TG 催促，避免漏報

### Module 2 — Per-Currency PSP Statistics
**代收（法幣 + 虛擬幣分開）**
- `分組分析_商戶幣別PSP統計完整()` — 法幣代收
- `分組分析_商戶幣別PSP統計虛擬幣()` — 虛擬幣代收

**代付（法幣 + 虛擬幣共用一套邏輯，filter 不同）**
- `執行所有商戶代付統計()` + `核心代付處理邏輯()` — 法幣與虛擬幣各傳入不同 filter

**輸出欄位**：商戶 × PSP × 類型 × 幣別 → 總筆數 / 成功筆數 / 失敗筆數 / 成功率 / PSP佔比 / 成功金額 / 平均完成時間

**日期偏移寫入**：每天寫進幣別分頁固定位置（日期 1 號 → 第 4 欄，每天 7 欄 block）

### Module 3 — Daily Volume Aggregation + TG Push
- `彙總所有幣別狀況()` — 法幣 & 虛擬幣各一套，邏輯相同，TG hashtag 不同
- 每幣別輸出：上分 / 回分 成功筆數、成功率、成功金額、回調時間
- **波動提醒**：與前一天比較，±10% 以上商戶自動分類推 TG（▲成長 N 家 / ▼下滑 N 家 / 新增商戶 / 減少商戶）
- `checkAllSheetsVolumeDrop()` — 各幣別 Tab 跑量下跌 >10% 即時告警
- `sendSlowOrdersToTG_ByAnalysisA1()` — VND 慢單 >10 分鐘且超 50 筆推 TG
- `analyzePSPGaps_AllCurrencies_WithTxnID()` — 代付 PSP 兩單之間間隔 >100 分鐘偵測（吊單率）

### Module 4 — PSP Scoring System（核心）
每日對所有 PSP 進行五維度量化評分：

| 維度 | 權重 | 計算方式 |
|------|------|----------|
| 時效分 | 30% | `max(0, 1 − avg_sec / limit_sec)` |
| 成功分 | 30% | 成功率 |
| 超時分 | 30% | `1 − 超時率`（成功率=0 時強制 0，避免無單可超時卻得滿分） |
| 金額分 | 5%  | PSP 金額 / 幣別總金額 |
| 筆數分 | 5%  | PSP 筆數 / 幣別總筆數 |

**8 級評等**：A+(≥90) · A(≥85) · B+(≥75) · B(≥65) · C+(≥55) · C(≥45) · D+(≥35) · D

超時閾值由 Dashboard `C2:E100` 設定，不同幣別可設不同分鐘數。

### Module 5 — Smart_Insights Heatmaps
- **24h 筆數熱點圖**：每小時顯示「提單筆數 / 成功筆數」，高峰時段自動變紅
- **24h 金額熱點圖**：每小時提單金額，高峰時段自動變綠
- **PSP 佔比熱點圖**：每小時 PSP 佔比，≥80% 黃色警示（集中度風險）
- 超時單：對應格變紅 + Note 記錄超時交易 ID，方便追單

### Module 6 — Monthly PSP History
- `recordPspHistory()` — 每日評分自動存入外部雲端 Spreadsheet
- 月份分頁：`幣別 | PSP | 平均分數 | 績效評級` + 日期橫向展開
- 新 PSP 自動插行（依幣別+PSP名排序），新日期自動插欄（依日期排序）
- 優（≥85）綠底 / 差（<60）紅底，視覺快速判讀

### Module 7 — Unified Long Table（整合中）
- `更新統一長表()` — 支援 Google Sheet / Excel(.xlsx) / CSV 三種來源
- 統一欄位：日期 / 商戶 / 幣別 / PSP / 類型 / 時段 / 小時序 / 總筆數 / 成功筆數 / 失敗筆數 / 提單金額 / 成功金額 / 代理費 / 平台費 / 總耗時秒 / 平均耗時
- 每 10 分鐘自動掃描資料夾，整天複寫（同日資料先刪後寫），Email 通知
- 目的：取代目前多檔分散架構，為未來網頁 Dashboard 鋪路

---

## Outcome

| Metric | Before | After |
|--------|--------|-------|
| 每日分析流程 | 手動下載 CSV → 手動匯入 → 手動分析 | 上傳 Drive → 15 分鐘內全自動完成 |
| PSP 評估方式 | 主觀判斷 / 口頭討論 | 5 維度量化評分 + 8 級評等 |
| 波動告警 | 需要盯著數字自己發現 | ±10% 商戶自動推 TG |
| 慢單偵測 | 無 | VND >10分鐘 / >50筆 即時通知 |
| 吊單監控 | 無 | 代付 PSP 兩單間隔 >100 分鐘自動偵測 |
| 月度績效追蹤 | 無 | 每日評分自動存入外部歷史表，可回溯 |
| 資料覆蓋 | 部分幣別手動 | 8 幣別 / 30+ 商戶 / 10+ PSP 全覆蓋 |

---

## Key Design Decisions

**1. 幣別分流邏輯**
INR / PHP 依 PSP 名稱自動路由到不同分頁（INR-Worldpay vs INR-喚醒 / PHP-Maya vs PHP-Gcash），同一幣別不同通道邏輯拆開，避免數據混雜影響分析。

**2. 超時率與成功率連動**
成功率為 0 時，超時分強制歸 0。原始設計缺這個 guard，會導致「無成功訂單的 PSP 因無超時而得高分」的反直覺結果。

**3. 代付 CSV 前處理**
代付 CSV 原始有「超时处理订单」欄，與分析無關且佔欄位。匯入時自動刪除，同時在最後新增「處理時間」欄（HH:MM:SS），供後續分析直接使用。

**4. 統一長表設計**
從多檔 → 一張長表，目的是讓後續網頁 Dashboard 只需對接一個資料源。欄位設計考慮多角色查詢需求：OPS 看即時狀況 / AM 看商戶趨勢 / 管理層看幣別跑量。

---

## Update Log

| Date | Update |
|------|--------|
| 2026-06-28 | 架構文件與現有全部 code 記錄完成（3 個 .gs 檔） |
| [TBD] | 統一長表系統整合完成後更新架構圖 |
| [TBD] | 網頁 Dashboard 完成後補截圖 + 說明 |

---

*Full case study with screenshots → [Notion Portfolio](https://notion.so) (link TBD)*
