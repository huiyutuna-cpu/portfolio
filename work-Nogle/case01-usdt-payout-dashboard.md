# Case 01 — USDT Payout Profit & Capital Flow Dashboard

> **Role**: Payment Operations Analyst @ Nogle
> **Tools**: Google Apps Script · Google Sheets · Chart.js · HTML/CSS
> **Timeline**: 2–3 months build → fully automated, runs monthly

---

## Problem

No centralized visibility into USDT payout profitability or merchant capital flow.
Exchange rate calculations were semi-automated and frequently inaccurate.
Fee rate adjustments were made without reliable margin data.

Scale: 30+ merchants, multiple currencies (CNY, VND, IDR, INR, PHP, THB, USDT).

---

## What I Built

**System 1 — USDT Payout Profit Analysis** ([code](./code/case01_usdt_payout_analysis.gs))
- Pulls raw payout data from source spreadsheet across 30+ merchants
- Groups by merchant × currency × month; calculates rate spread (匯差) per group
- Incremental update: historical months locked, never re-processed
- Auto-triggers on file open

**System 2 — Merchant Payment Ratio Analysis** ([code](./code/case01_usdt_payout_analysis.gs))
- Cross-references Payin/Payout sheets to calculate:
  - Deposited capital = Payin − Payout − U Return
  - Capital holding ratio (資金存放率) = Deposited / Payin
- Monthly breakdown per merchant

**Web Dashboard**
- Google Apps Script Web App with live Chart.js charts
- Filters by currency and merchant
- Built with Chart.js + Tailwind CSS

---

## Outcome

| Metric | Result |
|--------|--------|
| Exchange rate accuracy | Fully automated, zero manual correction |
| Capital visibility | Management sees merchant fund holdings monthly |
| Decisions enabled | Fee adjustments, channel analysis, liquidity management |
| Reporting effort | Auto-runs monthly, no manual input |

---

## Key Design Decision

Built closed-month locking early — historical data is locked once a month passes so incremental updates never accidentally overwrite finalized records. Critical at 30+ merchant scale.

---

*Full case study with screenshots → [Notion Portfolio](https://notion.so) (link TBD)*
