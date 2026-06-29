# Case 02 — Monthly Merchant Volume & Revenue Analysis System

> **Role**: Payment Operations Analyst @ BTSE (Bestpay)
> **Tools**: Google Apps Script · Google Sheets · Telegram Bot API
> **Timeline**: Built iteratively → runs monthly, output delivered via TG + Sheets
> **GitHub**: [code](./code/case02_monthly_revenue_analysis.gs)

---

## Problem

No monthly report existed for merchant volume changes or revenue trends.

Each month, Operations needed to manually:
- Track which merchants dropped off or went live across 8+ currencies
- Calculate revenue per currency/type for executive reporting
- Write analysis narratives with no automated data pipeline

Scale: 30+ merchants, 8 currencies (CNY, VND, IDR, INR, PHP, BDT, USDT, IDR), monthly report to management.

---

## What I Built

**System 1 — Merchant Volume Diff Report** ([code](./code/case02_monthly_revenue_analysis.gs))
- Compares two monthly Sheets across all currency tabs
- Outputs: 🟠 missing merchants / 🟢 new merchants / 🔵 stable merchants
- Filters test accounts; sorts by currency + merchant name
- Volume change per merchant: highlights >150% growth (yellow) / <−60% drop (pink)
- Auto-pushes summary to Telegram group on run

**System 2 — Monthly Revenue Analysis** ([code](./code/case02_monthly_revenue_analysis.gs))
- Reads Payin / Payout sheets for two months from a linked external Spreadsheet
- Aggregates by currency: total count, success count, volume, revenue
- Generates a new comparison sheet with Δ columns (Current − Previous)
- Covers: 上分 (Payin) / 回分 (Payout) / 提現 (Withdrawal)

**System 3 — Revenue Ratio Report** ([code](./code/case02_monthly_revenue_analysis.gs))
- Per-merchant revenue ratio (Revenue ÷ Amount) per currency
- Month-over-month comparison with ▲▼ color-coded indicators
- Outputs directly to "收益比" sheet; clears and rewrites on each run

---

## Outcome

| Metric | Result |
|--------|--------|
| Before | Zero monthly report — analysis was entirely ad hoc |
| Merchant diff detection | Automated across 8 currencies, runs in <1 min |
| Revenue report | Full 2-month comparison auto-built into a new sheet |
| TG delivery | Diff + volume change summary auto-sent to Ops group |
| Report depth | Volume, success rate, revenue, revenue ratio per merchant |

---

## Key Design Decision

Three independent systems that feed one analyst-written narrative.
Merchant diff answers *who dropped* — revenue analysis answers *how much revenue moved* — revenue ratio flags *where margin is being eroded*.
This maps directly to the three questions management asks every month.

---

*Full case study with screenshots → [Notion Portfolio](https://notion.so) (link TBD)*
