# Case Study 01 — USDT Payout Profit & Capital Flow Dashboard

> **Role**: Payment Operations Analyst @ Nogle (FinTech, Southeast Asia)
> **Timeline**: 2–3 months to build, now runs automatically every month
> **Tools**: Google Apps Script, AppSheet, Google Sheets, Chart.js, HTML/CSS

---

## Background

Working as the only analyst responsible for USDT payout reconciliation across 30+ merchants, I was in charge of tracking whether each payout transaction was profitable — and how much capital each merchant held in our system at any given time.

The previous process relied on a semi-automated setup, but the exchange rate calculations were inaccurate and required manual correction every month. There was no centralized view of:

- Which merchants were generating healthy exchange rate margins
- Which channels were underperforming
- How much capital was sitting idle (deposited but not yet returned)

Without accurate data, fee rate decisions were made on gut feel, and capital flow was largely invisible.

---

## Problem

| | Before |
|---|---|
| Exchange rate tracking | Semi-automated, frequently inaccurate |
| Merchant capital visibility | None — no one knew how much each merchant was holding |
| Decision making | Fee rate adjustments made without reliable data |
| Reporting | Required manual correction every month |
| Scale | 30+ merchants across multiple currencies |

**I identified this gap myself. No one had flagged it as a priority.**

---

## My Role

I owned this project end-to-end:

- **Discovery**: Identified that inaccurate exchange rate data was causing fee decisions to be made on unreliable baselines
- **Scoping**: Defined what metrics actually mattered to OPS and management (rate margin, fund holding ratio)
- **Design**: Designed the data model — how to calculate 匯差 (rate spread), 存放金額 (deposited capital), 資金存放率 (capital holding ratio)
- **Build**: Wrote all Google Apps Script logic, cross-spreadsheet data pulls, and the web dashboard
- **Adoption**: System was eventually handed off to a new Operations Finance team member

---

## Solution

Built two interconnected automated systems inside Google Sheets, with a web dashboard frontend:

### System 1 — USDT Payout Profit Analysis
- Automatically pulls raw payout data from source spreadsheet (30+ merchants)
- Groups by merchant × currency × month
- Calculates per group: transaction count, total payout amount, average 3rd-party exchange rate, actual exchange rate, **rate spread (profit margin)**
- Incremental update logic: historical months are locked and never re-processed
- Runs automatically on file open; can also be triggered manually

### System 2 — Merchant Payment Ratio Analysis
- Cross-references Payin and Payout sheets across all merchants
- Calculates for each merchant per month:
  - Total payin / payout amounts
  - Payout ratio (代付佔比)
  - Total U return amount & ratio (回U佔比)
  - **Deposited capital = Payin − Payout − U Return**
  - **Capital holding ratio = Deposited capital / Payin**

### Web Dashboard
- Deployed as a Google Apps Script Web App
- Interactive filters: by currency, by merchant
- 4 live charts: monthly transaction count, payout volume, rate spread trend, exchange rate fluctuation
- Built with Chart.js + Tailwind CSS

---

## Outcome

| | After |
|---|---|
| Exchange rate accuracy | Fully automated, no manual correction needed |
| Capital visibility | Management can see exactly how much each merchant holds, monthly |
| Business decisions enabled | Fee rate adjustments, underperforming channel identification, liquidity management |
| Reporting effort | Auto-runs monthly, zero manual input |
| Scale | 30+ merchants tracked simultaneously |

**Three types of decisions were made using this data:**
1. Fee rate adjustments for specific merchants based on actual margin data
2. Identification of channels/merchants with insufficient rate spread
3. Capital liquidity management — understanding which merchants had high fund holding ratios

---

## What I Learned

- **Data visibility changes decisions.** Before this system, fee rates were set without margin data. After, adjustments became data-driven.
- **Incremental update design matters at scale.** With 30+ merchants and growing history, rebuilding all data every month would be slow and error-prone. Building the closed-month lock-in logic early saved a lot of headaches.
- **Cross-functional utility.** I built this for OPS, but Finance and management both ended up using it for different decisions — which validated that the data model was comprehensive.

---

*Next: add screenshots of the dashboard and sample output (anonymized)*
