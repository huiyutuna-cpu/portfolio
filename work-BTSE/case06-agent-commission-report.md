# Case 06 — Monthly Agent Commission Report Automation

> **Role**: Payment Operations Analyst @ BTSE (Bestpay)
> **Tools**: Google Apps Script · Google Sheets
> **Timeline**: Run monthly (first week, covering prior month)
> **Scale**: 2 platforms (HTPay / BestPay) · multiple agents · 5+ currencies
> **GitHub**: [case06_htpay_agent_report.gs](./code/case06_htpay_agent_report.gs) · [case06_bestpay_agent_report.gs](./code/case06_bestpay_agent_report.gs)

---

## Problem

At month-end, agent commission reports needed to be manually compiled and sent to agents and Finance for verification.

Before this system:
- Data had to be manually filtered, grouped by agent and currency, and formatted
- Two platforms (HTPay, BestPay) have different data structures — no unified process
- Errors in aggregation delayed report delivery and caused disputes with agents
- Finance had no automated summary to verify total volumes vs. agent fees

---

## System Architecture

```
HTPay Flow:
[PayIn-Data sheet]  [Payout-Data sheet]
   (raw row data)      (raw row data)
         ↓
[generateAgentReports()]
  ├── Filter: exclude Chinese-name agents (internal accounts)
  ├── Per-agent tab: per-currency horizontal blocks
  │     └── Pay In block → Pay Out block → Total (Payin+Payout)
  ├── 總計 sheet: all data by Month+Currency (no agent filter)
  └── 代理總計 sheet: agent fee summary by Month+Agent+Currency

BestPay Flow:
[Payin-DD sheet]  [Payout-DD sheet]  [代理名單 sheet]
  (pivot table)     (pivot table)    (merchant→agent+rates)
         ↓
[generateAgentSheets()]
  ├── Determine main month from date frequency
  ├── Build full calendar: day 1 ~ month-end (zero-fill missing days)
  └── Per-agent tab: Payin (left) | Payout (right)
        ├── Payin: Date + per-merchant amounts + Subtotal + Fee
        └── Payout: Date + per-merchant amounts + Subtotal + Net Subtotal + Fee
                    (Net Subtotal = Amount ÷ 1.04; Fee = Net × rate%)
```

---

## What I Built

### HTPay Module

**Data reading** (`readSourceRowsForAgentSheets_`):
- Reads PayIn-Data (cols A/D/E/F/G/L/M) and Payout-Data (cols A/D/E/F/G/J/K)
- Normalizes date to `YYYY-MM` format, handles both Date objects and string formats
- Filters out agents whose names contain Chinese characters — these are internal accounts handled separately

**Per-agent sheet generation**:
- One tab per agent (A–Z sorted)
- Horizontal layout: each currency gets a 5-column block (Date/Month · accountId · total_count · total_amount · total_agentFee), separated by 2 blank columns
- Within each block: Pay In section → Pay Out section → Total row (cross-references Payin Total + Payout Total with Sheets formulas)
- Aggregation: rows with same Date + accountId are summed before writing

**Summary sheets**:
- `總計`: reads raw source directly, aggregates by Month + Currency regardless of agent — used by Finance to cross-check total platform volume
- `代理總計`: aggregates agent fees by Month + Agent + Currency — used as the single-line payment summary for Finance approval

**Chinese-agent exclusion design**: the filter only applies to individual agent tabs. The `總計` sheet reads all rows (including Chinese-name agents), so Finance always sees 100% of volume even when some agents are internal.

### BestPay Module

**Merchant-to-agent mapping** (`loadAgentMapping_`):
- Separate `代理名單` sheet: columns for 商戶名 / 代理名 / 代理代收手續費(%) / 代理代付手續費(%)
- Rates stored as-is (e.g. `0.2` = 0.2%) — zero-configuration for new merchants by editing the sheet
- Missing rate = blank cell = that merchant's fee not calculated (prevents incorrect charges)

**Pivot reading** (`readPivot_`):
- Scans first 20 rows to find the header row containing "Date" or "Transaction Date"
- Reads until hitting a "總和" (grand total) row
- Graceful degradation: if either pivot sheet has no recognizable header, shows a confirm dialog and continues with the other side rather than crashing

**Full-month calendar**:
- Determines the "main month" by finding the most-frequent YYYY-MM in the source data
- Generates day 1 through month-end as a date list
- Every day is written to the sheet; days with no transactions show 0 — agents can see the full calendar and verify days match

**INR Payout fee calculation**:
- BestPay's INR payout amounts include GST (×1.04)
- Net Subtotal = Amount ÷ 1.04 (GST-exclusive base)
- Agent fee = Net Subtotal × rate%
- An extra "Net Subtotal" column shows agents the actual base before GST

**Layout**: Payin block on the left, Payout block immediately to the right with 2-column gap. Total row at the bottom of each block uses `SUM` formula referencing the exact cell range.

---

## Outcome

| Metric | Before | After |
|--------|--------|-------|
| Report generation | Manual filtering + copy-paste | One button, < 1 minute |
| Error risk | Manual aggregation, easy to miss rows | Automated sum + Sheets formula verification |
| HTPay summary for Finance | Manual monthly reconciliation | Auto-generated 總計 sheet |
| BestPay missing-day handling | Gaps in report if no transactions | Full calendar always shown, zeros explicit |
| INR GST calculation | Manual ÷1.04 per cell | Auto Net Subtotal column |
| Agent fee rate updates | Code changes | Edit 代理名單 sheet only |
| Platform coverage | One-off scripts per platform | Standardized flow for both HTPay and BestPay |

---

## Key Design Decisions

**1. Chinese-name filter only for agent tabs, not for 總計**
Internal accounts use Chinese agent names. Excluding them from individual tabs keeps agent reports clean, but the `總計` sheet reads all data so Finance can always reconcile 100% of volume. If the filter applied everywhere, Finance figures would be understated.

**2. BestPay: pivot as input, not raw data**
BestPay's source data arrives as a pivot table (merchant columns × date rows). Rather than re-pivoting raw data, the script reads the pivot directly and maps merchant columns to agents using the `代理名單` lookup. This avoids double-aggregation and respects the source format.

**3. Full calendar with zero-fill**
Outputting only days with transactions would make it easy to miss gaps or errors. Showing the full month forces explicit zeros — agents and Finance can immediately spot unusual zero-transaction days without additional audit steps.

**4. Rate stored as percent value (0.2 = 0.2%), not decimal**
Keeping rates in human-readable percent form in the sheet (matching how Finance expresses rates verbally) avoids the confusion of 0.002 vs. 0.2. The code divides by 100 internally.

**5. Graceful degradation on missing pivot headers**
If one platform's data hasn't been uploaded yet, a confirm dialog lets the operator continue with just the available side. This handles the real-world case where Payin is ready but Payout is still being processed.

---

## Update Log

| Date | Update |
|------|--------|
| 2026-06-29 | System documented, both platform scripts recorded |

---

*Full case study with screenshots → [Notion Portfolio](https://notion.so) (link TBD)*
