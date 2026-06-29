# Case 05 — PSP Monthly Score & Water Level Analysis System

> **Role**: Payment Operations Analyst @ BTSE (Bestpay)
> **Tools**: Google Apps Script · Google Sheets · Google Drive API
> **Timeline**: Monthly use; run once per month to review and reset water levels
> **Scale**: 8+ currencies · 10+ PSPs · spans multiple months of daily PSP data
> **GitHub**: [case05_psp_watermark_analysis.gs](./code/case05_psp_watermark_analysis.gs)
> **Related**: → Extends [Case 03 — Daily OPS Data Analysis System](./case03-daily-ops-data-analysis.md) (source of daily PSP scores)

---

## Problem

PSP performance was tracked daily (Case 03), but water level thresholds were set manually and never updated systematically.

Before this system:
- No standard process to review PSP performance across a date range
- Water levels (routing thresholds) were set ad-hoc, not data-driven
- No way to identify which PSPs were "core" (consistently active) vs. occasional
- Transaction volume estimates were done manually by referencing scattered files

This meant water levels drifted out of sync with actual PSP capacity, causing over-routing to weak PSPs and under-utilization of strong ones.

---

## System Architecture

```
[Monthly Sheets: YYYY-MM]
  ├── Row per: Currency × PSP
  ├── Cols E+: Daily PSP scores (from Case 03 pipeline)
  └── onOpen() → auto-calculate col C (avg score) + col D (grade)
                → redraw currency-group border lines

[Data Sheet: user sets A2=start date, A3=end date, A4=days]
       ↓ [Button: Run Analysis]
[executeIndependentWatermarkAnalysis()]
  ├── Scan YYYY-MM sheets → find PSPs active > 1/3 of range days (core merchants)
  ├── fetchCurrencySuccessVolumeAdvanced()
  │     └── Open Drive folders: YYYY-MM/YYYY_MM PSP分析
  │         → 地毯式掃描: find 幣別: markers → read 24h hourly rows → sum 成功金額
  └── Write to Data sheet:
        F:I  → Currency list + auto volume + core PSP count + avg volume formula
        K:N  → Core PSP detail: currency / PSP name / avg score / grade
        O:R  → Water level formulas (raw): 100% / 80% / 40% / 15%
        T:W  → Rounded water levels (significant-figure rounding, light-blue highlight)
```

---

## What I Built

### Module 1 — Monthly Sheet Auto-Formatting (`onOpen`)

Triggers on every spreadsheet open or refresh. Scans all tabs; processes any tab named `YYYY-MM` (length 7, contains `-`).

**Score calculation** (`calculateScoresAndGrades`):
- Reads all daily score columns (E onwards)
- Calculates average score per PSP row, ignoring blanks
- Writes to column C: `avg` rounded to 1 decimal
- Writes to column D: grade based on thresholds

| Score | Grade |
|-------|-------|
| > 90  | A+    |
| > 80  | A     |
| > 70  | B+    |
| > 60  | B     |
| > 50  | C+    |
| > 40  | C     |
| > 20  | D+    |
| ≤ 20  | D     |

**Border rendering** (`autoRenderCurrencyBorders`):
- Resets entire data range to uniform light-grey thin borders first (clears stale borders from previous runs)
- Reads column A for currency values; draws a medium-weight bottom border on each currency's last row
- Solves the "stale border accumulation" problem that occurred when row counts changed between refreshes

### Module 2 — Core Merchant Identification

For a given date range:
- Counts how many days each PSP had a score > 0 (active days)
- Threshold: `total days in range / 3`
- PSPs exceeding the threshold are classified as **core merchants**
- Core merchant list groups by currency and feeds into the water level calculation

### Module 3 — Volume Fetching (`fetchCurrencySuccessVolumeAdvanced`)

Navigates the Drive folder hierarchy:
```
Parent Folder (2026 root)
  └── YYYY-MM/
        └── YYYY_MM PSP分析 (Spreadsheet)
              └── YYYY-MM-DD-代收分析 sheets (one per day)
```

**地毯式掃描 (full-grid scan)** for each daily sheet:
- Scans every cell in the sheet looking for cells starting with `幣別:`
- Once found, reads the sub-header row below to locate the `成功金額` column within that currency block
- Reads 24 rows downward (hourly data) and sums them
- Accumulates by currency across all days in range

This approach handles any layout — horizontally expanded, vertically stacked, or merged cells — without relying on fixed column positions.

### Module 4 — Water Level Formula Engine

For each core PSP, calculates four water level thresholds based on:

**Base volume per PSP** = avg daily volume for currency ÷ core PSP count

**Grade multipliers** (relative to C+ = 1.0 baseline):

| Grade | Multiplier | Rationale |
|-------|-----------|-----------|
| A+    | 2.0736    | 1.2^6 — very high capacity |
| A     | 1.728     | 1.2^5 |
| B+    | 1.44      | 1.2^4 |
| B     | 1.2       | 1.2^3 |
| C+    | 1.0       | Baseline |
| C     | 0.8333    | ~1/1.2 |
| D+    | 0.6944    | ~1/1.44 |
| D     | 0.5787    | ~1/1.728 |

**Four water levels** per PSP:
- `HIGH_CRITICAL` (100%): maximum expected volume
- `HIGH_WARNING` (80%): alert threshold
- `LOW_WARNING` (40%): under-utilization alert
- `LOW_CRITICAL` (15%): near-zero routing alert

**Significant-figure rounding** for T:W output:
```
=ROUND(x, 1 - INT(LOG10(ABS(x))))
```
Produces clean numbers (e.g. 48,230 → 50,000; 3,817 → 4,000) suitable for direct entry into routing config.

---

## Outcome

| Metric | Before | After |
|--------|--------|-------|
| Water level review | Ad-hoc, no schedule | Once/month, data-driven |
| Core PSP identification | Manual judgment | Auto: active > 1/3 of days |
| Volume estimation | Manual file lookup | Auto deep-scan across Drive folders |
| Water level values | Round numbers with no basis | Grade-weighted, significant-figure rounded |
| Border formatting | Accumulated stale borders | Full reset + redraw on every open |
| Score/grade updates | Manual formula maintenance | Auto-recalculated on open |

---

## Key Design Decisions

**1. Full-grid scan instead of fixed column indexing**
PSP analysis sheets have variable layout (some currencies have more PSPs → columns shift). Instead of assuming column positions, the scanner searches every cell for `幣別:` markers and re-locates `成功金額` within each block. This makes the system resilient to layout changes.

**2. Grade multiplier as geometric series (1.2x steps)**
The 1.2 ratio between adjacent grades was chosen to match the score gap between grade thresholds (10 points each). A+ PSPs are expected to handle roughly 4x the volume of D-grade PSPs, which matches observed real-world routing behavior.

**3. Core merchant threshold = total days / 3**
Requiring activity on >1/3 of days filters out PSPs that ran briefly as tests or during emergencies. Monthly periods (30 days) require 10+ active days; 3-month ranges require 30+ active days.

**4. Two-column output (O:R raw + T:W rounded)**
Raw values are kept for auditability; rounded values in T:W (with light-blue background) are what operators copy into the routing config. Separating them prevents rounding from hiding calculation errors.

**5. Border reset before redraw**
Calling `setBorder` with all thin-grey first ensures no ghost borders persist when row counts change between months. Without this, deleted rows would leave invisible heavy borders at wrong positions.

---

## Update Log

| Date | Update |
|------|--------|
| 2026-06-29 | System documented, code recorded |

---

*Full case study with screenshots → [Notion Portfolio](https://notion.so) (link TBD)*
*→ Related: Case 03 daily pipeline feeds the monthly score data consumed by this system*
