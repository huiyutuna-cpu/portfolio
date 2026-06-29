# Case 07 — Crypto Wallet Transaction Auto-Matching System

> **Role**: Payment Operations Analyst @ BTSE (Bestpay)
> **Tools**: Google Apps Script · Google Sheets · Gmail API
> **Trigger**: Auto-runs on data entry / manual force-run
> **Scale**: 3 wallet types · 2 platforms (BestPay / HTPay) · used by OPS + OPS Finance + Company Finance + Settlement Team
> **GitHub**: [case07_wallet_hash_matching.gs](./code/case07_wallet_hash_matching.gs)

---

## Problem

Crypto wallet transactions (USDT充值 / USDT下發 / INR OTC) come from multiple platforms. Each row in the wallet data table needs to be matched back to its originating merchant or account — but this required manually searching hash values across multiple source sheets.

Before this system:
- Someone had to copy each transaction hash and manually look it up in BestPay or HTPay export sheets
- Three separate wallet types, each with different source combinations, meant no unified process
- Missed matches went unnoticed — no alert when a row couldn't be identified
- Finance and Settlement team had no automated way to confirm transaction origin

---

## System Architecture

```
Trigger: onEdit (auto) or manualForceRun() (button)
                ↓
executeMappingEngine(isForceAll)
  ├── Load all source sheets into memory (one read per source)
  ├── For each unprocessed row in each Wallet Data sheet:
  │     ├── [Strategy 1] Special condition (INR OTC only)
  │     │     └── Match by Amount + Date against TXID source
  │     └── [Strategy 2] Hash/UUID extraction & lookup
  │           ├── Extract 64-char hex (TxHash) or 36-char UUID from target cols
  │           ├── Search across all configured source sheets
  │           └── On match: write merchant identifier to col A
  └── Unmatched rows → col A = "未配到" (red) + send error email

3 wallet configurations:
  Bestpayops  → BestPayOps Data + HTPay USDT充值
  Bestpayotc  → BestPayOTC Data + HTPay USDT下發
  bestpayinrotc → BestPayINROTC Data + TXID source (amount+date fallback)
```

---

## What I Built

### Auto-trigger logic (`onEdit`)
- Watches all configured wallet sheets for edits
- Scans column B: if 10+ consecutive rows have B filled but A empty → auto-run
- The 10-row threshold prevents triggering on every single paste/keystroke; only fires when a meaningful batch of new data has arrived

### Matching engine (`executeMappingEngine`)

**Source caching**: all source sheets are loaded into memory once at the start. This avoids repeated Sheets API calls in the inner loop — critical for performance when matching hundreds of rows.

**Two matching strategies**:

1. **Hash/UUID matching** (primary): extracts 64-char hex (blockchain TxHash) or 36-char UUID from configured target columns using regex. Searches the same pattern in each source sheet's hash columns. Case-insensitive, hyphen-normalized for UUID comparison.

2. **Amount + Date fallback** (INR OTC only): some INR transactions don't carry a standard hash — the TXID source uses amount and date instead. Matches `Number(target.amount) === Number(source.amount)` AND `formatDate(target.date) === formatDate(source.date)`.

**Result writing**:
- Match found → write merchant identifier to col A, grey background
- Row marked `Cancel` in col P → write "Cancel", skip matching
- No match → write "未配到", red background, log the row

**Force-run mode** (`manualForceRun`): re-processes all rows regardless of col A state — used when source data is updated after initial run, or for monthly full reconciliation.

**Error email** (`sendErrorEmail`): on normal (non-force) runs, sends a summary email listing every unmatched row by sheet name and row number. Finance and Settlement team can act on these immediately without opening the spreadsheet.

---

## Outcome

| Metric | Before | After |
|--------|--------|-------|
| Transaction identification | Manual hash lookup across sheets | Auto-matched on data paste |
| Coverage | Only when someone had time | Every row, every run |
| Unmatched alert | Silent — gaps went unnoticed | Email notification, row highlighted red |
| INR OTC handling | No standard hash → manual only | Amount + date fallback matching |
| Multi-wallet support | 3 separate ad-hoc processes | Single config-driven engine |
| Re-reconciliation | Redo manually | One-click force-run |
| Processing speed | Slow — repeated Sheets reads per row | All sources loaded once; 100+ rows complete in seconds |
| Match rate | Low — manual lookup misses rows | High — hash matching + amount/date fallback covers edge cases |

---

## Key Design Decisions

**1. Config-driven, not hardcoded**
All wallet-to-source mappings are in a `CONFIG` array at the top of the script. Adding a new wallet type means adding one object — no changes to the matching engine. This was important because wallet types were added incrementally as the team onboarded new platforms.

**2. 10-row threshold for auto-trigger**
`onEdit` fires on every cell change. A threshold of 10 consecutive unprocessed rows ensures the engine only runs when a real data batch has been pasted in — not on isolated edits, formula updates, or single-row entries.

**3. Source data loaded once per run**
Each source sheet is read into memory before the matching loop. Without this, a 200-row target sheet against 3 source sheets would make 600+ individual Sheets reads — well above script timeout limits.

**4. Two-tier matching for INR**
INR OTC transactions from a specific counterparty don't have a standard hash. Rather than treating these as always-unmatched, a secondary condition (amount + date, scoped by account name) catches them before falling through to the "未配到" path.

**5. Force-run vs. incremental**
Normal runs skip rows where col A is already filled (incremental update). Force-run re-evaluates all rows — needed when a source sheet is refreshed mid-month and previously unmatched rows may now have a match.

---

## Update Log

| Date | Update |
|------|--------|
| 2026-06-29 | System documented |

---

*Full case study → [Notion Portfolio](https://notion.so) (link TBD)*
