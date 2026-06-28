# Case 04 — Merchant Account Opening System

> **Role**: Payment Operations Analyst @ Nogle
> **Tools**: Google Apps Script · Google Sheets · HTML/CSS/JS · Telegram Bot API
> **Timeline**: Built as single-sprint feature; live and used by AM/CS teams
> **Scale**: 30+ merchants · 8 fiat + 5 crypto currencies · Daily use by AM team
> **GitHub**: [case04_account_opening_form.gs](./code/case04_account_opening_form.gs) · [case04_account_opening_form.html](./code/case04_account_opening_form.html) · [case04_telegram_bot_account_opening.gs](./code/case04_telegram_bot_account_opening.gs)

---

## Problem

Merchant onboarding had no standardized process.

Before this system:
- AM had to coordinate with OPS manually to fill in new merchant info
- Station ID was generated ad-hoc, no naming convention enforced
- CS had no structured checklist when creating accounts in the backend system
- OPS had no single source of truth for active merchants and their currencies
- CS Shift Lead could not audit what had been opened or pending

This caused duplicates, inconsistent data entry, and delayed onboarding.

---

## System Architecture

```
[AM → HTML Form (Web App)]
       ↓
[Google Apps Script Backend]
  ├── Auto-generate Station ID (currency prefix + 4-digit seq)
  ├── Create merchant tab in Account Book (Spreadsheet)
  ├── Sync row to 商戶＆代理 Management Sheet (rates, status, date)
  └── Push Telegram notification → CS Group
       ↓
[CS → Telegram Bot (/openacct)]
  ├── Search merchant by name → match Account Book tab
  ├── Auto-detect: New Merchant vs Existing Merchant (Add Currency)
  ├── Step-by-step checklist (4 fixed steps + dynamic data from sheet)
  ├── API key collection (new merchant only, 64-char validation)
  └── Generate final notice (ZH-CN / EN) → Copy to CS for merchant
```

---

## What I Built

### Module 1 — HTML Registration Form (Web App)

Two modes on entry:
- **New Merchant**: auto Station ID, full contact fields
- **Existing Merchant (Add Currency)**: Station ID required, contact fields hidden

Form sections:
- Basic info: name, contact, phone, email
- Services: Pay / Payout checkboxes
- Active currencies (loaded dynamically from 商戶＆代理 via `getUniqueCurrencies()`)
- Fiat sections: Payout / Withdraw / Pay — each with quick-add buttons per currency (CNY/VND/IDR/INR/PHP/THB/BRL/MYR/BDT)
- Crypto sections: shared-value flow → fill once, auto-expand to BTC/ETH/USDT/USDC/BNB
- Agent affiliation + IP whitelist (API / Login / Withdraw)

Anti-resubmission design:
- On submit, form disappears **immediately** (before server responds)
- Success page shown; backend processes async
- "Next merchant" button calls `window.location.reload()` — hard reload, no cached state

### Module 2 — GAS Backend (`submitData`)

**Station ID generation** (`generateStationId_`):

| Currency Scenario | Prefix |
|---|---|
| Crypto only (USDT/BTC/etc.) | `CY` |
| Single fiat currency | First 2 chars (e.g. `CN`, `IN`, `PH`) |
| Multiple fiat currencies | `MA` |

Sequential number: scans 商戶＆代理 column A for existing IDs with same prefix → takes `max + 1` → zero-pads to 4 digits.

**Duplicate guard**: checks if a tab named `merchantName` already exists in the spreadsheet. If yes, aborts with message — prevents F5 / reload resubmission from writing twice.

**Sheet writing**:
- Creates new tab `merchantName` with 6 blocks: basic info → services → payout → withdraw → payin → agent/IP
- Simultaneously writes one row per currency to 商戶＆代理 management sheet (Station ID, name, currency, date, agent name, agent rate %, service rate %, payout rate %, status = `待開戶`)

**Telegram notification** to CS group:
```
CY0031 - MerchantName is now ready to create account, please proceed via HTpay_Selfservice bot.

⚠️ Reminder: This is production account info. If the merchant needs to test, please ensure a test account is provided.
```

### Module 3 — Telegram Bot (CS Account Opening Flow)

**State machine** stored in `PropertiesService.getUserProperties()`:
- Steps: `ask_name` → `choose_tab` → `confirming` → `await_api_key` → `notify_lang`
- State persists across messages; cleared after completion or cancel

**Auto-detection of merchant type**:
- Reads `B3` of account sheet; if value contains `"舊商戶無須填寫"` → `isOldMerchant = true`
- Changes: skip API key step, different password hint (`舊商戶 無需填寫`), different completion message

**Fixed 4 steps** (prepended before dynamic checklist):
1. Login Account: `merchantName`
2. Password: `12345` (or `舊商戶 無需填寫`)
3. Confirmed password: same
4. API key: prompt for 64-char input (new) or auto-skip (existing)

**Dynamic checklist** built from account sheet:
- A1–A8 (basic info) from column B
- A9 merged B:G (active currencies)
- Per-currency columns: Payout block / Withdraw block
- Per-column payin blocks
- Misc rows A29+

**Completion flow**:
1. Sends group notification: `{ID} - {Name} Merchant account is now created, please check for assignment and routing.`
2. Writes today's date to column V of all matching rows in 商戶＆代理
3. Prompts CS to choose notification language: `ZH-CN` or `EN`
4. Generates formatted merchant notice (different template for new vs. existing) and displays for CS to copy

---

## Outcome

| Metric | Before | After |
|--------|--------|-------|
| AM onboarding process | Manual coordination with OPS | Self-service web form, anywhere |
| Station ID format | Ad-hoc | Auto-generated, prefix + 4-digit seq |
| CS checklist | None (verbal / memory) | Step-by-step bot guide, all fields |
| Merchant registry | Manually updated spreadsheet | Auto-written on form submit |
| Duplicate opens | Possible (no guard) | Blocked at form + sheet level |
| Merchant notification | Manually typed by CS | Auto-generated ZH/EN template |
| Open date tracking | Manual | Auto-stamped to 商戶＆代理 V column |

---

## Key Design Decisions

**1. Immediate form hide on submit**
The form container is hidden the moment the user clicks Submit — before the backend responds. This prevents CS from accidentally submitting twice via F5 or back button. The server still processes in the background; failure restores the form.

**2. Station ID prefix logic**
Crypto currencies are excluded from fiat prefix calculation. A merchant with both CNY and Crypto gets prefix `CN`, not `MA`. Only merchants with 2+ distinct fiat currencies get `MA`. This keeps IDs semantically meaningful.

**3. Merchant type detection via cell content**
Rather than adding a separate field or sheet, the bot detects `isOldMerchant` by reading whether `B3` (Contact Person) contains `"舊商戶無須填寫"` — a value the frontend injects as a hidden field when the "Existing Merchant" mode is selected. No schema change needed in the account sheet.

**4. Currency list loaded dynamically**
The checkbox options for "active currencies" are not hardcoded — they're fetched live from the 商戶＆代理 sheet's column D via `getUniqueCurrencies()`. This means new currencies added to the registry automatically appear in the form without code changes.

**5. Rates stored as decimals with number format**
Agent rates and service rates are stored in 商戶＆代理 as decimals (e.g. `0.003`) with `setNumberFormat('0.00%')`. This allows Sheets formulas to use them directly in calculations without conversion.

---

## Update Log

| Date | Update |
|------|--------|
| 2026-06-29 | System documented, all 3 code files recorded |

---

*Full case study with screenshots → [Notion Portfolio](https://notion.so) (link TBD)*
