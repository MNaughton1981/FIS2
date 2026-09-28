# FIS2 — Family Income Statement / FIRE Tracker

A Google Sheets–based personal finance system: monthly income statement, variable-spending
categorization, net-worth tracking, solar/asset ROI, and a FIRE (Financial Independence,
Retire Early) planning dashboard. This repo is the **archive** for the workbook snapshots and
the Apps Script code that powers them.

---

## ⚠️ Mental model: repo vs. live sheet

**This repo and your live Google Sheet are not connected.** Nothing here syncs automatically.

| | GitHub repo (this) | Live Google Sheet |
|---|---|---|
| What it is | An archive of `.xlsx` snapshots + `.gs` source | Where you actually work |
| How data gets in | You **download** the sheet and **commit** the `.xlsx` | You edit it directly |
| How code gets in | You **commit** the `.gs` files here | You **paste** the `.gs` into Extensions → Apps Script |
| Merging a PR here | Updates the repo only | **Does nothing to the live sheet** |

So: a merged PR never edits your sheet, never runs code, and never refreshes an `.xlsx`
snapshot. Downloading a sheet and committing it is the only thing that updates the snapshots.

---

## Files

### Workbook snapshots (`.xlsx`)
| File | What it is |
|---|---|
| `Family_Income_Statement_2026_v5_FIRE.xlsx` | Current personal workbook (19 tabs, incl. FIRE Dashboard, Pension, HSA, Scenarios, Spending Guardrail). **Contains personal data.** |
| `Family_Income_Statement_2026_v4.xlsx`, `… v4 (1).xlsx` | Prior-version snapshots (12 tabs). Kept for history. |
| `FIRE_Tracker_Template_v5.xlsx` | **Shareable, no-personal-data** blank template (11 tabs). Safe to give to others. |

### Apps Script (`apps_script/`)
| File | Belongs to | What it does |
|---|---|---|
| `Categorize.gs` | **Personal workbook** | Reads the `Imports` tabs, applies the `Rules`, and rolls spending up into `CC Detail` + the Income Statement. Resolves rows **by label** (immune to layout shifts) and treats Eversource as manual (see below). Owns the single `onOpen()` / `📊 Family Finances` menu. |
| `MonthlyEntry.gs` | **Personal workbook** | Guided monthly entry: fans the `Monthly Entry` staging tab out to Balances / Solar / Amazon / Comp & Benefits / Spending Guardrail. **Delete its `onOpen()`** so only `Categorize.gs` builds the menu. |
| `FIRE_Template.gs` | **Shareable template** | Self-contained generator: run **📊 FIRE Tracker → Setup / Rebuild Template** in a blank sheet and it builds all 11 template tabs from scratch. Zero personal data. |

**Rule of thumb:** `Categorize.gs` + `MonthlyEntry.gs` power *your* sheet; `FIRE_Template.gs`
is only for spinning up a fresh copy for someone else.

---

## Utilities are entered by hand (on purpose)

Gas and electric both bill as `EVERSOURCE WEB_PAY`, arrive on drifting dates, and can't be
split reliably from the bank feed. So the **Electric** and **Gas** rows on the Income
Statement are manual-entry, and `Categorize.gs` intercepts `EVERSOURCE` as `__MANUAL__` — it
never auto-writes those rows, so your typed bills survive every import. The **Utility Recon**
tab compares what you typed vs. what actually left the account and flags gaps (trust the YTD
flag — monthly gaps are usually just billing-date drift).

---

## Monthly workflow (personal sheet)

1. Export transactions from each account → paste into `Imports - Checking` / `Savings` / `Barclay` (data starts row 5).
2. **📊 Family Finances → Categorize & Roll Up.** Yellow rows in Imports = uncategorized; add a `Rules` entry if needed.
3. Enter this month's utility bills by hand on the Income Statement; enter balances via the `Monthly Entry` tab → **Post Monthly Entry**.
4. Split the month's Amazon lump on the **Amazon Reconciliation** tab.
5. **Generate / Refresh Charts.**
6. When you want to archive: **File → Download → .xlsx**, then commit it here.

---

## Upgrading the live sheet (e.g. V4 → V5) without re-sharing

Sharing lives on the **file**, not its contents. To keep collaborators' access, **edit the
file they're already on — never create a new one.**

- ✅ **Do:** add/edit tabs, paste updated `.gs`, or **File → Import → Replace spreadsheet**
  (imports new contents into the *same* file — same URL, same sharing).
- ❌ **Don't:** *File → Make a copy*, or upload an `.xlsx` as *Create new spreadsheet* — both
  make a new file with a new URL that collaborators aren't on.

V4 → V5 was purely additive (7 new tabs: Pension, FIRE Dashboard, Amazon Reconciliation,
Solar Reinvestment, HSA Election Planner, Monthly Entry, Spending Guardrail; nothing removed).
After importing/adding tabs, update the Apps Script (paste `Categorize.gs`, delete the
`onOpen()` in `MonthlyEntry.gs`) and run **Generate / Refresh Charts**.

---

## Sharing the template with someone else

1. Give them `FIRE_Tracker_Template_v5.xlsx` (open in Sheets) **or** a blank sheet + `FIRE_Template.gs`.
2. They **File → Make a copy** (their own copy — no personal data to worry about).
3. Extensions → Apps Script → paste `FIRE_Template.gs` → run **Setup / Rebuild Template**.
4. The `.xlsx` holds structure + formulas; the menu, categorizer, and charts come from the script.

> Note: the template's Asset tab is named `Asset ROI & Reinvestment` (Excel forbids `/` in
> sheet names, which the Google-Sheets generator uses).
