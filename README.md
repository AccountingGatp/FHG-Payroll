# Payroll → QuickBooks Journal Import Converter

A browser-based tool that turns a **Payworks payroll export** into a balanced,
import-ready **QuickBooks journal-entry workbook** — implementing the GATP
_"Payroll → QuickBooks Journal Import"_ SOP.

Everything runs **client-side in the browser**. Payroll data is never uploaded
to a server — the two files you drop in are parsed in memory and the finished
workbook is generated and downloaded locally. Safe for sensitive payroll data,
and it works fully offline.

**Live app:** https://fhg-payroll-qb-converter.vercel.app

> The hosted version loads the spreadsheet engine (`xlsx-js-style`) from a public
> CDN (jsDelivr, with an unpkg fallback). The copy in this repo is fully
> self-contained — it vendors the engine under `vendor/` so it runs with no
> network access at all. Both process your files entirely in the browser.

---

## What it does

Given two files:

1. **Payroll input** — the one-sheet Payworks export. Header row: a `YYYYMMDD`
   date in A1, then `Class` (B), `Code` (C), `Memo` (D), `Original …` (E),
   `Payworks Journal Export - New Month` (F), `… Old Month` (G).
2. **Master Account List** — columns `Books` (A), `Memo` (B), `AccountName` (C),
   with one row set per book (e.g. Canada / NB / ON).

…the app:

1. **Auto-detects the book** (Canada / NB / ON) and explains why — a memo that
   exists in only one book points at that book. You can override the pick.
2. **Derives the journal dates** — New Month = the A1 header date; Old Month =
   the last day of the prior month. Both are editable, and the `Journal No`
   (`Payroll` + `MMDDYYYY`) updates live.
3. **Builds two import sheets** — `New Month` (amounts from column F) and
   `Old Month` (amounts from column G), one journal line per input row:
   - `Journal No | Journal Date | Memo | Account | Debit | Credit | Description | Class`
   - Account is looked up from the confirmed book by exact Memo match.
   - Negative amounts become **Debits**; a positive non-clearing amount goes to
     **Credit**; the **Payroll Clearing Account** line is the balancing credit.
   - Zero-amount lines are dropped.
   - The clearing credit is set to `sum(debits) − sum(non-clearing credits)` so
     each journal balances to **$0.00 exactly**, absorbing any penny rounding.
   - Unmapped memos are left blank and **highlighted yellow**, and listed on the
     Summary — the app never guesses an account.
4. **Adds a Summary sheet** (first tab) — per-month journal no, date, line
   count, formula-driven Total Debit / Credit / Difference, a BALANCED status,
   the clearing rounding adjustment, and any unmapped memos.
5. **Optionally maps numeric Class IDs → QBO class names** (paste
   `ID = Name` lines). Blank classes stay blank.

Formatting follows the GATP house style: Arial, navy `#1F3864` header fill with
white bold text, alternating row fills, currency format
`#,##0.00;(#,##0.00);"-"`, and formula-driven totals (never hardcoded).

## Using it

**Offline single file (easiest)** — download **`Payroll-QB-Converter.html`** and
double-click it. Everything (the spreadsheet engine, logic, and styling) is
inlined into that one file, so it runs with no server, no internet, and no other
files. Ideal for handling payroll data on a locked-down machine — keep it on
your desktop or a shared drive and just open it.

**Locally from the repo** — or open `index.html` in any modern browser (it loads
its engine from `vendor/`, also fully offline). Double-click it, or serve the
folder:

```bash
python3 -m http.server 8000
# then visit http://localhost:8000
```

Drop in the two files, confirm the book and dates, click **Generate**. The
workbook downloads as `Payroll_QB_Import_<BOOK>_<MMDDYYYY>.xlsx`, ready to
import into QuickBooks.

**Deploying** — it's a static site (HTML/CSS/JS + one vendored library). Host it
on any static host (Vercel, Netlify, GitHub Pages, an internal share). No build
step and no backend required.

## Checklist before importing

- [ ] Confirmed the book (Canada / NB / ON)
- [ ] Confirmed the dates (New Month = header date, Old Month = prior month-end)
- [ ] Both sheets show **BALANCED** (Difference = 0.00)
- [ ] No yellow unmapped-memo cells (or each resolved)
- [ ] Reviewed the clearing rounding adjustment
- [ ] Mapped Class IDs → names if your QBO file needs names

## Project layout

```
index.html          UI
styles.css          GATP house-style theme
app.js              browser glue (file input, confirm step, download)
converter.js        pure conversion logic — parsing, book detection,
                    journal building, workbook generation (SOP rules live here)
vendor/
  xlsx.bundle.js    xlsx-js-style (SheetJS fork with cell styling) — vendored
                    so the app runs offline; the browser only needs this file
  cpexcel.js        optional codepage table (used only by Node-side tests)
```

`converter.js` is framework-agnostic and takes the XLSX library as a parameter,
so the exact shipping logic is unit-testable in Node.

## Notes

- **Class IDs vs names:** the input Class column carries numeric IDs. The app
  keeps them as-is unless you supply a name mapping.
- Dates are written as real Excel date values (so QuickBooks reads them
  correctly), displayed in short-date format.
