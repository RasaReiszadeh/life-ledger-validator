# Life Ledger Validator

A personal-finance data-quality project: an Excel workbook with a VBA validation macro, plus a Python/SQLite script that independently reconciles the same data. All data is fictional.

## What it does
- **VBA macro (`ValidateLedger`)** checks every transaction for missing/invalid dates, future dates, unknown categories, wrong types, text-formatted or negative amounts, out-of-range values, and possible duplicates. It labels each row (text and colour) and writes a plain-language `Report` sheet.
- **Reconciliation** cross-checks workbook totals and a manually reported net against recalculated figures.
- **Python/SQL script (`reconcile.py`)** loads the workbook into SQLite and repeats the checks with SQL, compares category totals to the Summary sheet, and exports `reconciliation_report.csv`.
- **Dashboard** shows spending vs limit and income vs expenses.
- **Accessibility-informed design**: descriptive sheet names, frozen headers, no colour-only status, chart titles with text equivalents, and a checklist sheet. This is a self-review, not a formal audit.

## Files
- `Life_Ledger_Validator.xlsm` - workbook with the macro
- `validator.bas` - the VBA module (import into the workbook)
- `make_ledger.py` - generates the sample workbook
- `reconcile.py` - SQL reconciliation

## Run it
1. Open the workbook, enable macros, press Alt+F8, run `ValidateLedger`.
2. `pip install openpyxl`, then `python reconcile.py`.

## Screenshots
![Report sheet](screenshots/report.png)
![Dashboard](screenshots/dashboard.png)
![Reconciliation output](screenshots/reconcile.png)
