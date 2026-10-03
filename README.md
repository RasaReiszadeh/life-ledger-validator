# Life Ledger Validator

A personal-finance data-quality project that combines **Excel/VBA validation** with an independent **Python/SQLite reconciliation pipeline**.

All data included in this project is fictional and intended for demonstration purposes.

## Overview

**Life Ledger Validator** is designed to catch common data-quality problems in personal-finance records and independently verify the results using two separate validation systems.

The project includes:

* An Excel workbook with a VBA validation macro
* A Python/SQLite reconciliation script
* Automated transaction-level validation
* Workbook-to-database reconciliation
* Category-level comparison against the Summary sheet
* Duplicate detection
* A visual financial dashboard
* An accessibility-informed workbook design
* A CSV reconciliation report for independent review

The goal is not to perform a formal financial audit. This is a **self-review and data-quality project** demonstrating practical validation, reconciliation, automation, and accessibility principles.

---

## Features

### Excel / VBA Validation

The `ValidateLedger` macro checks each transaction for:

* Missing or invalid dates
* Future-dated transactions
* Unknown categories
* Invalid transaction types
* Text-formatted amounts
* Negative amounts
* Out-of-range values
* Possible duplicate transactions
* Other data-quality issues

Each transaction receives a clear validation status using both **text and colour**, so status is not communicated through colour alone.

The macro also generates a plain-language `Report` sheet summarizing the findings.

---

### Financial Reconciliation

The workbook independently checks:

* Total income
* Total expenses
* Net cash flow
* Category totals
* Manually reported net amount

Calculated results are compared against the corresponding values in the workbook to identify discrepancies.

---

### Python / SQLite Validation

`reconcile.py` provides a second, independent validation layer.

The script:

1. Loads the Excel transaction data.
2. Creates a SQLite database.
3. Recalculates transaction totals using SQL.
4. Repeats key validation checks.
5. Calculates category totals.
6. Compares calculated values with the workbook's Summary sheet.
7. Checks the manually reported net amount.
8. Generates `reconciliation_report.csv`.

Using a separate Python/SQL implementation reduces reliance on the Excel/VBA calculations alone.

---

## Workbook Structure

| Sheet            | Purpose                                                                 |
| ---------------- | ----------------------------------------------------------------------- |
| **Dashboard**    | Visual overview of income, expenses, net cash flow, and spending limits |
| **Transactions** | Source transaction data and validation results                          |
| **Summary**      | Financial totals and category breakdowns                                |
| **Report**       | Plain-language validation findings                                      |
| **Checklist**    | Accessibility and quality-control self-review                           |

The workbook uses descriptive sheet names, frozen headers, consistent formatting, and clear status labels.

---

## Accessibility-Informed Design

The workbook follows several practical accessibility principles:

* Descriptive sheet names
* Frozen header rows
* Clear table structures
* Status communicated through **text as well as colour**
* Descriptive chart titles
* Text equivalents for visual information
* Consistent formatting
* A dedicated accessibility and quality checklist

This is an **accessibility-informed self-review**, not a formal accessibility audit.

---

## Project Files

```text
Life_Ledger_Validator/
│
├── Life_Ledger_Validator.xlsm
├── validator.bas
├── make_ledger.py
├── reconcile.py
└── reconciliation_report.csv
```

### `Life_Ledger_Validator.xlsm`

The main Excel workbook containing the financial data, dashboard, summaries, reports, and VBA validation macro.

### `validator.bas`

The VBA module containing the `ValidateLedger` macro.

### `make_ledger.py`

Python script used to generate the sample workbook and its fictional financial data.

### `reconcile.py`

Independent Python/SQLite validation and reconciliation script.

### `reconciliation_report.csv`

Machine-readable output containing the results of the Python/SQL reconciliation process.

---

## How to Run

### 1. Run the Excel Validator

Open:

```text
Life_Ledger_Validator.xlsm
```

Enable macros when prompted.

Then:

1. Press **Alt + F8**
2. Select `ValidateLedger`
3. Click **Run**
4. Review the generated `Report` sheet

---

### 2. Run the Python Reconciliation

Install the required dependency:

```bash
pip install openpyxl
```

Then run:

```bash
python reconcile.py
```

The script will independently validate the workbook data and generate:

```text
reconciliation_report.csv
```

---

## Validation Workflow

```text
Excel Transactions
        │
        ▼
   VBA Validation
        │
        ├── Transaction checks
        ├── Duplicate detection
        └── Validation Report
        │
        ▼
   Workbook Summary
        │
        │
        ▼
 Python / SQLite
        │
        ├── SQL validation
        ├── Recalculated totals
        ├── Category reconciliation
        └── Net reconciliation
        │
        ▼
reconciliation_report.csv
```

The two validation paths are intentionally independent so that the same calculations are not simply repeated using the same implementation.

---

## Example Validation Checks

| Check                   | Description                                                 |
| ----------------------- | ----------------------------------------------------------- |
| Date validation         | Detects missing or invalid transaction dates                |
| Future dates            | Flags transactions dated after the current date             |
| Category validation     | Identifies categories outside the approved list             |
| Type validation         | Checks whether transaction types are valid                  |
| Amount format           | Detects text-formatted monetary values                      |
| Negative amounts        | Flags unexpected negative values                            |
| Range validation        | Identifies unusually large or invalid amounts               |
| Duplicate detection     | Identifies potentially duplicated transactions              |
| Total reconciliation    | Compares calculated totals with reported totals             |
| Category reconciliation | Compares category totals with the Summary sheet             |
| Net reconciliation      | Compares recalculated net cash flow with the reported value |

---

## Dashboard

The dashboard provides a quick visual overview of the ledger, including:

* Total income
* Total expenses
* Net cash flow
* Spending by category
* Spending versus limits
* Income versus expenses
* Validation status

The dashboard is intended for **quick review**, while the Transactions and Report sheets provide the detailed information needed to investigate individual issues.

---

## Design Philosophy

The project intentionally combines **data validation, automation, reconciliation, and presentation** rather than treating the spreadsheet as a simple financial tracker.

The design aims to be:

* **Clear** — users should quickly understand what the workbook is showing.
* **Consistent** — terminology and validation rules should match across Excel and Python.
* **Auditable** — calculations should be independently reproducible.
* **Accessible** — important information should not depend on colour alone.
* **Practical** — validation results should be understandable to a non-technical user.

---

## Limitations

This project is a demonstration of data-quality practices and is **not a financial audit**.

The data is fictional, and validation rules are intentionally designed for this sample ledger. Real-world financial applications would require additional controls, security considerations, accounting rules, testing, and domain-specific requirements.

Possible duplicate transactions are also treated as **potential matches**, not confirmed duplicates. A human review may be required before deciding whether a transaction should actually be removed.

---

## Technologies

* **Microsoft Excel**
* **VBA**
* **Python**
* **SQLite**
* **OpenPyXL**
* **SQL**
* **CSV**

---

## Project Goal

The project demonstrates how a personal-finance spreadsheet can be transformed from a basic data-entry tool into a small **data-quality and reconciliation system** with:

> **Human-readable validation + automated checks + independent verification + accessible presentation.**

All financial data in this project is fictional.
