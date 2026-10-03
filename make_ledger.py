import datetime as dt
from openpyxl import Workbook
from openpyxl.chart import BarChart, Reference
from openpyxl.formatting.rule import FormulaRule
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.worksheet.datavalidation import DataValidation

NAVY = "1F3864"
HEAD_FONT = Font(bold=True, color="FFFFFF")
HEAD_FILL = PatternFill("solid", fgColor=NAVY)


def style_header(ws, row, ncols):
    for c in range(1, ncols + 1):
        cell = ws.cell(row, c)
        cell.font = HEAD_FONT
        cell.fill = HEAD_FILL


wb = Workbook()
dash = wb.active
dash.title = "Dashboard"
tx = wb.create_sheet("Transactions")
bud = wb.create_sheet("Budget")
rules = wb.create_sheet("Rules")
summ = wb.create_sheet("Summary")
acc = wb.create_sheet("Accessibility")

# ---------------- Transactions (fictional data with deliberate errors) ----------------
tx.append(["Date", "Description", "Category", "Type", "Amount", "Status"])
D = dt.date
data = [
    (D(2026, 9, 1), "Paycheque", "Income", "Income", 2400),
    (D(2026, 9, 1), "Rent", "Housing", "Expense", 1450),
    (D(2026, 9, 2), "Grocery run", "Groceries", "Expense", 86.40),
    (D(2026, 9, 3), "PRESTO reload", "Transport", "Expense", 60),
    (D(2026, 9, 4), "Netflix", "Subscriptions", "Expense", 16.49),
    (D(2026, 9, 5), "Lunch with friends", "Dining", "Expense", 42.75),
    (D(2026, 9, 7), "Grocery run", "Groceries", "Expense", 92.10),
    (D(2026, 9, 8), "Python course", "Education", "Expense", 39.99),
    (D(2026, 9, 10), "Ramen night", "Dining", "Expense", 38),
    (D(2026, 9, 12), "Grocery run", "Groceries", "Expense", 101.25),
    (D(2026, 9, 14), "Movie night", "Entertainment", "Expense", 27.50),
    (D(2026, 9, 15), "Paycheque", "Income", "Income", 2400),
    (D(2026, 9, 16), "PRESTO reload", "Transport", "Expense", 60),
    (D(2026, 9, 18), "Sushi", "Dining", "Expense", 68),
    (D(2026, 9, 19), "Grocery run", "Groceries", "Expense", 88.90),
    (D(2026, 9, 21), "Textbook", "Education", "Expense", 150),
    (D(2026, 9, 22), "Concert tickets", "Entertainment", "Expense", 145),
    (D(2026, 9, 25), "Spotify", "Subscriptions", "Expense", 11.99),
    # ---- deliberate problems below ----
    (D(2026, 9, 25), "Spotify", "Subscriptions", "Expense", 11.99),   # possible duplicate
    ("2026-09-31", "Grocery run", "Groceries", "Expense", 74.20),     # date stored as text
    (D(2026, 9, 26), "Bookstore", "Educaton", "Expense", 33.00),      # category typo
    (D(2026, 9, 27), "Mystery charge", None, "Expense", 19.99),       # missing category
    (D(2026, 9, 27), "Coffee", "Dining", "Expense", "12.50 CAD"),     # amount as text
    (D(2026, 9, 28), "Refund", "Groceries", "Expense", -30),          # negative amount
    (D(2026, 9, 28), "Laptop", "Education", "Expense", 99999),        # over maximum
    (D(2027, 1, 5), "Pre-paid gym", "Entertainment", "Expense", 45),  # future date
    (D(2026, 9, 29), "Side gig", "Income", None, 120),                # missing type
]
for row in data:
    tx.append(list(row))
style_header(tx, 1, 6)
for col, w in zip("ABCDEF", [14, 26, 18, 12, 14, 14]):
    tx.column_dimensions[col].width = w
for r in range(2, tx.max_row + 1):
    tx.cell(r, 1).number_format = "yyyy-mm-dd"
    tx.cell(r, 5).number_format = "#,##0.00"
tx.freeze_panes = "A2"
dv_type = DataValidation(type="list", formula1='"Income,Expense"', allow_blank=True)
tx.add_data_validation(dv_type)
dv_type.add("D2:D500")

# ---------------- Budget ----------------
bud.append(["Category", "Monthly Limit"])
for cat, lim in [("Housing", 1500), ("Groceries", 400), ("Transport", 180),
                 ("Dining", 150), ("Subscriptions", 60), ("Education", 200),
                 ("Entertainment", 100)]:
    bud.append([cat, lim])
style_header(bud, 1, 2)
bud.column_dimensions["A"].width = 20
bud.column_dimensions["B"].width = 16
for r in range(2, bud.max_row + 1):
    bud.cell(r, 2).number_format = "#,##0.00"

# ---------------- Rules ----------------
rules.append(["Rule", "Value"])
rules.append(["Maximum single transaction", 5000])
rules.append(["Reconciliation tolerance", 0.01])
style_header(rules, 1, 2)
rules.column_dimensions["A"].width = 30
rules.column_dimensions["B"].width = 14
rules["A5"] = "Edit only the Value column. The VBA macro reads these cells."

# ---------------- Summary ----------------
summ["A1"], summ["B1"] = "Total income", '=SUMIFS(Transactions!E:E,Transactions!D:D,"Income")'
summ["A2"], summ["B2"] = "Total expenses", '=SUMIFS(Transactions!E:E,Transactions!D:D,"Expense")'
summ["A3"], summ["B3"] = "Calculated net", "=B1-B2"
summ["A5"], summ["B5"] = "Reported net (typed in from the app)", 2310
for r in (1, 2, 3, 5):
    summ.cell(r, 1).font = Font(bold=True)
    summ.cell(r, 2).number_format = "#,##0.00"
summ["A6"] = "B5 is a fictional hardcoded input used to test reconciliation."
summ["A8"], summ["B8"], summ["C8"], summ["D8"], summ["E8"] = (
    "Category", "Spent", "Limit", "Remaining", "Status")
style_header(summ, 8, 5)
for i in range(7):
    r = 9 + i
    summ.cell(r, 1).value = f"=Budget!A{i + 2}"
    summ.cell(r, 2).value = (f'=SUMIFS(Transactions!$E:$E,Transactions!$C:$C,A{r},'
                             f'Transactions!$D:$D,"Expense")')
    summ.cell(r, 3).value = f"=Budget!B{i + 2}"
    summ.cell(r, 4).value = f"=C{r}-B{r}"
    summ.cell(r, 5).value = f'=IF(D{r}<0,"OVER BUDGET","Within budget")'
    for c in (2, 3, 4):
        summ.cell(r, c).number_format = "#,##0.00"
summ.conditional_formatting.add(
    "E9:E15",
    FormulaRule(formula=['$E9="OVER BUDGET"'],
                fill=PatternFill(start_color="FFC7CE", end_color="FFC7CE", fill_type="solid"),
                font=Font(bold=True)))
summ.column_dimensions["A"].width = 38
for col in "BCD":
    summ.column_dimensions[col].width = 14
summ.column_dimensions["E"].width = 16

# ---------------- Dashboard ----------------
dash["A1"] = "Life Ledger Dashboard - September 2026 (fictional data)"
dash["A1"].font = Font(bold=True, size=14)
labels = [("Total income", "=Summary!B1"), ("Total expenses", "=Summary!B2"),
          ("Calculated net", "=Summary!B3"),
          ("Categories over budget", '=COUNTIF(Summary!E9:E15,"OVER BUDGET")')]
for i, (lab, f) in enumerate(labels):
    dash.cell(3 + i, 1).value = lab
    dash.cell(3 + i, 1).font = Font(bold=True)
    dash.cell(3 + i, 2).value = f
    dash.cell(3 + i, 2).number_format = "#,##0.00"
dash.cell(6, 2).number_format = "0"
dash.column_dimensions["A"].width = 26
dash.column_dimensions["B"].width = 16

c1 = BarChart()
c1.type = "col"
c1.title = "Spending vs monthly limit by category"
c1.x_axis.title = "Category"
c1.y_axis.title = "Amount (CAD)"
c1.add_data(Reference(summ, min_col=2, max_col=3, min_row=8, max_row=15), titles_from_data=True)
c1.set_categories(Reference(summ, min_col=1, min_row=9, max_row=15))
c1.x_axis.delete = False
c1.y_axis.delete = False
c1.width, c1.height = 20, 9
dash.add_chart(c1, "A9")

c2 = BarChart()
c2.type = "col"
c2.title = "Total income vs total expenses"
c2.y_axis.title = "Amount (CAD)"
c2.add_data(Reference(summ, min_col=2, min_row=1, max_row=2), titles_from_data=False)
c2.set_categories(Reference(summ, min_col=1, min_row=1, max_row=2))
c2.legend = None
c2.x_axis.delete = False
c2.y_axis.delete = False
c2.width, c2.height = 14, 8
dash.add_chart(c2, "A29")

# ---------------- Accessibility checklist ----------------
acc.append(["Check", "How this workbook handles it", "Verified?"])
items = [
    ("Descriptive sheet names", "Sheets are named Dashboard, Transactions, Budget, Rules, Summary, Report, Accessibility."),
    ("Marked and frozen header rows", "Every data sheet has a bold header row; Transactions freezes it."),
    ("No merged cells or empty rows inside data", "Tables are plain grids with no merged cells."),
    ("Status is not shown by colour alone", "Row status and budget status are also written as text (OK / Check / OVER BUDGET)."),
    ("Text and background contrast", "Headers use white text on dark navy (1F3864)."),
    ("Charts have titles, axis titles and a text equivalent", "Both charts have titles and axis titles; the same figures appear as text on Dashboard and Summary."),
    ("Plain-language messages", "Validation report messages are full sentences that say what to fix."),
    ("Alt text on both charts", "MANUAL: right-click each chart > Edit Alt Text and describe the takeaway."),
    ("Run Excel's Accessibility Checker", "MANUAL: Review > Check Accessibility, then fix every warning."),
    ("Document title set", "MANUAL: File > Info > Properties > Title."),
]
for chk, how in items:
    acc.append([chk, how, "No"])
style_header(acc, 1, 3)
acc.column_dimensions["A"].width = 44
acc.column_dimensions["B"].width = 80
acc.column_dimensions["C"].width = 12
for r in range(2, acc.max_row + 1):
    acc.cell(r, 2).alignment = Alignment(wrap_text=True, vertical="top")
    acc.cell(r, 1).alignment = Alignment(wrap_text=True, vertical="top")
dv_ver = DataValidation(type="list", formula1='"Yes,No"', allow_blank=False)
acc.add_data_validation(dv_ver)
dv_ver.add("C2:C50")
acc.freeze_panes = "A2"

wb.save("Life_Ledger_Validator.xlsx")
print("Created Life_Ledger_Validator.xlsx")
