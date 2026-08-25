# =========================================
#      Student Performance Analyser
# =========================================


# STEP 1 — Loading the WorkBook
from openpyxl import load_workbook
wb = load_workbook("MarkList.xlsx")
ws = wb.active
print("===========================================================")
print("===================Excel Mini Project======================")
print("===========================================================")
print("Rows:",ws.max_row)
print("Columns:",ws.max_column)

# STEP 2 — Understand which columns are subjects
print("\nColumn Headers:")
for col in range(1, ws.max_column+1):
    print(col, ws.cell(row=1, column =col).value)
        #Output:
        # Column Headers:
        # 1 Student_ID
        # 2 Name
        # 3 Python
        # 4 SQL
        # 5 Excel
        # 6 PowerBI


