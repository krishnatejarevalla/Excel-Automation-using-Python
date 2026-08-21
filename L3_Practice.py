from openpyxl import load_workbook
from openpyxl.utils import get_column_interval
wb = load_workbook("MarkList.xlsx")
ws = wb.active

#Write code to find the columns between B and F
print(get_column_interval("B","F"))
        #Output:
        # ['B', 'C', 'D', 'E', 'F']

#Printing column names btw C and F from column letters
columns = get_column_interval("C","F")
for column in columns:
    print(ws[column+"1"].value)
    #Output:
    # Python
    # SQL
    # Excel
    # PowerBI


