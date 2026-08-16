#Automating Excel using openpyxl module
from openpyxl import Workbook, load_workbook
wb = load_workbook("Marks.xlsx")
ws = wb.active
print(ws['A2'].value) #Arjun
print(ws['B6'].value) #69
print(ws.max_row) #9 rows
print(ws.max_column) #5 cols