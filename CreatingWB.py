#Creating a new Excel workbook and sheets in it and inserting cellvalues
from openpyxl import Workbook, load_workbook
wb = Workbook()
ws = wb.active
ws.title = "XYZ Company"
ws.append(['Hello','Welcome','To','Our','Organisation'])
ws.append(['Hello','Welcome','To','Our','Organisation'])
ws.append(['Hello','Welcome','To','Our','Organisation'])
ws.append(['Hello','Welcome','To','Our','Organisation'])
ws.append(['Hello','Welcome','To','Our','Organisation'])
ws.append(['Hello','Welcome','To','Our','Organisation'])
wb.save('Demo.xlsx')
