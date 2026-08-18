from openpyxl import workbook, load_workbook
wb= load_workbook("Demo.xlsx")
ws = wb.active
# ws.insert_rows(3) #inserted empty row at 3
# ws.insert_rows(5) #inserted empty row at 5
# ws.insert_cols(2) #inserted empty col at B column
# ws.delete_rows(5) #deleted row at 5
ws.delete_cols(2) #deleted col at 2
ws.delete_cols(3) #deleted col at 3
# ws.move_range("F1:G6", rows=3, cols=3) #moving cells from F1:G6 right by 3 cols and down by 3 rows
ws.move_range("F1:G6", rows=0, cols=-3) #moving cells left by 3 cols
wb.save("Demo.xlsx")