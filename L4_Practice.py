from openpyxl import load_workbook
from openpyxl.utils import range_boundaries
wb = load_workbook("MarkList.xlsx")
ws = wb.active

#Printing the boundaries of a given range
range = input("Enter the range of cells: ")
min_col, min_row, max_col, max_row = range_boundaries(range)
print(f"Starting row : {min_row}",
      f"Starting col : {min_col}",
      f"ending row :{max_row}",
      f"ending col : {max_col}"
      , end =",")
        #Output:
        #Enter the range of cells: A1:F11
        #Starting row : 1 Starting col : 1 ending row :11 ending col : 6,