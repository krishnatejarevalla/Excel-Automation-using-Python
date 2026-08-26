# ============================================================
#          Student Performance Analyzer
# ============================================================
# Description:
#This Project helps the teachers to get all the progress of a specific 
#student using their Name or ID
# ============================================================


from openpyxl import load_workbook

# STEP 1 - Load the Excel Workbook

wb = load_workbook("MarkList.xlsx")
ws = wb.active
print("=" * 60)
print("           STUDENT PERFORMANCE ANALYZER")
print("=" * 60)
print("Workbook loaded successfully!")
print("Worksheet:", ws.title)

#Finding total no of rows and columns that have data
print("Rows:", ws.max_row)
print("Columns:", ws.max_column)

        # OUTPUT:
        # Rows: 11
        # Columns: 6

#Printing Column Headers for Teacher's Reference
print("\nColumn Headers:")
for col in range(1, ws.max_column + 1):
    print(col, ws.cell(row=1, column=col).value)

        # OUTPUT:
        # Column Headers:
        # 1 Student_ID
        # 2 Name
        # 3 Python
        # 4 SQL
        # 5 Excel
        # 6 PowerBI

# STEP 2 - Identify Subjects Automatically

def get_subject_columns(ws):
    #Creating an empty dictionary
    subject_columns = {}
    for col in range(1, ws.max_column + 1):
        header = ws.cell(row=1, column=col).value
        if header not in ["Student_ID", "Name"]:
            #Storing Subject names and their column nums as key-value pairs in dict
            subject_columns[header] = col

    return subject_columns
subjects = get_subject_columns(ws)

#Printing the subjects-column numbers from dictionary
print("\nSubjects:")
for subject, column in subjects.items():
    print(subject, "-> Column", column)

        # OUTPUT:
        # Subjects:
        # Python -> Column 3
        # SQL -> Column 4
        # Excel -> Column 5
        # PowerBI -> Column 6

# STEP 3 - Find Student
def find_student(ws, search_value):
    for row in range(2, ws.max_row + 1):
        student_id = ws.cell(row=row, column=1).value
        student_name = ws.cell(row=row, column=2).value
        if str(student_id) == search_value or str(student_name).lower() == search_value.lower():
            return row

    return None

search_value = input("\nEnter Student ID or Name: ").strip()
student_row = find_student(ws, search_value)

#Printing the Search Result
if student_row is None:
    print("\nStudent not found.")
else:
    print("\nStudent found!")
    print("Excel Row:", student_row)

        # OUTPUT:
        # Enter Student ID or Name: priya
        # Student found!
        # Excel Row: 3
        # Enter Student ID or Name: Krishna
        # Student not found.
    
    