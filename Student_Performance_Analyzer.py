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


# STEP 4 - Get Student Subject Marks
def get_student_marks(ws, student_row, subjects):
    student_marks = {}
    for subject, column in subjects.items():
        marks = ws.cell(row=student_row, column=column).value
        student_marks[subject] = marks

    return student_marks   

# STEP 5 - Calculate Student Performance

def calculate_performance(student_marks):
    total = sum(student_marks.values())
    subject_count = len(student_marks)
    max_marks = subject_count * 100
    percentage = (total / max_marks) * 100
    average = total / subject_count
    return total, average, percentage

#Printing the Search Result
if student_row is None:
    print("\nStudent not found.")
else:
    print("\nStudent found!")
    print("Excel Row:", student_row)
    student_marks = get_student_marks(ws, student_row, subjects)
    print("\nSubject Marks:")
    for subject, marks in student_marks.items():
        print(subject, ":", marks)
    total, average, percentage = calculate_performance(student_marks)
    print("\nPerformance:")
    print("Total:", total)
    print("Average:", average)
    print("Percentage:", percentage)

        # OUTPUT:
        # Enter Student ID or Name: 103

        # Student found!
        # Excel Row: 4

        # Subject Marks:
        # Python : 65
        # SQL : 72
        # Excel : 70
        # PowerBI : 68

        # Performance:
        # Total: 275
        # Average: 68.75
        # Percentage: 68.75
    
