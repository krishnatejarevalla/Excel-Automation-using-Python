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

# STEP 8 - Analyze Student Performance

def analyze_student_performance(student_marks):

    highest_subject = max(student_marks, key=student_marks.get)
    lowest_subject = min(student_marks, key=student_marks.get)

    highest_marks = student_marks[highest_subject]
    lowest_marks = student_marks[lowest_subject]

    if highest_marks >= 90:
        performance_level = "Excellent"
    elif highest_marks >= 75:
        performance_level = "Good"
    elif highest_marks >= 50:
        performance_level = "Average"
    else:
        performance_level = "Needs Improvement"

    return (
        highest_subject,
        highest_marks,
        lowest_subject,
        lowest_marks,
        performance_level
    )

# STEP 6 - Calculate Student Rank

def calculate_student_totals(ws, subjects):

    student_totals = []

    for row in range(2, ws.max_row + 1):

        student_id = ws.cell(row=row, column=1).value
        student_name = ws.cell(row=row, column=2).value

        total = 0

        for column in subjects.values():

            marks = ws.cell(row=row, column=column).value
            total = total + marks

        student_totals.append({
            "id": student_id,
            "name": student_name,
            "total": total
        })

    return student_totals

def calculate_student_rank(student_totals, student_row, ws):

    target_total = 0

    for student in student_totals:

        if student["id"] == ws.cell(row=student_row, column=1).value:
            target_total = student["total"]
            break

    rank = 1

    for student in student_totals:

        if student["total"] > target_total:
            rank = rank + 1

    return rank



# STEP 7 - Generate Complete Student Report

def display_student_report(
    ws,
    student_row,
    student_marks,
    total,
    average,
    percentage,
    rank,
    highest_subject,
    highest_marks,
    lowest_subject,
    lowest_marks,
    performance_level
):

    student_id = ws.cell(row=student_row, column=1).value
    student_name = ws.cell(row=student_row, column=2).value

    print("\n" + "=" * 60)
    print("              STUDENT PERFORMANCE REPORT")
    print("=" * 60)

    print("\nStudent ID :", student_id)
    print("Student    :", student_name)

    print("\n" + "-" * 60)
    print("Subject Marks")
    print("-" * 60)

    for subject, marks in student_marks.items():
        print(f"{subject:<12}: {marks}")

    print("\n" + "-" * 60)
    print("Performance Summary")
    print("-" * 60)

    max_marks = len(student_marks) * 100

    print(f"Total      : {total} / {max_marks}")
    print(f"Average    : {average:.2f}")
    print(f"Percentage : {percentage:.2f}%")
    print(f"Rank       : {rank}")

    print("\n" + "=" * 60) 
# Printing the Search Result

if student_row is None:

    print("\nStudent not found.")

else:

    print("\nStudent found!")
    print("Excel Row:", student_row)

    # Get student's subject marks
    student_marks = get_student_marks(ws, student_row, subjects)

    print("\nSubject Marks:")

    for subject, marks in student_marks.items():
        print(subject, ":", marks)

    # Calculate student's performance
    total, average, percentage = calculate_performance(student_marks)

    # Calculate totals of all students
    student_totals = calculate_student_totals(ws, subjects)

    # Calculate student's rank
    rank = calculate_student_rank(student_totals, student_row, ws)

    # Analyze student's performance
    highest_subject, highest_marks, lowest_subject, lowest_marks, performance_level = analyze_student_performance(student_marks)

    display_student_report(
    ws,
    student_row,
    student_marks,
    total,
    average,
    percentage,
    rank,
    highest_subject,
    highest_marks,
    lowest_subject,
    lowest_marks,
    performance_level
)
    print("\n" + "-" * 60)
    print("Performance Insights")
    print("-" * 60)

    print(f"Highest Subject : {highest_subject} ({highest_marks})")
    print(f"Lowest Subject  : {lowest_subject} ({lowest_marks})")
    print(f"Performance     : {performance_level}")

    print("\n" + "=" * 60)

        # OUTPUT:
        # Enter Student ID or Name: anjali

        # Student found!
        # Excel Row: 7

        # Subject Marks:
        # Python : 95
        # SQL : 91
        # Excel : 96
        # PowerBI : 94

        # ============================================================
        #               STUDENT PERFORMANCE REPORT
        # ============================================================

        # Student ID : 106
        # Student    : Anjali

        # ------------------------------------------------------------
        # Subject Marks
        # ------------------------------------------------------------
        # Python      : 95
        # SQL         : 91
        # Excel       : 96
        # PowerBI     : 94

        # ------------------------------------------------------------
        # Performance Summary
        # ------------------------------------------------------------
        # Total      : 376 / 400
        # Average    : 94.00
        # Percentage : 94.00%
        # Rank       : 1

        # ============================================================

        # ------------------------------------------------------------
        # Performance Insights
        # ------------------------------------------------------------
        # Highest Subject : Excel (96)
        # Lowest Subject  : SQL (91)
        # Performance     : Excellent

        # ============================================================
        # PS D:\Excel Automation with Python> 
    
# STEP 9 - Main Menu

while True:

    print("\n" + "=" * 60)
    print("        STUDENT PERFORMANCE ANALYZER")
    print("=" * 60)

    print("1. Search Student")
    print("2. View Class Statistics")
    print("3. View Top Performers")
    print("4. Exit")

    choice = input("\nEnter your choice: ").strip()

    if choice == "1":
        print("\nSearch Student selected.")

    elif choice == "2":
        print("\nClass Statistics selected.")

    elif choice == "3":
        print("\nTop Performers selected.")

    elif choice == "4":
        print("\nThank you for using Student Performance Analyzer!")
        break

    else:
        print("\nInvalid choice. Please enter 1, 2, 3 or 4.")