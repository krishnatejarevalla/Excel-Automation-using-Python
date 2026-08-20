from openpyxl import Workbook
wb = Workbook()
ws = wb.active
ws.title = "MarkList"

#Creating column names and entering data
ws.append(["Student_ID","Name","Python","SQL","Excel","PowerBI"])
ws.append([101, "Ravi", 78, 85, 90, 82])
ws.append([102, "Priya", 92, 88, 84, 91])
ws.append([103, "Arjun", 65, 72, 70, 68])
ws.append([104, "Sneha", 88, 94, 91, 89])
ws.append([105, "Kiran", 74, 69, 80, 76])
ws.append([106, "Anjali", 95, 91, 96, 94])
ws.append([107, "Rahul", 81, 77, 85, 79])
ws.append([108, "Divya", 89, 93, 88, 90])
ws.append([109, "Vijay", 70, 68, 75, 72])
ws.append([110, "Neha", 86, 90, 92, 87])
wb.save("MarkList.xlsx")

