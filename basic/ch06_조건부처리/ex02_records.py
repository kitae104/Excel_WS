import openpyxl
from openpyxl.styles import PatternFill

workbook = openpyxl.load_workbook("file.xlsx")
sheet = workbook.active

total_rows = sheet.max_row
start_row = max(total_rows-3, 1)
end_row = total_rows

for row in range(start_row,end_row):
    for col in range(2,4): # A to D
        sheet.cell(row=row, column = col).fill = PatternFill(start_color="ffccaa", end_color="123456", fill_type="solid")

workbook.save("file.xlsx")