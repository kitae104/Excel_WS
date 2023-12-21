import openpyxl 
from openpyxl.worksheet.table import Table
from openpyxl.styles import PatternFill

wb = openpyxl.Workbook()
sheet = wb.active

data = [
  ["이름", "나이", "도시"],
  ["홍길동", 20, "서울"],
  ["김철수", 30, "인천"],
  ["박성진", 40, "대전"],
  ["이영희", 50, "대구"],
  ["최선이", 60, "부산"]
]

# 데이터 입력
for row in data:
  sheet.append(row)

# 테이블 만들기 
table_range = f"A1:C{len(data)}" # A1 to C5
table = Table(displayName="MyTableData", ref=table_range)

# 행 삽입 후 데이터 입력 
sheet.insert_rows(2)
sheet.cell(row = 2 , column=1, value="테스터")
sheet.cell(row = 2 , column=2, value=10)
sheet.cell(row = 2 , column=3, value="강원")

sheet.insert_rows(3)
sheet.cell(row = 3 , column=1, value="New1")
sheet.cell(row = 3 , column=2, value="New2")
sheet.cell(row = 3 , column=3, value="New3")

# 열 삽입 후 데이터 입력
sheet.insert_cols(2)
sheet.cell(row=1, column =2, value="테이블")

sheet.cell(row = 2 , column=2, value=23)
sheet.cell(row = 3 , column=2, value=24)
sheet.cell(row = 4 , column=2, value=55)
sheet.cell(row = 5 , column=2, value=13)
sheet.cell(row = 6 , column=2, value=44)
sheet.cell(row = 7 , column=2, value=65)

# 테이블 스타일 지정 - 2번째 행에 대해서만 적용
for cell in sheet[2]:
    cell.fill = PatternFill(start_color="ffaabb", end_color="123456", fill_type="solid")

for cell in sheet[7]:
    cell.fill = PatternFill(start_color="ffccbb", end_color="ff3456", fill_type="solid")

for row in sheet.iter_rows(min_col = 3, max_col = 3):
    for cell in row:
        cell.fill = PatternFill(start_color="ffaabb", end_color="123456", fill_type="solid")

sheet.add_table(table)
wb.save("table.xlsx")