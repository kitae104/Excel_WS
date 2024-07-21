import openpyxl
file = "basic\ch02_엑셀기초\mydata.xlsx"
wb = openpyxl.load_workbook(file)
ws = wb.active

# 전체 데이터 읽어오기(min_row : 시작행, max_row : 끝 행)
for row in ws.iter_rows(min_row = 2, values_only = True):
    name, age, city = row     # row의 튜플 정보를 각 변수에 할당 
    print(name, end=" ")
    print(age, end=" ")
    print(city, end=" ")
    print()