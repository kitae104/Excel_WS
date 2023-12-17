import openpyxl
file = "basic\ch02_엑셀기초\mydata.xlsx"
workbook = openpyxl.load_workbook(file)
sheet = workbook.active

# 전체 데이터 읽어오기(min_row : 시작행, max_row : 끝 행)
for row in sheet.iter_rows(min_row = 2, values_only = True):
    name, age, city = row     # row의 튜플 정보를 각 변수에 할당 
    print(name)
    print(age)
    print(city)