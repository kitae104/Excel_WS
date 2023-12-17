import openpyxl
wb = openpyxl.load_workbook('test.xlsx')    # 엑셀 파일 로드 
sheet = wb.active                           # 시트 활성화 
cell_value1 = sheet['A1'].value             # 셀 값
cell_value2 = sheet['B2'].value
print(cell_value1)
print(cell_value2)
