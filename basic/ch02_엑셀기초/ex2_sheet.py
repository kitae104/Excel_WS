import openpyxl
workbook = openpyxl.Workbook()            
sheet = workbook.active
sheet["A1"] = "First sheet"                 # 기본 시트의 셀에 데이터 입력 
sheet1 = workbook.create_sheet("test1")     # 새로운 시트 생성 
sheet2 = workbook.create_sheet("test2")
sheet3 = workbook.create_sheet("test3")
sheet1['A4'] = "This is the test1 sheet"
sheet2['B1'] = "This is the test2 sheet"
sheet3['C3'] = "This is the test3 sheet"

workbook.save("sheets.xlsx")