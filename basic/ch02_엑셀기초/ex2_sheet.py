import openpyxl
workbook = openpyxl.Workbook()            
ws = workbook.active
ws["A1"] = "First ws"                 # 기본 시트의 셀에 데이터 입력 
ws1 = workbook.create_ws("test1")     # 새로운 시트 생성 
ws2 = workbook.create_ws("test2")
ws3 = workbook.create_ws("test3")
ws1['A4'] = "This is the test1 ws"
ws2['B1'] = "This is the test2 ws"
ws3['C3'] = "This is the test3 ws"

workbook.save("wss.xlsx")