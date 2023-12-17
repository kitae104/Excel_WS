import openpyxl
workbook = openpyxl.Workbook()        # 워크북 생성
sheet = workbook.active               # 시트 생성 
sheet.title = "python"                # 시트 타이틀 
sheet["A1"] = "Name"                  # 특정 셀에 데이터 입력 
sheet["C2"] = 20
sheet["A3"] = "gil dong"
sheet["F7"] = "test data"
workbook.save("input_test.xlsx")      # 파일 저장 