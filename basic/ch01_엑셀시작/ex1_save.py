# 엑셀 파일 저장하기
import openpyxl
wb = openpyxl.Workbook()  # 새로운 워크북 생성 

sheet = wb.active         # 활성화된 시트 선택(첫번째 시트가 기본)

list = [1,2,3,4]              # 리스트 데이터 
list2 = ["A", "B", "C", "D"]  #리스트 데이터

sheet.append(list)        # 시트에 추가 
sheet.append(list2)        # 시트에 추가 
sheet.append(list2)        # 시트에 추가 
sheet.append(list2)        # 시트에 추가 
sheet.append(list2)        # 시트에 추가 

wb.save("test.xlsx")    # 저장하기