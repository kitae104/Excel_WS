import openpyxl 
from openpyxl.styles import Font

wb = openpyxl.Workbook()
sheet = wb.active

sheet["A1"] = "data"

# 병합하기 & 병합 풀기 
sheet.merge_cells("B1:E1")
sheet["B1"] = "엑셀 병합 테스트"
sheet.unmerge_cells("B1:E1")

# 포맷 설정하기 
font1 = Font(
  name="Times New Roman",
  bold = True,
  italic=True,
    u = "double",
    color = "ff0011",
    size = 14,
    strike= True 
)

font2 = Font(
    name = "Arial",
    bold = False,
    italic=True,
    u = "single",
    color = "0011FF",
    size = 20,
    strike= False 
)

sheet["A1"].font = font1
sheet["B1"].font = font2

wb.save("merging.xlsx")