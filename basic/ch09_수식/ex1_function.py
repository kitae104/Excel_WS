import openpyxl

wb = openpyxl.load_workbook("data.xlsx")
sheet = wb.active

sheet['D1'].value = "=SUM(A1:A7)"
sheet['B1'].value = "=A1+A2"
var = 5
sheet['B3'].value = f"=A1*{var}"  # f-string
sheet['D2'].value = "=AVERAGE(A1:A7)"
sheet['D3'].value = "=MIN(A1:A7)"

sheet['E1'].value = '=IF(A1>6, "True", "False")'
sheet['F1'].value = '=CONCATENATE(A1," ",B1)'
sheet['C1'].value = '=COUNT(A1:A7)'
wb.save("data.xlsx")