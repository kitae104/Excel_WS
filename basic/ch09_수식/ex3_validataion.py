from openpyxl import Workbook
from openpyxl.worksheet.datavalidation import DataValidation

wb = Workbook()
ws = wb.active

data_validation = DataValidation(type="decimal", operator="between", formula1=0, formula2=50)
data_validation.add("A1:A10")

ws.add_data_validation(data_validation)

wb.save("basic/ch09_수식/data_validation.xlsx")