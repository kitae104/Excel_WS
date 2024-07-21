from openpyxl import load_workbook
from openpyxl.workbook.protection import WorkbookProtection

wb = load_workbook("basic/ch09_수식/myfile.xlsx")
ws = wb.active

wb.security = WorkbookProtection(workbookPassword="12345", lockStructure=True)   # 구조 잠금

ws.protection.set_password("123")    # 시트 잠금
ws.protection.enable()

wb.save("basic/ch09_수식/myfile_protected.xlsx")
