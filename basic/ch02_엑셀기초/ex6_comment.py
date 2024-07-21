import openpyxl
from openpyxl.comments import Comment
file = "basic\ch02_엑셀기초\wss2.xlsx"
wb = openpyxl.load_workbook(file)
ws = wb.active

cell = ws['A5']
comment_obj = Comment("이 곳에 숫자가 있습니다","홍 길동")
cell.comment = comment_obj      # 주석 추가 
wb.save(file)