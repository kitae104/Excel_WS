import openpyxl
from openpyxl.drawing.image import Image
wb = openpyxl.Workbook()
ws = wb.active

img_list = ["img1.png", "img2.png", "img3.png"]  # 이미지 리스트
for i, image in enumerate(img_list, start=1):    # 반복 처리
    img = Image(image)
    img.width = 50
    img.height = 50
    ws.add_image(img, f"A{i}")           # 시트에 이미지 추가

wb.save("inserting_image.xlsx")
