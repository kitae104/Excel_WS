from openpyxl import load_workbook
wb = load_workbook("basic/ch07_정렬_필터/print_setting.xlsx")
ws = wb.active

ws.print_options.horizontalCentered = True  # 수평 중앙 정렬
ws.print_options.verticalCentered = True    # 수직 중앙 정렬

# 홀수 페이지 헤더 설정
ws.oddHeader.left.text = "Page &[Page] of &N"  # 홀수 페이지 헤더
ws.oddHeader.left.size = 14
ws.oddHeader.left.font = "Tahoma,Bold"
ws.oddHeader.left.color = "CC3366"

# 홀수 페이지 푸터 설정
ws.oddFooter.right.text = "Odd Page Footer - &[File]"	# evenFooter
ws.oddFooter.right.size = 12
ws.oddFooter.right.font = "Verdana"
ws.oddFooter.right.color = "CC0000"

ws.print_title_cols = 'A:D' # 처음 두 열
ws.print_title_rows = '1:1' # 첫 번째 행

ws.print_area = 'A1:D10'    # 인쇄 범위

ws.page_setup.orientation = ws.ORIENTATION_LANDSCAPE  # 용지 방향
ws.page_setup.paperSize = ws.PAPERSIZE_A5  # 용지 크기

wb.save("basic/ch07_정렬_필터/print_setting2.xlsx")