import os, openpyxl
from openpyxl.worksheet.properties import PageSetupProperties
from openpyxl.utils import get_column_letter

IN = os.environ.get("INFILE")
if not IN:
    raise SystemExit("INFILE 환경변수에 대상 xlsx 경로를 지정하세요")
wb = openpyxl.load_workbook(IN)

for name, last in [("(TV)편성기획", 32), ("(TV)편성기획-팀편성", 30)]:
    ws = wb[name]
    W = sum(round((ws.column_dimensions[get_column_letter(c)].width or 8.43)*7)+5
            for c in range(1, 44))
    H = sum(round((ws.row_dimensions[r].height or 15)*96/72) for r in range(1, last+1))
    # 여백 0 + 시트보다 살짝 큰 사용자 용지 → 100% 배율(축소 없음) 강제
    ws.page_setup.orientation = "portrait"
    ws.page_setup.fitToPage = False
    ws.page_setup.scale = 100
    ws.page_setup.paperWidth = f"{W*1.12/96*25.4:.2f}mm"
    ws.page_setup.paperHeight = f"{H*1.12/96*25.4:.2f}mm"
    ws.sheet_properties.pageSetUpPr = PageSetupProperties(fitToPage=False)
    ws.print_area = f"A1:AQ{last}"
    ws.page_margins.left = ws.page_margins.right = 0
    ws.page_margins.top = ws.page_margins.bottom = 0
    ws.page_margins.header = ws.page_margins.footer = 0

wb["Sheet1"].sheet_state = "hidden"
wb.save(os.environ.get("OUTFILE", "render_exact.xlsx"))
print("ok")
