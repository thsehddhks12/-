import openpyxl
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

# 1. 원본 엑셀 파일 로드
file_path = "▣ 26년 월간 실적보고 (7월).xlsx"
wb = openpyxl.load_workbook(file_path, data_only=True)
ws = wb.active

# ----------------------------------------------------
# 2. 데이터 추출 (Data Extraction)
# ----------------------------------------------------
# (1) 보고서 타이틀 / 기준월
report_title = ws.cell(row=2, column=2).value or "26년 07월 생산실적 보고"

# (2) 생산계획 및 실적 (7~12행, O열 = 15번째 열)
prod_data = {}
for r in range(7, 13):
    label = ws.cell(row=r, column=2).value or f"항목_{r}"
    val = ws.cell(row=r, column=15).value  # O열
    prod_data[label] = val

# (3) 투입시간 (19행 F열=6, K열=11)
plan_time = ws.cell(row=19, column=6).value   # F19
actual_time = ws.cell(row=19, column=11).value # K19

# (4) 비가동 요인 (21~29행, I~R열 = 9~18번째 열)
downtime_headers = [ws.cell(row=20, column=c).value for c in range(9, 19)]
downtime_rows = []
for r in range(21, 30):
    row_vals = [ws.cell(row=r, column=c).value for c in range(9, 19)]
    if any(row_vals):
        downtime_rows.append(row_vals)

# (5) UPH / PPH 관리 (31~38행, B~Q열 = 2~17번째 열)
uph_headers = [ws.cell(row=30, column=c).value for c in range(2, 18)]
uph_rows = []
for r in range(31, 39):
    row_vals = [ws.cell(row=r, column=c).value for c in range(2, 17)]
    if any(row_vals):
        uph_rows.append(row_vals)

# (6) 완제품 출고 현황 (50행 K열=11, M열=13)
ship_plan = ws.cell(row=50, column=11).value  # K50
ship_actual = ws.cell(row=50, column=13).value # M50

print("✅ 원본 데이터 추출 완료!")

# ----------------------------------------------------
# 3. 임원보고용 요약 대시보드 시트 생성 & 데이터 작성
# ----------------------------------------------------
if "Executive_Dashboard" in wb.sheetnames:
    del wb["Executive_Dashboard"]

dash = wb.create_sheet(title="Executive_Dashboard", index=0)
dash.views.sheetView[0].showGridLines = True

# 서식 스타일 정의
font_title = Font(name="맑은 고딕", size=16, bold=True, color="1F497D")
font_section = Font(name="맑은 고딕", size=11, bold=True, color="FFFFFF")
font_header = Font(name="맑은 고딕", size=10, bold=True)
font_data = Font(name="맑은 고딕", size=10)

fill_section = PatternFill(start_color="1F497D", end_color="1F497D", fill_type="solid") # Dark Blue
fill_header = PatternFill(start_color="F2F2F2", end_color="F2F2F2", fill_type="solid")  # Light Gray
fill_highlight = PatternFill(start_color="FFF2CC", end_color="FFF2CC", fill_type="solid") # Yellow Highlight

thin_border = Border(
    left=Side(style='thin', color='D9D9D9'),
    right=Side(style='thin', color='D9D9D9'),
    top=Side(style='thin', color='D9D9D9'),
    bottom=Side(style='thin', color='D9D9D9')
)

# 1. 타이틀 작성
dash.merge_cells("B2:H2")
dash["B2"] = f"📊 [임원 보고용] {report_title} 요약 대시보드"
dash["B2"].font = font_title
dash["B2"].alignment = Alignment(vertical="center")

# Helper 함수: 섹션 헤더 생성
def create_section_header(ws, start_cell, end_cell, text):
    ws.merge_cells(f"{start_cell}:{end_cell}")
    cell = ws[start_cell]
    cell.value = text
    cell.font = font_section
    cell.fill = fill_section
    cell.alignment = Alignment(horizontal="left", vertical="center", indent=1)

# --- 섹션 1: 생산계획 및 실적 / 투입시간 / 출고 현황 요약 KPI ---
create_section_header(dash, "B4", "H4", "1. 핵심 생산 KPI & 출고 요약")

headers_kpi = ["구분", "발주량", "생산계획", "CAPA계획", "생산실적", "투입시간 (계획/실적)", "완제품출고 (계획/실적)"]
for col_idx, h in enumerate(headers_kpi, start=2):
    cell = dash.cell(row=5, column=col_idx, value=h)
    cell.font = font_header
    cell.fill = fill_header
    cell.alignment = Alignment(horizontal="center", vertical="center")
    cell.border = thin_border

# KPI 데이터 입력
kpi_row = [
    "월간 합계",
    prod_data.get("발주량", 0),
    prod_data.get("생산계획수량", 0),
    prod_data.get("capa 계획 수량", 0),
    prod_data.get("생산실적 수량", 0),
    f"{plan_time or 0}h / {actual_time or 0}h",
    f"{ship_plan or 0}건 / {ship_actual or 0}건"
]

for col_idx, val in enumerate(kpi_row, start=2):
    cell = dash.cell(row=6, column=col_idx, value=val)
    cell.font = font_data
    cell.alignment = Alignment(horizontal="center" if col_idx in [2, 7, 8] else "right", vertical="center")
    cell.border = thin_border
    if col_idx == 6: # 생산실적 음영 강조
        cell.fill = fill_highlight

# --- 섹션 2: 비가동 요인 종합 ---
create_section_header(dash, "B8", "H8", "2. 비가동 요인 종합 분석")

if downtime_headers:
    for col_idx, h in enumerate(downtime_headers[:7], start=2):
        cell = dash.cell(row=9, column=col_idx, value=h)
        cell.font = font_header
        cell.fill = fill_header
        cell.border = thin_border

curr_row = 10
for r_data in downtime_rows:
    for col_idx, val in enumerate(r_data[:7], start=2):
        cell = dash.cell(row=curr_row, column=col_idx, value=val)
        cell.font = font_data
        cell.border = thin_border
    curr_row += 1

# --- 섹션 3: UPH / PPH 라인별 현황 ---
uph_start_row = curr_row + 2
create_section_header(dash, f"B{uph_start_row}", f"H{uph_start_row}", "3. 라인별 UPH / PPH 상세 관리 현황")

header_uph_row = uph_start_row + 1
if uph_headers:
    for col_idx, h in enumerate(uph_headers[:7], start=2):
        cell = dash.cell(row=header_uph_row, column=col_idx, value=h)
        cell.font = font_header
        cell.fill = fill_header
        cell.border = thin_border

curr_row = header_uph_row + 1
for r_data in uph_rows:
    for col_idx, val in enumerate(r_data[:7], start=2):
        cell = dash.cell(row=curr_row, column=col_idx, value=val)
        cell.font = font_data
        cell.border = thin_border
    curr_row += 1

# 열 너비 자동 조절
for col in dash.columns:
    max_len = max(len(str(cell.value or '')) for cell in col)
    col_letter = get_column_letter(col[0].column)
    dash.column_dimensions[col_letter].width = max(max_len + 3, 12)

# 저장
output_path = "▣_26년_월간_실적보고_대시보드_완성본.xlsx"
wb.save(output_path)
print(f"🎉 성공적으로 임원 보고용 대시보드가 생성되었습니다: {output_path}")
