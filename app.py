import streamlit as st
import pandas as pd
import openpyxl

st.set_page_config(page_title="생산실적 요약 대시보드", layout="wide")

file_path = "▣ 26년 월간 실적보고 (7월).xlsx"

# 1. 엑셀 파일 로드 및 안전 추출 함수
wb = openpyxl.load_workbook(file_path, data_only=True)
ws = wb.active

report_title = ws.cell(row=2, column=2).value or "26년 07월 생산실적 보고"

def safe_val(row, col):
    val = ws.cell(row=row, column=col).value
    if val is None or str(val).startswith('='):
        return 0
    try:
        return float(val)
    except:
        return 0

# UI 제목
st.title(f"📊 [임원 보고용] {report_title} 요약 대시보드")
st.markdown("---")

# ----------------------------------------------------
# 1. 비가동 요인 종합 분석
# ----------------------------------------------------
st.subheader("1. 비가동 요인 (Hr)")

# 20, 21행 조합 헤더
downtime_headers = ["리워크", "결품", "보급지연", "불량", "스페어파츠 포장", "교육", "기타", "합계"]

downtime_rows = []
for r in range(22, 30):
    row_vals = [ws.cell(row=r, column=c).value for c in range(9, 17)]
    # 수치 데이터 정제
    clean_vals = [round(v, 2) if isinstance(v, (int, float)) else "-" for v in row_vals]
    if any(v != "-" for v in clean_vals):
        downtime_rows.append(clean_vals)

if not downtime_rows: # 비가동 내역이 없을 경우 예시 기본행 처리
    downtime_rows = [["-"] * len(downtime_headers)]

df_downtime = pd.DataFrame(downtime_rows, columns=downtime_headers)
st.dataframe(df_downtime, use_container_width=True)

st.markdown("---")

# ----------------------------------------------------
# 2. 라인별 UPH / PPH 관리 현황
# ----------------------------------------------------
st.subheader("2. 라인별 UPH / PPH 관리 현황")

uph_cols = [
    "생산과", "가조립(분해)", "1과_SUB", "1과_조립", "1과_TEST", "1과_옵션", "1과_포장",
    "2과_SUB", "2과_조립", "2과_TEST&옵션", "2과_포장", "2과_뉴메틱", "3과_헤비듀티", "3과_ITM"
]

uph_rows = []
row_titles = ["투입시간(H)", "생산수량(EA)", "인원수(명)", "UPH", "PPH", "UPD"]

for idx, r in enumerate(range(12, 18)): # UPH 영역 행 범위
    row_vals = [ws.cell(row=r, column=c).value for c in range(2, 16)]
    formatted_vals = [row_titles[idx]]
    
    for val in row_vals:
        if isinstance(val, (int, float)):
            formatted_vals.append(f"{val:,.1f}")
        elif str(val).startswith("#") or val is None: # #DIV/0! 오류 정제
            formatted_vals.append("-")
        else:
            formatted_vals.append(str(val))
            
    uph_rows.append(formatted_vals)

df_uph = pd.DataFrame(uph_rows, columns=uph_cols)
st.dataframe(df_uph, use_container_width=True)

st.markdown("---")

# ----------------------------------------------------
# 3. 완제품 출고 현황 요약
# ----------------------------------------------------
st.subheader("3. 완제품 출고 현황 요약")

ship_headers = ["정규 오더 준수율", "정기_계획", "정기_실적", "정기_준수율", "긴급_계획", "긴급_실적", "긴급_준수율", "합계_계획", "합계_실적", "합계_준수율", "비고"]

ship_rows = []
for r in range(20, 25): # 출고 현황 데이터 행 탐색
    group_name = ws.cell(row=r, column=2).value
    if group_name:
        row_data = [
            group_name,
            safe_val(r, 4), safe_val(r, 5), f"{safe_val(r, 6)*100:.0f}%",
            safe_val(r, 7), safe_val(r, 8), f"{safe_val(r, 9)*100:.0f}%",
            safe_val(r, 10), safe_val(r, 11), f"{safe_val(r, 12)*100:.0f}%",
            ws.cell(row=r, column=13).value or "-"
        ]
        ship_rows.append(row_data)

df_ship = pd.DataFrame(ship_rows, columns=ship_headers)
st.dataframe(df_ship, use_container_width=True)
