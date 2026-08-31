import streamlit as st
import pandas as pd
import openpyxl

st.set_page_config(page_title="생산실적 요약 대시보드", layout="wide")

# 엑셀 파일 읽기 (GitHub 저장소 내 엑셀 파일 자동 추출)
file_path = "▣ 26년 월간 실적보고 (7월).xlsx"
wb = openpyxl.load_workbook(file_path, data_only=True)
ws = wb.active

report_title = ws.cell(row=2, column=2).value or "26년 07월 생산실적 보고"

# 1번 항목 수치
order_qty = ws.cell(row=7, column=15).value or 0
plan_qty = ws.cell(row=8, column=15).value or 0
capa_qty = ws.cell(row=9, column=15).value or 0
actual_qty = ws.cell(row=10, column=15).value or 0
achieve_rate = (actual_qty / plan_qty * 100) if plan_qty > 0 else 0

# 투입시간 & 완제품 출고
plan_time = ws.cell(row=19, column=6).value or 0
actual_time = ws.cell(row=19, column=11).value or 0
ship_plan = ws.cell(row=50, column=11).value or 0
ship_actual = ws.cell(row=50, column=13).value or 0

# 메인 UI
st.title(f"📊 [임원 보고용] {report_title} 요약 대시보드")
st.markdown("---")

# 1. KPI 카드
st.subheader("1. 핵심 생산 KPI & 출고 요약")
k1, k2, k3, k4, k5 = st.columns(5)
k1.metric("발주량", f"{order_qty:,} EA")
k2.metric("생산계획", f"{plan_qty:,} EA")
k3.metric("CAPA계획", f"{capa_qty:,} EA")
k4.metric("생산실적", f"{actual_qty:,} EA", delta=f"달성률 {achieve_rate:.1f}%")
k5.metric("완제품출고 (실적/계획)", f"{ship_actual:,} / {ship_plan:,} 건")

st.markdown("---")

# 2. 비가동 요인 분석
st.subheader("2. 비가동 요인 종합 분석")
downtime_headers = [ws.cell(row=20, column=c).value for c in range(9, 19)]
downtime_rows = []
for r in range(21, 30):
    row_vals = [ws.cell(row=r, column=c).value for c in range(9, 19)]
    if any(v is not None and str(v).strip() != '' for v in row_vals):
        downtime_rows.append(row_vals)

df_downtime = pd.DataFrame(downtime_rows, columns=downtime_headers)
st.dataframe(df_downtime, use_container_width=True)

st.markdown("---")

# 3. UPH / PPH 라인별 현황
st.subheader("3. 라인별 UPH / PPH 상세 관리 현황")
uph_headers_r1 = [ws.cell(row=29, column=c).value for c in range(2, 18)]
uph_headers_r2 = [ws.cell(row=30, column=c).value for c in range(2, 18)]

# 헤더 결합
combined_uph_headers = []
for h1, h2 in zip(uph_headers_r1, uph_headers_r2):
    if h1 and h2 and h1 != h2:
        combined_uph_headers.append(f"{h1} ({h2})")
    else:
        combined_uph_headers.append(h1 or h2 or "")

uph_rows = []
for r in range(31, 39):
    row_vals = [ws.cell(row=r, column=c).value for c in range(2, 18)]
    if any(v is not None and str(v).strip() != '' for v in row_vals):
        uph_rows.append(row_vals)

df_uph = pd.DataFrame(uph_rows, columns=combined_uph_headers)
st.dataframe(df_uph, use_container_width=True)
