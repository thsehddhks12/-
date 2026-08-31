import streamlit as st
import pandas as pd
import openpyxl

st.set_page_config(page_title="생산실적 요약 대시보드", layout="wide")

file_path = "▣ 26년 월간 실적보고 (7월).xlsx"

# 1. 엑셀 데이터 추출 (data_only=True 지원)
wb = openpyxl.load_workbook(file_path, data_only=True)
ws = wb.active

report_title = ws.cell(row=2, column=2).value or "26년 07월 생산실적 보고"

# 데이터 추출 및 수식 예외 처리
def get_val(row, col):
    val = ws.cell(row=row, column=col).value
    if val is None or str(val).startswith('='):
        return 0
    try:
        return float(val)
    except:
        return 0

order_qty = get_val(7, 15)
plan_qty = get_val(8, 15)
capa_qty = get_val(9, 15)
actual_qty = get_val(10, 15)

# 백업: O열 수치가 0일 경우 N열 수치 확인
if order_qty == 0 and plan_qty == 0:
    order_qty = get_val(7, 14)
    plan_qty = get_val(8, 14)
    capa_qty = get_val(9, 14)
    actual_qty = get_val(10, 14)

achieve_rate = (actual_qty / plan_qty * 100) if plan_qty > 0 else 0

plan_time = get_val(19, 6)
actual_time = get_val(19, 11)
ship_plan = get_val(50, 11)
ship_actual = get_val(50, 13)

# UI 영역
st.title(f"📊 [임원 보고용] {report_title} 요약 대시보드")
st.markdown("---")

st.subheader("1. 핵심 생산 KPI & 출고 요약")
k1, k2, k3, k4, k5 = st.columns(5)
k1.metric("발주량", f"{order_qty:,.0f} EA")
k2.metric("생산계획", f"{plan_qty:,.0f} EA")
k3.metric("CAPA계획", f"{capa_qty:,.0f} EA")
k4.metric("생산실적", f"{actual_qty:,.0f} EA", delta=f"달성률 {achieve_rate:.1f}%")
k5.metric("완제품출고 (실적/계획)", f"{ship_actual:,.0f} / {ship_plan:,.0f} 건")

st.markdown("---")

st.subheader("2. 비가동 요인 종합 분석")
downtime_headers = [ws.cell(row=20, column=c).value for c in range(9, 19)]
downtime_rows = []
for r in range(21, 30):
    row_vals = [ws.cell(row=r, column=c).value for c in range(9, 19)]
    if any(v is not None and str(v).strip() != '' for v in row_vals):
        downtime_rows.append(row_vals)

df_downtime = pd.DataFrame(downtime_rows, columns=downtime_headers).fillna("-")
st.dataframe(df_downtime, use_container_width=True)

st.markdown("---")

st.subheader("3. 라인별 UPH / PPH 상세 관리 현황")
uph_r29 = [ws.cell(row=29, column=c).value for c in range(2, 18)]
uph_r30 = [ws.cell(row=30, column=c).value for c in range(2, 18)]

uph_cols = []
last_valid = ""
for h1, h2 in zip(uph_r29, uph_r30):
    if h1:
        last_valid = str(h1).strip()
    col_name = last_valid if not h2 else f"{last_valid} ({str(h2).strip()})"
    uph_cols.append(col_name)

uph_rows = []
for r in range(31, 39):
    row_vals = [ws.cell(row=r, column=c).value for c in range(2, 18)]
    if any(v is not None and str(v).strip() != '' for v in row_vals):
        formatted_row = [round(v, 2) if isinstance(v, float) else (v if v is not None else "-") for v in row_vals]
        uph_rows.append(formatted_row)

df_uph = pd.DataFrame(uph_rows, columns=uph_cols)
st.dataframe(df_uph, use_container_width=True)
