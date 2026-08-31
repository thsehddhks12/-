import streamlit as st
import pandas as pd
import plotly.express as px

# 1. 웹페이지 기본 설정 (넓은 화면 레이아웃)
st.set_page_config(page_title="생산관리 KPI & 비가동 모니터링", layout="wide")

st.title("🏭 생산공정 KPI & 비가동 시간 실시간 모니터링")
st.markdown("---")

# 2. 사이드바 - 엑셀 파일 업로드
st.sidebar.header("📁 데이터 업로드")
uploaded_file = st.sidebar.file_uploader("월간 실적보고 엑셀 파일을 선택하세요", type=["xlsx", "xls"])

if uploaded_file is not None:
    # 엑셀 데이터 불러오기 (시트명 및 데이터 구조에 맞게 자동 처리)
    df_kpi = pd.read_excel(uploaded_file, sheet_name=0) # 필요시 시트명 지정
    
    # -------------------------------------------------------------
    # 🚨 [가장 중요한 3번 항목] 비가동 시간 & 원인 집중 분석 섹션
    # -------------------------------------------------------------
    st.error("🚨 3. 비가동 시간 및 로스(Loss) 현황 분석")
    
    col_loss1, col_loss2 = st.columns([1, 2])
    
    with col_loss1:
        total_downtime = 42.5  # 예시 데이터 (엑셀 수식 연동 가능)
        st.metric(label="총 비가동 시간", value=f"{total_downtime} 시간", delta="-3.5 시간 (전월 대비)")
        st.caption(" 주요 비가동 원인: 설비 정기 점검, 자재 공급 지연, 라인 교체")

    with col_loss2:
        # 비가동 원인별 비중 파이 차트 예시
        loss_data = pd.DataFrame({
            '비가동 사유': ['설비 고장/정비', '자재 대기', '품질 불량 조치', '모델 교체(C/O)', '기타'],
            '시간(H)': [18.5, 12.0, 6.0, 4.5, 1.5]
        })
        fig_loss = px.pie(loss_data, values='시간(H)', names='비가동 사유', title="비가동 사유별 점유율", hole=0.4)
        st.plotly_chart(fig_loss, use_container_width=True)

    st.markdown("---")

    # -------------------------------------------------------------
    # 📊 1, 2, 4번 주요 KPI 요약 섹션
    # -------------------------------------------------------------
    st.subheader("📈 주요 생산 KPI 현황")
    
    kpi_col1, kpi_col2, kpi_col3 = st.columns(3)
    
    with kpi_col1:
        st.metric(label="1. 계획 대비 달성률", value="94.2%", delta="1.2%")
    with kpi_col2:
        st.metric(label="2. CAPA 대비 달성률", value="88.7%", delta="-0.5%")
    with kpi_col3:
        st.metric(label="완제품 출고 현황", value="12,450 개", delta="850 개")

    st.markdown("---")

    # -------------------------------------------------------------
    # ⚙️ 4번 UPH 라인별 전수 관리 섹션
    # -------------------------------------------------------------
    st.subheader("⚙️ 4. 라인별 UPH (시간당 생산량) 전수 관리")
    
    # UPH 샘플 데이터표
    uph_data = pd.DataFrame({
        '라인/설비명': ['조립 1라인', '조립 2라인', '검사 A설비', '검사 B설비', '포장 라인'],
        '표준 UPH': [120, 120, 150, 150, 200],
        '실적 UPH': [115, 122, 142, 148, 195],
        '달성률(%)': [95.8, 101.7, 94.7, 98.7, 97.5]
    })
    
    col_uph1, col_uph2 = st.columns([2, 1])
    
    with col_uph1:
        st.dataframe(uph_data, use_container_width=True)
        
    with col_uph2:
        fig_uph = px.bar(uph_data, x='라인/설비명', y=['표준 UPH', '실적 UPH'], barmode='group', title="라인별 UPH 비교")
        st.plotly_chart(fig_uph, use_container_width=True)

else:
    st.info("👈 왼쪽 사이드바에서 엑셀 파일을 업로드하시면 대시보드가 즉시 생성됩니다.")