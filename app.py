import streamlit as st
import pandas as pd

st.set_page_config(page_title="생산실적 요약 대시보드", layout="wide")

# 1. 구글 스프레드시트 CSV 직접 연동 URL
SHEET_ID = "11qmxisehmu99WG__5BgippmlqMFreFQdycWsn4BVbrQ"
GID = "0" # 시트 탭 ID
CSV_URL = f"https://docs.google.com/spreadsheets/d/{SHEET_ID}/export?format=csv&gid={GID}"

@st.cache_data(ttl=10) # 10초마다 실시간 데이터 갱신
def load_data():
    try:
        # 구글 시트 데이터 로드 (첫번째 행 헤더 처리)
        df = pd.read_csv(CSV_URL, header=None)
        return df
    except Exception as e:
        st.error(f"구글 시트를 불러오는데 실패했습니다. 구글 시트 공유 권한('링크가 있는 모든 사용자 공개')을 확인해 주세요.\n에러 내용: {e}")
        return None

df_raw = load_data()

st.title("📊 [임원 보고용] 생산실적 실시간 요약 대시보드")
st.markdown("---")

if df_raw is not None:
    # 데이터 파싱 및 UI 구성
    st.subheader("1. 1차 입력 데이터 미리보기 및 요약")
    
    # 구글 시트에서 입력받은 전체 데이터 표 출력
    st.dataframe(df_raw.fillna("-"), use_container_width=True)
    
    st.markdown("---")
    
    # 구글 시트 데이터 기반 새로고침 안내
    st.caption("💡 구글 스프레드시트에서 수치를 입력하거나 변경하면 웹 대시보드에 실시간 반영됩니다.")
