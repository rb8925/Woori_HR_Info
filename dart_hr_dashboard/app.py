import streamlit as st

st.set_page_config(
    page_title="증권사 인력 현황 비교",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded",
)

pg = st.navigation([
    st.Page("pages/1_인력현황.py",  title="인력현황",        icon="📊", default=True),
    st.Page("pages/2_추이분석.py",  title="4개년 추이분석",  icon="📈"),
    st.Page("pages/3_임원인력.py",  title="임원인력 상세현황", icon="👔"),
])
pg.run()
