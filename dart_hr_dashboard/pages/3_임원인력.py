import sys
import os

import streamlit as st

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
from _shared import (
    format_df_t3, _render_table3,
    _load, _load_t3, render_sidebar, render_debug,
)
from data_parser import RECENT_YEAR, build_table3

render_sidebar()

st.title(f"주요 증권사 임원인력 상세현황 ({RECENT_YEAR}년)")
st.caption("임원 인원수: 상근임원 기준 | 자기자본: 조원 단위 | 급여·근속: DART 직원현황 신고 기준")

emp_data, fin_data, _src_main = _load()
t3_emp, t3_fin, t3_exec, _src_t3 = _load_t3()

render_debug(emp_data, t3_emp, _src_main, _src_t3)

t3_raw = build_table3(t3_emp, t3_fin, t3_exec, RECENT_YEAR)
t3_fmt = format_df_t3(t3_raw)
st.html(_render_table3(t3_fmt))
