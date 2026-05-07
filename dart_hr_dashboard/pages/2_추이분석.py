import sys
import os

import streamlit as st

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
from _shared import (
    format_df, add_group_averages, _mark_partial, _render_table,
    _apply_woori_override,
    _load, _load_t3, render_sidebar, render_debug,
)
from data_parser import (
    PREV_YEAR, RECENT_YEAR, TREND_YEARS,
    build_table1, build_table2,
)

woori_ov = render_sidebar()

st.title("주요 증권사 4개년 인당생산성 및 임직원수 / 순영업수익 추이")

emp_data, fin_data, _src_main = _load()
t3_emp, t3_fin, t3_exec, _src_t3 = _load_t3()

render_debug(emp_data, t3_emp, _src_main, _src_t3)

t1_raw      = build_table1(emp_data, fin_data, RECENT_YEAR)
t1_raw_prev = build_table1(emp_data, fin_data, PREV_YEAR)
t2_raw      = build_table2(emp_data, fin_data)
t1_raw, t1_raw_prev, t2_raw = _apply_woori_override(t1_raw, t1_raw_prev, t2_raw, woori_ov)

y0, y1, y2, y3 = TREND_YEARS
T2_SECTIONS = {
    "인당생산성(백만원/인)":    "인당생산성 (백만원/인)",
    f"임직원수 {y0}년":         f"임직원수  {y0}년",
    f"임직원수 {y1}년":         f"임직원수  {y1}년",
    f"임직원수 {y2}년":         f"임직원수  {y2}년",
    f"임직원수 {y3}년":         f"임직원수  {y3}년",
    "임직원수 평균":             "임직원수  평균",
    f"순영업수익 {y0}년(억원)":  f"순영업수익  {y0}년 (억원)",
    f"순영업수익 {y1}년(억원)":  f"순영업수익  {y1}년 (억원)",
    f"순영업수익 {y2}년(억원)":  f"순영업수익  {y2}년 (억원)",
    f"순영업수익 {y3}년(억원)":  f"순영업수익  {y3}년 (억원)",
    "순영업수익 평균(억원)":      "순영업수익  평균 (억원)",
}
t2_fmt = format_df(add_group_averages(t2_raw)).rename(index=T2_SECTIONS)
t2_fmt = _mark_partial(t2_fmt, fin_data)

st.html(_render_table(t2_fmt, section_rows=True))

st.caption("**순영업수익 계산식**: 세전이익 − 영업외손익 + 판관비")
st.caption("**인당생산성**: 순영업수익 평균(백만원) ÷ 임직원수 평균 (조회기간 평균 기준)")
st.caption("**\\*** 판관비 미공시로 판관비 제외 계산 (세전이익 − 영업외손익)")
