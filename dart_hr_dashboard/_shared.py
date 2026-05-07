import os
import datetime

import pandas as pd
import streamlit as st

from data_parser import (
    CACHE_PATH,
    COMPANIES,
    PREV_YEAR,
    RECENT_YEAR,
    TREND_YEARS,
    build_table1,
    build_table2,
    build_table3,
    fetch_all_raw,
    fetch_t3_raw,
)

# ── 상수 ──────────────────────────────────────────────────────────────────────
HIGHLIGHT_CO = "우리투자증권"
HIGHLIGHT_BG = "#FFF8DC"
SUB_REG_BG   = "#E8F8F5"
SUB_CNT_BG   = "#F5EEF8"

GROUPS_DEF: dict[str, list[str]] = {
    "대형사":    ["미래에셋증권", "한국투자증권", "NH투자증권", "KB증권"],
    "중·대형사": ["키움증권", "메리츠증권"],
    "중·소형사": ["유안타증권", "현대차증권", "IBK투자증권", "유진투자증권", "대신증권", "교보증권"],
}
AVG_COL_NAME: dict[str, str] = {g: f"{g} 평균" for g in GROUPS_DEF}
AVG_COLS: set[str] = set(AVG_COL_NAME.values())

COMPANY_GROUP: dict[str, str] = {
    "우리투자증권":              "",
    "미래에셋증권":              "대형사",
    "한국투자증권":              "대형사",
    "NH투자증권":                "대형사",
    "KB증권":                    "대형사",
    AVG_COL_NAME["대형사"]:      "대형사",
    "키움증권":                  "중·대형사",
    "메리츠증권":                "중·대형사",
    AVG_COL_NAME["중·대형사"]:   "중·대형사",
    "유안타증권":                "중·소형사",
    "현대차증권":                "중·소형사",
    "IBK투자증권":               "중·소형사",
    "유진투자증권":              "중·소형사",
    "대신증권":                  "중·소형사",
    "교보증권":                  "중·소형사",
    AVG_COL_NAME["중·소형사"]:   "중·소형사",
}

T3_GROUPS_DEF: dict[str, list[str]] = {
    "대형사":                 ["한국투자증권", "미래에셋증권", "삼성증권"],
    "5대금융지주 산하 대형사": ["NH투자증권", "KB증권", "하나증권", "신한투자증권"],
    "중대형사":               ["메리츠증권", "키움증권"],
    "중형사":                 ["대신증권", "교보증권", "한화투자증권", "신영증권"],
    "중소형사":               ["현대차증권", "IBK투자증권", "유안타증권", "유진투자증권"],
}
T3_COMPANY_GROUP: dict[str, str] = {
    co: g for g, cos in T3_GROUPS_DEF.items() for co in cos
}

T1_ROW_LABELS = {
    "자기자본(억원)": "자기자본 (억원)",
    "리테일 인원":    "리테일 인원",
    "본사영업 인원":  "본사영업 인원",
    "본사관리 인원":  "본사관리 인원",
    "총인원":         "총인원",
}

# ── 포맷 함수 ─────────────────────────────────────────────────────────────────
def _fmt(v, kind: str) -> str:
    if v is None or (isinstance(v, float) and pd.isna(v)):
        return "-"
    try:
        if kind == "int":     return f"{int(round(v)):,}"
        if kind == "pct":     return f"{v:,.1f}%"
        return f"{v:,.1f}"
    except Exception:
        return "-"

_ROW_FMT: dict[str, str] = {
    "자기자본(억원)":         "int",
    "리테일 인원":            "int",
    "리테일 비중(%)":         "pct",
    "리테일 정규직 인원":      "int",
    "리테일 정규직 비중(%)":   "pct",
    "리테일 기간제 인원":      "int",
    "리테일 기간제 비중(%)":   "pct",
    "본사영업 인원":          "int",
    "본사영업 비중(%)":       "pct",
    "본사영업 정규직 인원":    "int",
    "본사영업 정규직 비중(%)": "pct",
    "본사영업 기간제 인원":    "int",
    "본사영업 기간제 비중(%)": "pct",
    "본사관리 인원":          "int",
    "본사관리 비중(%)":       "pct",
    "본사관리 정규직 인원":    "int",
    "본사관리 정규직 비중(%)": "pct",
    "본사관리 기간제 인원":    "int",
    "본사관리 기간제 비중(%)": "pct",
    "총인원":                 "int",
    "인당생산성(백만원/인)":   "decimal",
    "임직원수 평균":           "int",
    "순영업수익 평균(억원)":   "int",
    **{f"임직원수 {y}년":        "int" for y in TREND_YEARS},
    **{f"순영업수익 {y}년(억원)": "int" for y in TREND_YEARS},
}

def format_df(df: pd.DataFrame) -> pd.DataFrame:
    out = df.copy().astype(object)
    for row_label in out.index:
        kind = _ROW_FMT.get(row_label, "decimal")
        out.loc[row_label] = out.loc[row_label].apply(lambda v: _fmt(v, kind))
    return out

def _merge_pct_rows(df: pd.DataFrame) -> pd.DataFrame:
    pairs = [
        ("리테일 인원",          "리테일 비중(%)"),
        ("리테일 정규직 인원",    "리테일 정규직 비중(%)"),
        ("리테일 기간제 인원",    "리테일 기간제 비중(%)"),
        ("본사영업 인원",        "본사영업 비중(%)"),
        ("본사영업 정규직 인원",  "본사영업 정규직 비중(%)"),
        ("본사영업 기간제 인원",  "본사영업 기간제 비중(%)"),
        ("본사관리 인원",        "본사관리 비중(%)"),
        ("본사관리 정규직 인원",  "본사관리 정규직 비중(%)"),
        ("본사관리 기간제 인원",  "본사관리 기간제 비중(%)"),
    ]
    skip = {p for _, p in pairs}
    pair_map = {h: p for h, p in pairs}
    rows_data: dict[str, dict] = {}
    for lbl in df.index:
        if lbl in skip:
            continue
        pct_lbl = pair_map.get(lbl)
        if pct_lbl and pct_lbl in df.index:
            merged: dict[str, str] = {}
            for col in df.columns:
                hv, pv = df.loc[lbl, col], df.loc[pct_lbl, col]
                merged[col] = "-" if hv == "-" else (hv if pv == "-" else f"{hv} ({pv})")
            rows_data[lbl] = merged
        else:
            rows_data[lbl] = df.loc[lbl].to_dict()
    return pd.DataFrame(rows_data).T

def add_group_averages(df: pd.DataFrame) -> pd.DataFrame:
    out = df.copy()
    cols_ordered: list[str] = []
    for co in df.columns:
        cols_ordered.append(co)
        for g, members in GROUPS_DEF.items():
            if co == members[-1]:
                avg_col = AVG_COL_NAME[g]
                valid = [c for c in members if c in df.columns]
                out[avg_col] = df[valid].mean(axis=1, skipna=True)
                cols_ordered.append(avg_col)
    return out[cols_ordered]

# ── HTML 테이블 렌더러 ────────────────────────────────────────────────────────
def _render_table(df_fmt: pd.DataFrame, section_rows: bool = False) -> str:
    import re as _re
    companies = list(df_fmt.columns)
    col_groups: list[list] = []
    for co in companies:
        g = COMPANY_GROUP.get(co, "")
        if col_groups and col_groups[-1][0] == g:
            col_groups[-1][1].append(co)
        else:
            col_groups.append([g, [co]])

    PROD_BG    = "#E9F7EF"
    AVG_BG     = "#EEF2F7"
    SEC_HDR_BG = "#EDF4FC"
    S = {
        "tbl":         "border-collapse:collapse;width:100%;font-size:13px;font-family:sans-serif;",
        "th":          "border:1px solid #ccc;padding:6px 10px;font-weight:600;white-space:nowrap;text-align:center;background:#F2F3F4;",
        "th_hl":       "border:1px solid #ccc;padding:6px 10px;font-weight:600;white-space:nowrap;text-align:center;background:#FFF8DC;",
        "th_grp":      "border:1px solid #ccc;padding:6px 10px;font-weight:700;white-space:nowrap;text-align:center;background:#D6EAF8;",
        "th_idx":      "border:1px solid #ccc;padding:6px 10px;font-weight:700;white-space:nowrap;text-align:center;background:#F2F3F4;",
        "th_avg":      f"border:1px solid #ccc;padding:6px 10px;font-weight:600;white-space:nowrap;text-align:center;background:{AVG_BG};font-style:italic;",
        "td":          "border:1px solid #ccc;padding:5px 10px;text-align:right;white-space:nowrap;",
        "td_hl":       "border:1px solid #ccc;padding:5px 10px;text-align:right;white-space:nowrap;background:#FFF8DC;font-weight:600;",
        "td_avg":      f"border:1px solid #ccc;padding:5px 10px;text-align:right;white-space:nowrap;background:{AVG_BG};font-style:italic;",
        "td_row":      "border:1px solid #ccc;padding:5px 10px;font-weight:600;text-align:left;white-space:nowrap;",
        "td_row_sec":  f"border:1px solid #ccc;border-top:2px solid #8BADC8;padding:5px 10px;font-weight:700;text-align:left;white-space:nowrap;background:{SEC_HDR_BG};",
        "td_sec":      "border:1px solid #ccc;padding:5px 10px;font-weight:700;text-align:center;vertical-align:middle;background:#EBF5FB;white-space:nowrap;",
        "td_prod":     f"border:1px solid #ccc;padding:5px 10px;font-weight:700;text-align:left;white-space:nowrap;background:{PROD_BG};",
        "td_prod_d":   f"border:1px solid #ccc;padding:5px 10px;text-align:right;white-space:nowrap;background:{PROD_BG};",
        "td_prod_hl":  f"border:1px solid #ccc;padding:5px 10px;text-align:right;white-space:nowrap;background:{PROD_BG};font-weight:600;",
        "td_prod_avg": f"border:1px solid #ccc;padding:5px 10px;text-align:right;white-space:nowrap;background:{PROD_BG};font-style:italic;",
    }
    n_idx = 2 if section_rows else 1
    r1, r2 = [], []
    r1.append(f'<th rowspan="2" colspan="{n_idx}" style="{S["th_idx"]}">구분</th>')
    for g, cos in col_groups:
        if not g:
            for co in cos:
                s = S["th_hl"] if co == HIGHLIGHT_CO else S["th"]
                r1.append(f'<th rowspan="2" style="{s}">{co}</th>')
        else:
            r1.append(f'<th colspan="{len(cos)}" style="{S["th_grp"]}">{g}</th>')
            for co in cos:
                s = S["th_avg"] if co in AVG_COLS else (S["th_hl"] if co == HIGHLIGHT_CO else S["th"])
                display = "평균" if co in AVG_COLS else co
                r2.append(f'<th style="{s}">{display}</th>')
    thead = f'<thead><tr>{"".join(r1)}</tr><tr>{"".join(r2)}</tr></thead>'

    def _td(val, co, prod=False):
        is_avg = co in AVG_COLS
        if prod:
            s = S["td_prod_avg"] if is_avg else S["td_prod_hl"] if co == HIGHLIGHT_CO else S["td_prod_d"]
        else:
            s = S["td_avg"] if is_avg else S["td_hl"] if co == HIGHLIGHT_CO else S["td"]
        return f'<td style="{s}">{val}</td>'

    body_rows = []
    labels = list(df_fmt.index)
    if section_rows:
        def _sec(lbl):
            if "임직원수" in lbl:   return "임직원수"
            if "순영업수익" in lbl: return "순영업수익"
            return ""
        sec_groups: list[list] = []
        for lbl in labels:
            s = _sec(lbl)
            if sec_groups and sec_groups[-1][0] == s:
                sec_groups[-1][1].append(lbl)
            else:
                sec_groups.append([s, [lbl]])
        SEC_DISPLAY = {"임직원수": "임직원수 (명)", "순영업수익": "순영업수익 (억원)"}
        def _disp(lbl, sec):
            if sec in ("임직원수", "순영업수익"):
                if "평균" in lbl: return "평균"
                m = _re.search(r"(\d{4})", lbl)
                return f"{m.group(1)[2:]}년 12월" if m else lbl
            return lbl
        for sec, lbls in sec_groups:
            is_prod = (sec == "")
            for j, lbl in enumerate(lbls):
                cells = []
                if is_prod:
                    cells.append(f'<td colspan="2" style="{S["td_prod"]}">{_disp(lbl, sec)}</td>')
                else:
                    if j == 0:
                        cells.append(f'<td rowspan="{len(lbls)}" style="{S["td_sec"]}">{SEC_DISPLAY.get(sec, sec)}</td>')
                    cells.append(f'<td style="{S["td_row"]}">{_disp(lbl, sec)}</td>')
                cells += [_td(df_fmt.loc[lbl, co], co, prod=is_prod) for co in companies]
                body_rows.append(f'<tr>{"".join(cells)}</tr>')
    else:
        _SECTION_HDRS = {"리테일 인원", "본사영업 인원", "본사관리 인원"}
        def _sub_kind(lbl):
            for sec in ("리테일", "본사영업", "본사관리"):
                for et in ("정규직", "기간제"):
                    if lbl == f"{sec} {et} 인원": return et
            return None
        for lbl in labels:
            kind = _sub_kind(lbl)
            if kind:
                bg = SUB_REG_BG if kind == "정규직" else SUB_CNT_BG
                tc = "#1A8C6E" if kind == "정규직" else "#7D3C98"
                lbl_cell = (f'<td style="border:1px solid #ddd;padding:4px 8px;text-align:left;'
                            f'white-space:nowrap;font-size:12px;background:{bg};">'
                            f'<span style="color:{tc};padding-left:10px;">▸ {kind}</span></td>')
                cells = [lbl_cell]
                for co in companies:
                    val = df_fmt.loc[lbl, co]
                    is_avg = co in AVG_COLS
                    cell_bg = AVG_BG if is_avg else (HIGHLIGHT_BG if co == HIGHLIGHT_CO else bg)
                    fw = "font-weight:600;" if co == HIGHLIGHT_CO and not is_avg else ""
                    fi = "font-style:italic;" if is_avg else ""
                    cell_s = (f"border:1px solid #ddd;padding:4px 9px;text-align:right;"
                              f"white-space:nowrap;background:{cell_bg};font-size:12px;{fw}{fi}")
                    cells.append(f'<td style="{cell_s}">{val}</td>')
            elif lbl in _SECTION_HDRS:
                cells = [f'<td style="{S["td_row_sec"]}">{lbl}</td>']
                cells += [_td(df_fmt.loc[lbl, co], co) for co in companies]
            else:
                cells = [f'<td style="{S["td_row"]}">{lbl}</td>']
                cells += [_td(df_fmt.loc[lbl, co], co) for co in companies]
            body_rows.append(f'<tr>{"".join(cells)}</tr>')

    tbody = f'<tbody>{"".join(body_rows)}</tbody>'
    return f'<div style="overflow-x:auto;"><table style="{S["tbl"]}">{thead}{tbody}</table></div>'


# ── Table 3 포맷 & 렌더러 ────────────────────────────────────────────────────
_T3_ROW_FMT: dict[str, str] = {
    "자기자본(조원)":           "decimal2",
    "총인원":                   "int",
    "임원 인원수":              "int",
    "임원 비중(%)":             "pct2",
    "리테일 직원수":            "int",
    "리테일 비중(%)":           "pct1",
    "리테일 제외 임원 비중(%)": "pct2",
    "평균급여(백만원)":         "int",
    "평균근속(년)":             "decimal1",
    **{k: "int" for k in [
        "리테일 정규직", "리테일 계약직", "리테일 소계",
        "본사영업 정규직", "본사영업 계약직", "본사영업 소계",
        "본사관리 정규직", "본사관리 계약직", "본사관리 소계",
        "정규직 총합", "계약직 총합", "총합",
    ]},
}

def _fmt_t3(v, kind: str) -> str:
    if v is None or (isinstance(v, float) and pd.isna(v)):
        return "-"
    try:
        if kind == "int":      return f"{int(round(v)):,}"
        if kind == "pct1":     return f"{float(v):.1f}%"
        if kind == "pct2":     return f"{float(v):.2f}%"
        if kind == "decimal1": return f"{float(v):.1f}"
        if kind == "decimal2": return f"{float(v):.2f}"
        return str(v)
    except Exception:
        return "-"

def format_df_t3(df: pd.DataFrame) -> pd.DataFrame:
    out = df.copy().astype(object)
    for lbl in out.index:
        out.loc[lbl] = out.loc[lbl].apply(lambda v: _fmt_t3(v, _T3_ROW_FMT.get(lbl, "int")))
    return out

def _render_table3(df_fmt: pd.DataFrame) -> str:
    companies = list(df_fmt.columns)
    col_groups: list[list] = []
    for co in companies:
        g = T3_COMPANY_GROUP.get(co, "")
        if col_groups and col_groups[-1][0] == g:
            col_groups[-1][1].append(co)
        else:
            col_groups.append([g, [co]])
    S = {
        "tbl":     "border-collapse:collapse;width:100%;font-size:13px;font-family:sans-serif;",
        "th_idx":  "border:1px solid #ccc;padding:6px 10px;font-weight:700;white-space:nowrap;text-align:center;background:#F2F3F4;",
        "th_grp":  "border:1px solid #ccc;padding:6px 10px;font-weight:700;white-space:nowrap;text-align:center;background:#D6EAF8;",
        "th":      "border:1px solid #ccc;padding:6px 10px;font-weight:600;white-space:nowrap;text-align:center;background:#F2F3F4;",
        "td":      "border:1px solid #ccc;padding:5px 10px;text-align:right;white-space:nowrap;",
        "td_lbl":  "border:1px solid #ccc;padding:5px 10px;font-weight:600;text-align:left;white-space:nowrap;",
        "td_sec":  "border:1px solid #ccc;padding:5px 8px;font-weight:700;text-align:center;vertical-align:middle;background:#EBF5FB;white-space:nowrap;",
        "td_sub":  "border:1px solid #ccc;padding:5px 8px;font-weight:600;text-align:center;vertical-align:middle;background:#EDF4FC;white-space:nowrap;",
        "td_type": "border:1px solid #ccc;padding:5px 10px;text-align:left;white-space:nowrap;",
    }
    r1 = [f'<th rowspan="2" colspan="3" style="{S["th_idx"]}">구분</th>']
    r2 = []
    for g, cos in col_groups:
        r1.append(f'<th colspan="{len(cos)}" style="{S["th_grp"]}">{g}</th>')
        for co in cos:
            r2.append(f'<th style="{S["th"]}">{co}</th>')
    thead = f'<thead><tr>{"".join(r1)}</tr><tr>{"".join(r2)}</tr></thead>'

    RATIO_BG = "#FEF5E7"   # 임원 비중 행 (연한 황금색)
    GRAY_BG  = "#EAECEE"   # 평균급여·평균근속 행 (연한 회색)

    SIMPLE_LABELS = [
        ("자기자본(조원)",           "자기자본 (단위: 조원)",  "normal"),
        ("총인원",                   "총인원",                "normal"),
        ("임원 인원수",              "임원 인원수",           "normal"),
        ("임원 비중(%)",             "총인원 대비 임원 비중", "ratio"),
        ("리테일 직원수",            "리테일(영업) 직원수",   "normal"),
        ("리테일 비중(%)",           "총인원 대비 리테일 비중","normal"),
        ("리테일 제외 임원 비중(%)", "리테일 제외 임원 비중", "ratio"),
        ("평균급여(백만원)",         "평균급여 (단위: 백만원)","gray"),
        ("평균근속(년)",             "평균근속 (단위: 년)",   "gray_last"),
    ]
    INEON_STRUCTURE = [
        ("리테일",   ["리테일 정규직",   "리테일 계약직",   "리테일 소계"]),
        ("본사영업", ["본사영업 정규직", "본사영업 계약직", "본사영업 소계"]),
        ("본사관리", ["본사관리 정규직", "본사관리 계약직", "본사관리 소계"]),
        ("총계",     ["정규직 총합",     "계약직 총합",     "총합"]),
    ]
    TYPE_LABELS = {
        "리테일 정규직": "정규직", "리테일 계약직": "계약직", "리테일 소계": "소계",
        "본사영업 정규직": "정규직", "본사영업 계약직": "계약직", "본사영업 소계": "소계",
        "본사관리 정규직": "정규직", "본사관리 계약직": "계약직", "본사관리 소계": "소계",
        "정규직 총합": "정규직", "계약직 총합": "계약직", "총합": "총합",
    }
    INEON_TOTAL = sum(len(lbls) for _, lbls in INEON_STRUCTURE)
    body_rows = []
    for lbl, display, row_kind in SIMPLE_LABELS:
        if lbl not in df_fmt.index:
            continue
        if row_kind == "ratio":
            s_lbl  = f"border:1px solid #ccc;padding:5px 10px;font-weight:700;text-align:left;white-space:nowrap;background:{RATIO_BG};"
            s_cell = f"border:1px solid #ccc;padding:5px 10px;text-align:right;white-space:nowrap;background:{RATIO_BG};"
        elif row_kind == "gray":
            s_lbl  = f"border:1px solid #ccc;padding:5px 10px;font-weight:600;text-align:left;white-space:nowrap;background:{GRAY_BG};"
            s_cell = f"border:1px solid #ccc;padding:5px 10px;text-align:right;white-space:nowrap;background:{GRAY_BG};"
        elif row_kind == "gray_last":
            s_lbl  = f"border:1px solid #ccc;border-bottom:3px double #888;padding:5px 10px;font-weight:600;text-align:left;white-space:nowrap;background:{GRAY_BG};"
            s_cell = f"border:1px solid #ccc;border-bottom:3px double #888;padding:5px 10px;text-align:right;white-space:nowrap;background:{GRAY_BG};"
        else:
            s_lbl  = S["td_lbl"]
            s_cell = S["td"]
        cells = [f'<td colspan="3" style="{s_lbl}">{display}</td>']
        cells += [f'<td style="{s_cell}">{df_fmt.loc[lbl, co]}</td>' for co in companies]
        body_rows.append(f'<tr>{"".join(cells)}</tr>')
    first_ineon = True
    for sub_sec, lbls in INEON_STRUCTURE:
        for j, lbl in enumerate(lbls):
            if lbl not in df_fmt.index:
                continue
            cells = []
            if first_ineon and j == 0:
                cells.append(f'<td rowspan="{INEON_TOTAL}" style="{S["td_sec"]}">인원</td>')
                first_ineon = False
            if j == 0:
                cells.append(f'<td rowspan="3" style="{S["td_sub"]}">{sub_sec}</td>')
            cells.append(f'<td style="{S["td_type"]}">{TYPE_LABELS.get(lbl, lbl)}</td>')
            cells += [f'<td style="{S["td"]}">{df_fmt.loc[lbl, co]}</td>' for co in companies]
            body_rows.append(f'<tr>{"".join(cells)}</tr>')
    tbody = f'<tbody>{"".join(body_rows)}</tbody>'
    return f'<div style="overflow-x:auto;"><table style="{S["tbl"]}">{thead}{tbody}</table></div>'


# ── 데이터 가공 ───────────────────────────────────────────────────────────────
def _mark_partial(df: pd.DataFrame, fin_data: dict) -> pd.DataFrame:
    df = df.copy()
    missing_by_co: dict[str, list[int]] = {}
    for company in COMPANIES:
        missing = [
            y for y in TREND_YEARS
            if fin_data.get(company, {}).get(y) is not None
            and fin_data.get(company, {}).get(y, {}).get("판관비") is None
        ]
        if missing:
            missing_by_co[company] = missing

    def _apply(col, missing_years):
        if col not in df.columns:
            return
        for y in missing_years:
            lbl = f"순영업수익  {y}년 (억원)"
            if lbl in df.index and df.loc[lbl, col] != "-":
                df.loc[lbl, col] += "*"
        avg_lbl = "순영업수익  평균 (억원)"
        if avg_lbl in df.index and df.loc[avg_lbl, col] != "-":
            df.loc[avg_lbl, col] += "*"
        if RECENT_YEAR in missing_years:
            prod_lbl = "인당생산성 (백만원/인)"
            if prod_lbl in df.index and df.loc[prod_lbl, col] != "-":
                df.loc[prod_lbl, col] += "*"

    for co, missing in missing_by_co.items():
        _apply(co, missing)
    for g, members in GROUPS_DEF.items():
        all_missing = sorted({y for co in members for y in missing_by_co.get(co, [])})
        if all_missing:
            _apply(AVG_COL_NAME[g], all_missing)
    return df


def _apply_woori_override(t1_recent, t1_prev, t2, ov):
    co = "우리투자증권"
    for raw, yr in [(t1_recent, RECENT_YEAR), (t1_prev, PREV_YEAR)]:
        d = ov.get(yr, {})
        if d.get("자기자본") is not None:
            raw.loc["자기자본(억원)", co] = float(d["자기자본"])
        if d.get("총인원") is not None:
            raw.loc["총인원", co] = int(d["총인원"])
        total_val = raw.loc["총인원", co]
        total_ok = (total_val is not None
                    and not (isinstance(total_val, float) and pd.isna(total_val))
                    and total_val > 0)
        for sec in ("리테일", "본사영업", "본사관리"):
            reg = d.get(f"{sec}_정규직")
            cnt = d.get(f"{sec}_기간제")
            if reg is None and cnt is None:
                continue
            reg = reg or 0
            cnt = cnt or 0
            sec_tot = reg + cnt
            if sec_tot > 0:
                raw.loc[f"{sec} 인원", co] = sec_tot
                raw.loc[f"{sec} 비중(%)", co] = (
                    round(sec_tot / total_val * 100, 1) if total_ok else None)
            raw.loc[f"{sec} 정규직 인원", co] = reg if reg > 0 else None
            raw.loc[f"{sec} 기간제 인원", co] = cnt if cnt > 0 else None
            raw.loc[f"{sec} 정규직 비중(%)", co] = (
                round(reg / sec_tot * 100, 1) if sec_tot > 0 else None)
            raw.loc[f"{sec} 기간제 비중(%)", co] = (
                round(cnt / sec_tot * 100, 1) if sec_tot > 0 else None)
    for yr in TREND_YEARS:
        d = ov.get(yr, {})
        if d.get("총인원") is not None:
            t2.loc[f"임직원수 {yr}년", co] = int(d["총인원"])
        if d.get("순영업수익") is not None:
            t2.loc[f"순영업수익 {yr}년(억원)", co] = float(d["순영업수익"])
    def _avg2(vals):
        valid = [v for v in vals
                 if v is not None and not (isinstance(v, float) and pd.isna(v)) and v != 0]
        return round(sum(valid) / len(valid), 0) if valid else None
    emp_vals = [t2.loc[f"임직원수 {y}년", co] for y in TREND_YEARS]
    nor_vals = [t2.loc[f"순영업수익 {y}년(억원)", co] for y in TREND_YEARS]
    emp_avg  = _avg2(emp_vals)
    nor_avg  = _avg2(nor_vals)
    t2.loc["임직원수 평균", co]         = emp_avg
    t2.loc["순영업수익 평균(억원)", co] = nor_avg
    if emp_avg and nor_avg:
        t2.loc["인당생산성(백만원/인)", co] = round(float(nor_avg) * 100 / float(emp_avg), 1)
    return t1_recent, t1_prev, t2


# ── 데이터 로드 (캐싱) ────────────────────────────────────────────────────────
@st.cache_data(show_spinner="DART에서 데이터를 불러오는 중...")
def _load(force_refresh: bool = False):
    emp, fin, src = fetch_all_raw(use_cache=not force_refresh)
    return emp, fin, src

@st.cache_data(show_spinner="임원인력 데이터를 불러오는 중...")
def _load_t3(force_refresh: bool = False):
    t3e, t3f, t3x, src = fetch_t3_raw(use_cache=not force_refresh)
    return t3e, t3f, t3x, src


# ── 사이드바 & 디버그 ─────────────────────────────────────────────────────────
def _refresh_password() -> str:
    try:
        return st.secrets.get("REFRESH_PASSWORD", "")
    except Exception:
        return os.getenv("REFRESH_PASSWORD", "")


def render_sidebar() -> dict:
    """사이드바를 렌더링하고 woori_ov 딕셔너리를 반환한다."""
    with st.sidebar:
        st.title("설정")

        with st.expander("데이터 재수집 (관리자)"):
            correct_pw = _refresh_password()
            if correct_pw:
                pw_input = st.text_input("비밀번호", type="password", key="refresh_pw")
                unlocked = pw_input == correct_pw
                if pw_input and not unlocked:
                    st.error("비밀번호가 틀렸습니다.")
            else:
                unlocked = True
            if unlocked:
                if st.button("DART 데이터 재수집", type="primary", width="stretch"):
                    _load.clear()
                    _load_t3.clear()
                    if CACHE_PATH.exists():
                        CACHE_PATH.unlink()
                    st.rerun()

        st.divider()
        st.markdown(f"""
**기준연도**: {RECENT_YEAR}년
**추이 기간**: {TREND_YEARS[-1]}~{TREND_YEARS[0]}년
**재무제표**: 개별(OFS)
**출처**: DART OpenAPI
""")
        st.divider()
        st.markdown("""
**데이터 유의사항**
- 우리투자증권: 2024년이 첫 사업보고서 (이전 연도 N/A)
- 메리츠증권: 판관비 미공시 → 순영업수익 N/A
- 순영업수익 2022년: DART API 미제공 (전사 N/A)
- 우리투자증권 섹션 구분: DART 미신고 → 아래 직접입력으로 반영 가능
""")
        st.divider()
        woori_ov: dict[int, dict] = {}
        with st.expander("우리투자증권 직접입력"):
            st.caption("0 입력 시 DART 데이터 그대로 사용")
            for y in TREND_YEARS:
                st.markdown(f"**{y}년**")
                d: dict = {}
                if y in (RECENT_YEAR, PREV_YEAR):
                    v = st.number_input("자기자본 (억원)", min_value=0, value=0,
                                        step=100, key=f"ov_eq_{y}")
                    d["자기자본"] = int(v) if v else None
                v = st.number_input("총인원 (명)", min_value=0, value=0,
                                    step=10, key=f"ov_hc_{y}")
                d["총인원"] = int(v) if v else None
                v = st.number_input("순영업수익 (억원)", min_value=0, value=0,
                                    step=10, key=f"ov_nr_{y}")
                d["순영업수익"] = float(v) if v else None
                if y in (RECENT_YEAR, PREV_YEAR):
                    st.caption("섹션별 인원 (정규직 / 기간제)")
                    for sec in ("리테일", "본사영업", "본사관리"):
                        st.markdown(f"<small style='color:#555;'>**{sec}**</small>",
                                    unsafe_allow_html=True)
                        c1, c2 = st.columns(2)
                        with c1:
                            vr = st.number_input("정규직", min_value=0, value=0, step=1,
                                                 key=f"ov_{sec}_reg_{y}",
                                                 label_visibility="visible")
                            d[f"{sec}_정규직"] = int(vr) if vr else None
                        with c2:
                            vc = st.number_input("기간제", min_value=0, value=0, step=1,
                                                 key=f"ov_{sec}_cnt_{y}",
                                                 label_visibility="visible")
                            d[f"{sec}_기간제"] = int(vc) if vc else None
                woori_ov[y] = d
    return woori_ov


def render_debug(emp_data: dict, t3_emp: dict, src_main: str, src_t3: str):
    """디버그 expander를 렌더링한다."""
    with st.expander("🔍 디버그", expanded=False):
        from dart_api import _get_api_key
        key = _get_api_key()
        st.write("**API 키**:", f"길이 {len(key)} / 앞 4자: `{key[:4] if key else '(없음)'}`")
        st.write("**캐시 파일**:", str(CACHE_PATH))
        st.write("**캐시 존재**:", CACHE_PATH.exists())
        if CACHE_PATH.exists():
            mtime = CACHE_PATH.stat().st_mtime
            st.write("**캐시 최종 수정**:",
                     datetime.datetime.fromtimestamp(mtime).strftime("%Y-%m-%d %H:%M:%S"))
        st.write("**기본 테이블 데이터 출처**:",
                 "✅ 캐시" if src_main == "캐시" else "🔄 DART 실시간 수집")
        st.write("**임원인력(T3) 데이터 출처**:",
                 "✅ 캐시" if src_t3 == "캐시" else "🔄 DART 실시간 수집")
        st.write(f"**기본 테이블 회사 수**: {len(emp_data)}개  —  {', '.join(emp_data.keys())}")
        st.write(f"**T3 테이블 회사 수**: {len(t3_emp)}개  —  {', '.join(t3_emp.keys())}")
