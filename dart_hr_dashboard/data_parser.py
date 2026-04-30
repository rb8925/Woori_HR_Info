import json
import time
from pathlib import Path

import pandas as pd

from dart_api import COMPANIES, fetch_employee_info, fetch_financial_stmt, fetch_pangwanbi_from_html

# ── 상수 ──────────────────────────────────────────────────────────────────────
RECENT_YEAR = 2025
PREV_YEAR   = 2024
TREND_YEARS = [2025, 2024, 2023, 2022]   # 가장 최근 → 가장 오래된 순

CACHE_PATH = Path("data/raw_cache.json")

# ── 사업부 섹션 매핑 ──────────────────────────────────────────────────────────
# 값: 해당 카테고리에 속하는 fo_bbm 패턴 목록 (공백 정규화 후 부분일치)
# None이면 섹션 구분 없음 (전체만 표시)
SECTION_MAP: dict[str, dict[str, list[str]] | None] = {
    "우리투자증권": None,                           # DART 제출상 섹션 구분 없음
    "미래에셋증권": {
        "리테일":  ["Retail"],
        "본사영업": ["본사영업"],
        "본사관리": ["관리/지원"],
    },
    "한국투자증권": {
        "리테일":  ["Retail영업"],
        "본사영업": ["본사영업"],
        "본사관리": ["관리지원", "기타"],             # 관리지원(Retail,본사지원) + 기타
    },
    "NH투자증권": {
        "리테일":  ["PWM사업부", "WM사업부"],        # 2023: PWM사업부, 2024: WM사업부로 변경됨
        "본사영업": ["본사영업"],
        "본사관리": ["본사지원"],
    },
    "KB증권": {
        "리테일":  ["Retail"],
        "본사영업": ["본사영업"],
        "본사관리": ["본사관리"],
    },
    "키움증권": {
        "리테일":  [],                              # 리테일 구분 없음
        "본사영업": ["위탁매매", "자기매매", "인수업무"],
        "본사관리": ["지원부문"],
    },
    "메리츠증권": {
        "리테일":  ["지점"],                        # "지      점" → 정규화 후 "지점"
        "본사영업": ["본사영업"],
        "본사관리": ["본사관리"],
    },
    "유안타증권": {
        "리테일":  ["지점"],
        "본사영업": ["본사영업"],
        "본사관리": ["본사관리"],
    },
    "현대차증권": {
        "리테일":  ["지점"],
        "본사영업": ["본사영업"],
        "본사관리": ["본사관리"],
    },
    "IBK투자증권": {
        "리테일":  ["지점영업"],
        "본사영업": ["본사영업"],
        "본사관리": ["본사관리"],
    },
    "유진투자증권": {
        "리테일":  ["WM"],
        "본사영업": ["본점영업"],
        "본사관리": ["본사지원"],
    },
}

# ── 유틸리티 ──────────────────────────────────────────────────────────────────

def _norm(s: str) -> str:
    """공백(스페이스, 탭, 줄바꿈 등) 제거 후 소문자 통일."""
    return "".join(str(s).split()).lower()


def _match(fo_bbm: str, patterns: list[str]) -> bool:
    """공백 정규화 후 패턴이 fo_bbm에 포함되어 있으면 True."""
    n = _norm(fo_bbm)
    return any(_norm(p) in n for p in patterns)


def _safe_int(s) -> int:
    if s is None or str(s).strip() in ("-", ""):
        return 0
    try:
        return int(str(s).replace(",", "").strip())
    except ValueError:
        return 0


def _avg(values: list) -> float | None:
    valid = [v for v in values if v is not None and v != 0]
    return round(sum(valid) / len(valid), 0) if valid else None


def _억(value: int | None) -> float | None:
    """원 단위 → 억원 (소수점 1자리)."""
    if value is None:
        return None
    return round(value / 1e8, 1)

# ── API 수집 & 캐싱 ──────────────────────────────────────────────────────────

def _load_cache() -> dict:
    if CACHE_PATH.exists():
        return json.loads(CACHE_PATH.read_text(encoding="utf-8"))
    return {}


def _save_cache(data: dict):
    CACHE_PATH.parent.mkdir(exist_ok=True)
    CACHE_PATH.write_text(json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8")


def fetch_all_raw(use_cache: bool = True) -> tuple[dict, dict]:
    """
    모든 회사의 직원현황(emp)과 재무데이터(fin)를 수집한다.

    캐시가 있으면 API 호출 없이 반환하고, 없으면 DART에서 가져와 저장한다.

    반환:
        emp_data[company][year] = [rows...]
        fin_data[company][year] = {"자기자본": int, "세전이익": int, "영업외손익": int, "판관비": int}
    """
    cache = _load_cache() if use_cache else {}
    if cache.get("emp_data") and cache.get("fin_data"):
        first_co = next(iter(COMPANIES))
        if str(TREND_YEARS[0]) in cache["emp_data"].get(first_co, {}):
            print("캐시에서 데이터 로드")
            emp_raw = cache["emp_data"]
            fin_raw = cache["fin_data"]
            emp_data = {co: {int(y): rows for y, rows in yrs.items()} for co, yrs in emp_raw.items()}
            fin_data = {co: {int(y): fin  for y, fin  in yrs.items()} for co, yrs in fin_raw.items()}
            return emp_data, fin_data
        print("캐시에 최신 연도 데이터 없음 — 재수집")

    emp_data: dict[str, dict[int, list]] = {}
    fin_data: dict[str, dict[int, dict]] = {}

    for company, corp_code in COMPANIES.items():
        print(f"\n[{company}] 수집 중...")
        emp_data[company] = {}
        fin_data[company] = {}

        # 직원 현황: 연도별 개별 조회
        for year in TREND_YEARS:
            rows = fetch_employee_info(corp_code, year)
            emp_data[company][year] = rows
            print(f"  직원 {year}: {len(rows)}건")
            time.sleep(0.25)

        # 재무: bsns_year=2025 → 2025(당기)+2024(전기)
        #        bsns_year=2024 → 2024(당기)+2023(전기)
        #        bsns_year=2023 → 2023(당기)+2022(전기)
        # bsns_year=2022 이전은 DART API에서 013(데이터없음)을 반환해 미제공
        for query_year, (cur_yr, prv_yr) in [(2025, (2025, 2024)), (2024, (2024, 2023)), (2023, (2023, 2022))]:
            cur, prv = fetch_financial_stmt(corp_code, query_year)
            if cur:
                fin_data[company][cur_yr] = cur
            if prv:
                fin_data[company][prv_yr] = prv
            print(f"  재무 {query_year}→ cur={'OK' if cur else '-'}, prv={'OK' if prv else '-'}")

        # 판관비 HTML 폴백: fnlttSinglAcntAll에 판관비가 없는 회사(메리츠증권 등)는
        # 영업보고서 HTML에서 판매관리비를 추출한다.
        # bsns_year=2024 보고서 한 번으로 2024·2023·2022 세 연도 값을 커버.
        missing_years = [y for y in TREND_YEARS if fin_data[company].get(y, {}).get("판관비") is None]
        if missing_years:
            html_pg = fetch_pangwanbi_from_html(corp_code, TREND_YEARS[1])  # 2024 사보
            for yr, amount in html_pg.items():
                if yr in fin_data[company] and fin_data[company][yr].get("판관비") is None:
                    fin_data[company][yr]["판관비"] = amount
            if html_pg:
                covered = [y for y in missing_years if y in html_pg]
                print(f"  HTML판관비 보완: {covered}")

    _save_cache({"emp_data": {
        co: {str(y): rows for y, rows in yrs.items()}
        for co, yrs in emp_data.items()
    }, "fin_data": {
        co: {str(y): fin for y, fin in yrs.items()}
        for co, yrs in fin_data.items()
    }})
    print("\n캐시 저장 완료")
    return emp_data, fin_data

# ── 집계 함수 ─────────────────────────────────────────────────────────────────

def _headcount_by_section(rows: list[dict], sec_map: dict | None) -> dict:
    """
    rows: empSttus API rows (성별×섹션 단위)
    반환: {"리테일": n|None, "본사영업": n|None, "본사관리": n|None, "총인원": n}
    """
    total = sum(_safe_int(r.get("sm", 0)) for r in rows)

    if sec_map is None:
        return {"리테일": None, "본사영업": None, "본사관리": None, "총인원": total}

    counts = {cat: (0 if patterns else None) for cat, patterns in sec_map.items()}
    counts.setdefault("리테일", None)
    counts.setdefault("본사영업", None)
    counts.setdefault("본사관리", None)
    for row in rows:
        bbm = row.get("fo_bbm", "")
        sm  = _safe_int(row.get("sm", 0))
        matched = False
        for cat, patterns in sec_map.items():
            if _match(bbm, patterns):
                counts[cat] += sm
                matched = True
                break
        if not matched:
            # 매핑 안 된 섹션은 경고만 출력 (데이터 손실 방지)
            print(f"  [매핑 없음] fo_bbm='{bbm}', sm={sm}")

    return {**counts, "총인원": total}


def _net_op_revenue(fin: dict | None) -> int | None:
    """순영업수익 = 세전이익 − 영업외손익 + 판관비 (판관비 없으면 제외 계산)"""
    if not fin:
        return None
    a, b = fin.get("세전이익"), fin.get("영업외손익")
    if None in (a, b):
        return None
    c = fin.get("판관비")
    # 판관비는 비용이므로 항상 양수여야 함 — 음수로 파싱된 경우 절대값 사용
    if c is not None and c < 0:
        c = abs(c)
    return a - b + (c if c is not None else 0)

# ── 테이블 빌더 ───────────────────────────────────────────────────────────────

def build_table1(emp_data: dict, fin_data: dict, year: int = RECENT_YEAR) -> pd.DataFrame:
    """
    Table 1 — 회사별 스냅샷 (기준연도: year)

    행: 자기자본, 리테일인원, 리테일비중, 본사영업인원, 본사영업비중,
        본사관리인원, 본사관리비중, 총인원
    열: 회사명
    """
    rows = {}
    for company in COMPANIES:
        emp_rows = emp_data.get(company, {}).get(year, [])
        fin       = fin_data.get(company, {}).get(year)
        sec_map   = SECTION_MAP.get(company)

        hc    = _headcount_by_section(emp_rows, sec_map)
        total = hc["총인원"]

        def pct(n):
            return round(n / total * 100, 1) if (n is not None and total > 0) else None

        rows[company] = {
            "자기자본(억원)":     _억(fin.get("자기자본") if fin else None),
            "리테일 인원":        hc["리테일"],
            "리테일 비중(%)":     pct(hc["리테일"]),
            "본사영업 인원":      hc["본사영업"],
            "본사영업 비중(%)":   pct(hc["본사영업"]),
            "본사관리 인원":      hc["본사관리"],
            "본사관리 비중(%)":   pct(hc["본사관리"]),
            "총인원":             total,
        }

    # 열 = 회사, 행 = 지표
    return pd.DataFrame(rows)


def build_table2(emp_data: dict, fin_data: dict) -> pd.DataFrame:
    """
    Table 2 — 4개년 추이 비교

    행:  인당생산성, 임직원수(각 연도 + 평균), 순영업수익(각 연도 + 평균)
    열:  회사명
    """
    rows = {}
    for company in COMPANIES:
        emp_by_yr = {}
        nor_by_yr = {}

        for year in TREND_YEARS:
            emp_rows = emp_data.get(company, {}).get(year, [])
            total    = sum(_safe_int(r.get("sm", 0)) for r in emp_rows)
            emp_by_yr[year] = total if total > 0 else None

            fin = fin_data.get(company, {}).get(year)
            nor_by_yr[year] = _net_op_revenue(fin)

        emp_avg = _avg(list(emp_by_yr.values()))
        nor_avg = _avg(list(nor_by_yr.values()))

        productivity = (
            round(nor_avg / emp_avg / 1e6, 1)
            if nor_avg and emp_avg else None
        )

        y0, y1, y2, y3 = TREND_YEARS
        rows[company] = {
            "인당생산성(백만원/인)":   productivity,
            f"임직원수 {y0}년":        emp_by_yr[y0],
            f"임직원수 {y1}년":        emp_by_yr[y1],
            f"임직원수 {y2}년":        emp_by_yr[y2],
            f"임직원수 {y3}년":        emp_by_yr[y3],
            "임직원수 평균":           emp_avg,
            f"순영업수익 {y0}년(억원)": _억(nor_by_yr[y0]),
            f"순영업수익 {y1}년(억원)": _억(nor_by_yr[y1]),
            f"순영업수익 {y2}년(억원)": _억(nor_by_yr[y2]),
            f"순영업수익 {y3}년(억원)": _억(nor_by_yr[y3]),
            "순영업수익 평균(억원)":    _억(nor_avg),
        }

    return pd.DataFrame(rows)


# ── 단독 실행 테스트 ──────────────────────────────────────────────────────────

if __name__ == "__main__":
    print("=== 데이터 수집 시작 ===")
    emp, fin = fetch_all_raw(use_cache=False)

    print("\n=== Table 1 ===")
    t1 = build_table1(emp, fin)
    print(t1.to_string())

    print("\n=== Table 2 ===")
    t2 = build_table2(emp, fin)
    print(t2.to_string())
