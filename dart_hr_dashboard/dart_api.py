import os
import re
import time
from pathlib import Path
import requests
from dotenv import load_dotenv
from bs4 import BeautifulSoup

load_dotenv(Path(__file__).parent / ".env")

def _get_api_key() -> str:
    # Streamlit Cloud 환경: st.secrets 우선 사용
    try:
        import streamlit as st
        val = st.secrets.get("DART_API_KEY", "")
        if val:
            return val
    except Exception:
        pass
    # 로컬 환경: .env / 환경변수
    # streamlit import가 os.environ을 초기화하는 경우를 대비해 .env 재로드
    key = os.getenv("DART_API_KEY", "")
    if not key:
        load_dotenv(Path(__file__).parent / ".env", override=True)
        key = os.getenv("DART_API_KEY", "")
    return key

API_KEY = _get_api_key()
BASE_URL = "https://opendart.fss.or.kr/api"

# corp_code는 DART XML 전체 목록에서 확인한 정확한 값 사용
# 사용자 제공 코드 중 일부가 달라 수정함:
#   한국투자증권 00134550→00160144, KB증권 00164377→00164876(케이비증권)
#   키움증권 00296014→00296290, 메리츠증권 00080422→00163682
#   유안타증권 00108320→00117601, 현대차증권 00146377→00137997
#   IBK투자증권 00657518→00684918(아이비케이투자증권)
#   유진투자증권 00137997→00131054(유진증권)
#   우리투자증권 00283707→01015364 (2024년이 첫 사업보고서)
COMPANIES = {
    "우리투자증권": "01015364",
    "미래에셋증권": "00111722",
    "한국투자증권": "00160144",
    "NH투자증권":   "00120182",
    "KB증권":       "00164876",
    "키움증권":     "00296290",
    "메리츠증권":   "00163682",
    "유안타증권":   "00117601",
    "현대차증권":   "00137997",
    "IBK투자증권":  "00684918",
    "유진투자증권": "00131054",
}

# 우리투자증권은 2024년이 첫 사업보고서이므로 2024~2020 조회
# 나머지 회사는 2023~2019 조회 (2024 사보는 2025년 3월 제출 → 대부분 있음)
TARGET_YEARS = [2025, 2024, 2023, 2022, 2021]


def fetch_employee_info(corp_code: str, year: int) -> list[dict]:
    """
    DART empSttus API로 직원 현황을 조회한다.
    반환: row 리스트 (성별×사업부문 단위로 분리돼 있음)
    """
    url = f"{BASE_URL}/empSttus.json"
    params = {
        "crtfc_key": API_KEY,
        "corp_code": corp_code,
        "bsns_year": str(year),
        "reprt_code": "11011",  # 사업보고서 고정 코드
    }
    try:
        resp = requests.get(url, params=params, timeout=10)
        resp.raise_for_status()
        data = resp.json()

        if data.get("status") == "000":
            return data.get("list", [])

        # 010/013: 데이터 없음 (정상 케이스)
        if data.get("status") not in ("010", "013"):
            print(f"[API 오류] corp={corp_code}, year={year}, "
                  f"status={data.get('status')}, msg={data.get('message')}")
        return []
    except Exception as e:
        print(f"[통신 오류] corp={corp_code}, year={year}: {e}")
        return []


def fetch_financial_stmt(corp_code: str, year: int, delay: float = 0.3) -> tuple[dict | None, dict | None]:
    """
    개별재무제표(OFS)에서 자기자본·세전이익·영업외손익·판관비를 조회한다.
    연 1회 API 호출로 당기(year)와 전기(year-1) 데이터를 동시에 반환한다.
    반환: (current_year_dict, prior_year_dict)  — 데이터 없으면 (None, None)

    계정명은 회사마다 다르므로 공백 제거 후 부분일치로 매칭한다.
    영업외손익이 직접 공시되지 않는 경우 영업외수익 - 영업외비용으로 산출한다.
    """
    url = f"{BASE_URL}/fnlttSinglAcntAll.json"
    params = {
        "crtfc_key": API_KEY,
        "corp_code": corp_code,
        "bsns_year": str(year),
        "reprt_code": "11011",
        "fs_div": "OFS",
    }
    _empty = lambda: {
        "자기자본": None, "세전이익": None,
        "영업외손익": None, "판관비": None,
        "_영업외수익": None, "_영업외비용": None,
    }
    try:
        resp = requests.get(url, params=params, timeout=10)
        resp.raise_for_status()
        data = resp.json()
        time.sleep(delay)

        if data.get("status") != "000":
            return None, None

        cur, prv = _empty(), _empty()
        rows_list = data.get("list", [])

        for row in rows_list:
            nm   = row.get("account_nm", "").strip()
            nm_n = "".join(nm.split()).lower()          # 공백 제거 + 소문자
            sj   = row.get("sj_div", "")
            th   = _to_int(row.get("thstrm_amount", ""))
            frm  = _to_int(row.get("frmtrm_amount", ""))

            # 자기자본: BS 자본총계 — "자    본    총    계" 처럼 내부 공백이 있는 회사도 처리
            if sj == "BS" and nm_n == "자본총계":
                if th is not None and abs(th) > abs(cur["자기자본"] or 0):
                    cur["자기자본"] = th
                    prv["자기자본"] = frm

            # 세전이익: "법인세" + ("순이익" or "순손익") 포함 → 다양한 표기 통일
            elif sj in ("CIS", "IS") and "법인세" in nm_n and any(k in nm_n for k in ("순이익", "순손익")):
                if cur["세전이익"] is None:
                    cur["세전이익"] = th
                    prv["세전이익"] = frm

            # 판관비: IS/CIS에서 우선 탐색 (판매관리비 / 판매비와관리비 등 모두 처리)
            elif sj in ("CIS", "IS") and "판매" in nm_n and "관리비" in nm_n:
                if cur["판관비"] is None:
                    cur["판관비"] = th
                    prv["판관비"] = frm

            # 영업외손익: 직접 공시 (일부 회사만)
            elif sj in ("CIS", "IS") and nm_n == "영업외손익":
                if cur["영업외손익"] is None:
                    cur["영업외손익"] = th
                    prv["영업외손익"] = frm

            # 영업외수익/비용 총계 (직접 공시 없을 때 계산용)
            # 기타영업외수익 같은 하위 항목은 "기타" 필터로 제외
            elif sj in ("CIS", "IS") and "영업외수익" in nm_n and "기타" not in nm_n:
                if cur["_영업외수익"] is None:
                    cur["_영업외수익"] = th
                    prv["_영업외수익"] = frm

            elif sj in ("CIS", "IS") and "영업외비용" in nm_n and "기타" not in nm_n:
                if cur["_영업외비용"] is None:
                    cur["_영업외비용"] = th
                    prv["_영업외비용"] = frm

        # 판관비 2차 탐색: IS/CIS에서 못 찾은 경우 주석 등 전체 섹션에서 재탐색
        # (메리츠증권 등 재무제표 주석에만 판관비를 기재하는 회사 대응)
        if cur["판관비"] is None:
            for row in rows_list:
                nm_n = "".join(row.get("account_nm", "").strip().split()).lower()
                sj   = row.get("sj_div", "")
                if sj not in ("CIS", "IS") and "판매" in nm_n and "관리비" in nm_n:
                    th  = _to_int(row.get("thstrm_amount", ""))
                    frm = _to_int(row.get("frmtrm_amount", ""))
                    if th is not None:
                        cur["판관비"] = th
                        prv["판관비"] = frm
                        break

        # 영업외손익 미공시 회사: 수익 - 비용으로 산출
        for d in (cur, prv):
            if d["영업외손익"] is None:
                rev = d["_영업외수익"]
                exp = d["_영업외비용"]
                if rev is not None and exp is not None:
                    d["영업외손익"] = rev - exp
            del d["_영업외수익"]
            del d["_영업외비용"]

        return cur, prv
    except Exception as e:
        print(f"[재무 오류] corp={corp_code}, year={year}: {e}")
        return None, None


def _to_int(s) -> int | None:
    if not s or str(s).strip() in ("-", ""):
        return None
    try:
        return int(str(s).replace(",", "").strip())
    except ValueError:
        return None


def fetch_all_companies(years: list[int] = TARGET_YEARS, delay: float = 0.3) -> dict:
    """
    모든 증권사 × 연도 조합의 직원 현황을 수집한다.

    반환 형식:
    {
        "우리투자증권": {2024: [...rows], 2023: [...rows], ...},
        "미래에셋증권": {2024: [...rows], ...},
        ...
    }
    """
    result = {}
    for company, corp_code in COMPANIES.items():
        result[company] = {}
        for year in years:
            rows = fetch_employee_info(corp_code, year)
            result[company][year] = rows
            status = f"{len(rows)}건" if rows else "없음"
            print(f"  [{company}] {year}년: {status}")
            time.sleep(delay)
    return result


def fetch_pangwanbi_from_html(corp_code: str, bsns_year: int, delay: float = 0.5) -> dict[int, int]:
    """
    DART 영업보고서 HTML의 '손익현황' 표에서 판매관리비를 추출한다.
    fnlttSinglAcntAll API에 판관비가 없는 회사(메리츠증권 등 금융업)에서 사용.

    반환: {year: amount_in_won}  (억원 단위 표기 → 원으로 변환)
    """
    _html_headers = {"User-Agent": "Mozilla/5.0"}

    # 1. 사업보고서 접수번호 조회
    try:
        resp = requests.get(f"{BASE_URL}/list.json", params={
            "crtfc_key": API_KEY,
            "corp_code": corp_code,
            "bgn_de": f"{bsns_year}0101",
            "end_de": f"{bsns_year + 1}0630",
            "pblntf_ty": "A",
            "page_count": 5,
        }, timeout=10)
        items = resp.json().get("list", [])
    except Exception as e:
        print(f"[HTML판관비] list.json 오류: {e}")
        return {}

    rcept_no = None
    for item in items:
        if "사업보고서" in item.get("report_nm", ""):
            rcept_no = item["rcept_no"]
            break
    if not rcept_no:
        return {}
    time.sleep(delay)

    # 2. 영업보고서 dcmNo 조회 (DART 공시 메인 페이지 파싱)
    try:
        main_resp = requests.get(
            "https://dart.fss.or.kr/dsaf001/main.do",
            params={"rcpNo": rcept_no},
            headers=_html_headers, timeout=10,
        )
        main_soup = BeautifulSoup(main_resp.text, "html.parser")
    except Exception as e:
        print(f"[HTML판관비] main.do 오류: {e}")
        return {}

    dcm_no = None
    for opt in main_soup.find_all("option"):
        if "영업보고서" in opt.get_text():
            m = re.search(r"dcmNo=(\d+)", opt.get("value", ""))
            if m:
                dcm_no = m.group(1)
                break
    if not dcm_no:
        return {}
    time.sleep(delay)

    # 3. 영업보고서 HTML 다운로드
    try:
        doc_resp = requests.get(
            "https://dart.fss.or.kr/report/viewer.do",
            params={"rcpNo": rcept_no, "dcmNo": dcm_no, "eleId": "0",
                    "offset": "0", "length": "0", "dtd": "dart6.dtd"},
            headers=_html_headers, timeout=20,
        )
        doc_soup = BeautifulSoup(doc_resp.text, "html.parser")
    except Exception as e:
        print(f"[HTML판관비] viewer.do 오류: {e}")
        return {}

    # 4. '판매관리비' 행이 있는 테이블 파싱
    # 열 순서: [구분, 제N기(=bsns_year), 제N-1기, 제N-2기]  (억원 단위)
    for table in doc_soup.find_all("table"):
        if "판매관리비" not in table.get_text():
            continue
        rows = table.find_all("tr")
        if not rows:
            continue

        n_data_cols = len(rows[0].find_all(["td", "th"])) - 1
        col_years = {i + 1: bsns_year - i for i in range(n_data_cols)}

        for row in rows:
            cells = [td.get_text(strip=True) for td in row.find_all(["td", "th"])]
            if not cells:
                continue
            label = "".join(cells[0].split()).replace("\xa0", "")
            if "판매관리비" not in label and "판매비" not in label:
                continue
            result = {}
            for col_idx, year in col_years.items():
                if col_idx >= len(cells):
                    continue
                val_str = cells[col_idx].replace(",", "").strip()
                try:
                    result[year] = int(val_str) * 10 ** 8
                except ValueError:
                    pass
            if result:
                print(f"[HTML판관비] {corp_code} {bsns_year}보고서 → {result}")
                return result

    return {}


if __name__ == "__main__":
    print("=== DART API 전체 수집 테스트 ===")
    data = fetch_all_companies()
    total = sum(len(rows) for co in data.values() for rows in co.values())
    print(f"\n총 수집 row 수: {total}")
