"""행정규칙 연혁 웹 스크래퍼 — 법제처 API 미지원 시 fallback

법제처 API의 lawHistory.do가 admrul에 대해 동작하지 않는 경우,
law.go.kr 웹 인터페이스에서 행정규칙 연혁 목록을 스크래핑한다.

대상 URL:
  https://www.law.go.kr/admRulBylInfoR.do?admRulSeq={mst}
  → 행정규칙 시행일자별 버전 목록 테이블 파싱

주의:
  - law.go.kr 페이지 구조가 변경되면 selector 수정 필요
  - JavaScript 렌더링 페이지는 파싱 불가 (Selenium 없이는 한계)
  - 요청 간격 1초 권장 (서버 부하 방지)
"""
import logging
import time
from typing import Optional

import requests

logger = logging.getLogger(__name__)

_HEADERS = {
    "User-Agent": (
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
        "AppleWebKit/537.36 (KHTML, like Gecko) "
        "Chrome/124.0.0.0 Safari/537.36"
    ),
    "Accept-Language": "ko-KR,ko;q=0.9",
    "Referer": "https://www.law.go.kr/",
}

# 시도할 URL 패턴 목록 (mst = 법령MST = admRulSeq)
_HISTORY_URL_PATTERNS = [
    "https://www.law.go.kr/admRulBylInfoR.do?admRulSeq={mst}",
    "https://www.law.go.kr/admRulLstv2R.do?admRulSeq={mst}",
]


def _try_import_bs4():
    try:
        from bs4 import BeautifulSoup
        return BeautifulSoup
    except ImportError:
        return None


def _parse_date_str(s: str) -> Optional[str]:
    """'2024.01.01' 또는 '20240101' → '20240101' 형식으로 정규화"""
    s = s.strip().replace(".", "").replace("-", "").replace(" ", "")
    if len(s) == 8 and s.isdigit():
        return s
    return None


def scrape_admrul_history(mst: str, timeout: int = 15) -> list[dict]:
    """행정규칙 연혁 목록 스크래핑.

    Args:
        mst: 법령MST (admRulSeq)
        timeout: HTTP 요청 타임아웃 (초)

    Returns:
        [{"법령ID": str, "시행일자": str, "공포번호": str, "법령명한글": str}, ...]
        파싱 실패 시 빈 목록 반환
    """
    BeautifulSoup = _try_import_bs4()
    if BeautifulSoup is None:
        logger.warning("beautifulsoup4 미설치 — 스크래핑 불가 (pip install beautifulsoup4)")
        return []

    session = requests.Session()
    session.headers.update(_HEADERS)

    for url_template in _HISTORY_URL_PATTERNS:
        url = url_template.format(mst=mst)
        try:
            resp = session.get(url, timeout=timeout)
            resp.raise_for_status()
            resp.encoding = "utf-8"
            time.sleep(0.5)  # 서버 부하 방지

            soup = BeautifulSoup(resp.text, "html.parser")
            results = _parse_history_table(soup, mst)
            if results:
                logger.info(
                    "[스크래퍼] 연혁 %d건 파싱 성공: %s", len(results), url
                )
                return results

        except requests.exceptions.RequestException as exc:
            logger.debug("[스크래퍼] 요청 실패 (%s): %s", url, exc)
        except Exception as exc:
            logger.debug("[스크래퍼] 파싱 오류 (%s): %s", url, exc)

    logger.debug("[스크래퍼] 모든 URL 패턴 실패 (mst=%s)", mst)
    return []


def _parse_history_table(soup, mst: str) -> list[dict]:
    """HTML에서 연혁 테이블 파싱.

    law.go.kr 행정규칙 페이지의 테이블 구조:
    - 시행일자 / 공포번호 / 법령명 컬럼이 포함된 테이블
    - 각 행에 법령ID 링크(admRulSeq 또는 lsiSeq) 포함
    """
    results = []

    # 연혁 테이블 탐색 — 여러 선택자 시도
    tables = soup.find_all("table")
    for table in tables:
        rows = table.find_all("tr")
        if len(rows) < 2:
            continue

        # 헤더 행에서 컬럼 위치 파악
        header_row = rows[0]
        headers = [th.get_text(strip=True) for th in header_row.find_all(["th", "td"])]

        col_enforce = _find_col(headers, ["시행일", "시행일자"])
        col_announce = _find_col(headers, ["공포번호", "제정·개정번호", "번호"])
        col_name = _find_col(headers, ["행정규칙명", "법령명", "규칙명"])

        if col_enforce is None:
            continue  # 연혁 테이블이 아님

        for row in rows[1:]:
            cells = row.find_all(["td", "th"])
            if len(cells) <= max(
                c for c in [col_enforce, col_announce, col_name] if c is not None
            ):
                continue

            enforce_raw = cells[col_enforce].get_text(strip=True) if col_enforce is not None else ""
            announce_raw = cells[col_announce].get_text(strip=True) if col_announce is not None else ""
            name_raw = cells[col_name].get_text(strip=True) if col_name is not None else ""

            enforce_date = _parse_date_str(enforce_raw)
            if not enforce_date:
                continue

            # 법령ID: 링크 href에서 추출
            law_id = _extract_id_from_row(row)

            results.append({
                "법령ID": law_id or "",
                "법령MST": mst,
                "시행일자": enforce_date,
                "공포번호": announce_raw,
                "법령명한글": name_raw,
            })

    return results


def _find_col(headers: list[str], keywords: list[str]) -> Optional[int]:
    """헤더 목록에서 키워드가 포함된 컬럼 인덱스 반환"""
    for kw in keywords:
        for i, h in enumerate(headers):
            if kw in h:
                return i
    return None


def _extract_id_from_row(row) -> Optional[str]:
    """table row의 링크 href에서 법령ID(admRulSeq 또는 lsiSeq) 추출"""
    import re
    for a in row.find_all("a", href=True):
        href = a["href"]
        # admRulSeq=XXXXX 또는 lsiSeq=XXXXX 패턴
        m = re.search(r'(?:admRulSeq|lsiSeq|ID)=(\d+)', href)
        if m:
            return m.group(1)
        # javascript:fn_detail('XXXXX') 패턴
        m = re.search(r"fn_\w+\('(\d+)'", href)
        if m:
            return m.group(1)
    return None
