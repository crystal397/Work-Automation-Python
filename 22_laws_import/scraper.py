"""법령·행정규칙 연혁 웹 스크래퍼 — 법제처 API 미지원 시 fallback

전략:
  requests + BeautifulSoup (정적 HTML 파싱)
  - UTF-8 / EUC-KR 인코딩 자동 판별

공개 함수:
  scrape_admrul_history(mst)  — 행정규칙 연혁
  scrape_law_history(mst)     — 법령 연혁

행정규칙 대상 URL (admRulSeq=행정규칙일련번호 사용):
  https://www.law.go.kr/admRulHstListR.do?admRulSeq={mst}  ← 주 엔드포인트
  https://www.law.go.kr/admRulBylInfoR.do?admRulSeq={mst}
  https://www.law.go.kr/admRulInfoP.do?admRulSeq={mst}
  https://www.law.go.kr/admRulInfoR.do?admRulSeq={mst}

법령 대상 URL:
  https://www.law.go.kr/lsBylInfoR.do?lsiSeq={mst}
  https://www.law.go.kr/lsInfoP.do?lsiSeq={mst}
"""
import logging
import re
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

_HISTORY_URL_PATTERNS = [
    "https://www.law.go.kr/admRulBylInfoR.do?admRulSeq={mst}",
    "https://www.law.go.kr/admRulInfoP.do?admRulSeq={mst}",
    "https://www.law.go.kr/admRulInfoR.do?admRulSeq={mst}",
]

_LAW_HISTORY_URL_PATTERNS = [
    "https://www.law.go.kr/lsBylInfoR.do?lsiSeq={mst}",
    "https://www.law.go.kr/lsInfoP.do?lsiSeq={mst}",
]


# ── 라이브러리 동적 임포트 ────────────────────────────────────────────────────

def _try_bs4():
    try:
        from bs4 import BeautifulSoup
        return BeautifulSoup
    except ImportError:
        return None


# ── 날짜 파싱 ─────────────────────────────────────────────────────────────────

def _parse_date_str(s: str) -> Optional[str]:
    """다양한 날짜 표기 → 'YYYYMMDD' 정규화.

    지원 형식:
      '2024.01.01', '2024-01-01', '20240101'  → '20240101'
      '2024년 1월 1일', '2024년01월01일'       → '20240101'
    """
    s = s.strip()

    ko = re.match(r"(\d{4})\s*년\s*(\d{1,2})\s*월\s*(\d{1,2})\s*일", s)
    if ko:
        y, m, d = ko.group(1), ko.group(2).zfill(2), ko.group(3).zfill(2)
        return f"{y}{m}{d}"

    normalized = s.replace(".", "").replace("-", "").replace(" ", "")
    return normalized if (len(normalized) == 8 and normalized.isdigit()) else None


# ── 공개 진입점 ───────────────────────────────────────────────────────────────

def scrape_admrul_history(mst: str, timeout: int = 15) -> list[dict]:
    """행정규칙 연혁 목록 스크래핑.

    1단계: admRulHstListR.do?admRulSeq={mst} — 법제처 내부 연혁 목록 엔드포인트 (정적 HTML)
    2단계: 기존 URL 패턴 정적 파싱 (fallback)

    Args:
        mst: 행정규칙일련번호 (버전별 고유 순번, admRulSeq= 파라미터)

    Returns:
        [{"법령ID": str, "법령MST": str, "시행일자": str,
          "공포번호": str, "법령명한글": str}, ...]
        시행일자 기준 오름차순 정렬, 동일 시행일 중복 제거
    """
    results = _scrape_hst_list(mst, timeout)
    if results:
        return _dedup(results)

    results = _scrape_static(mst, timeout)
    if results:
        return _dedup(results)

    logger.debug("[스크래퍼] 모든 방법 실패 (mst=%s)", mst)
    return []


def _dedup(items: list[dict]) -> list[dict]:
    """시행일자 기준 중복 제거 + 오름차순 정렬"""
    seen: set[str] = set()
    unique = []
    for item in sorted(items, key=lambda x: x.get("시행일자", "")):
        key = item.get("시행일자", "")
        if key and key not in seen:
            seen.add(key)
            unique.append(item)
    return unique


# ── 1단계: admRulHstListR.do 연혁 목록 엔드포인트 ────────────────────────────

def _scrape_hst_list(mst: str, timeout: int) -> list[dict]:
    """admRulHstListR.do?admRulSeq={mst} 로 연혁 목록 조회.

    법제처 내부 AJAX 엔드포인트이지만 정적 HTML 응답.
    mst = 행정규칙일련번호 (admRulSeq 파라미터).
    """
    BeautifulSoup = _try_bs4()
    if BeautifulSoup is None:
        return []

    url = f"https://www.law.go.kr/admRulHstListR.do?admRulSeq={mst}"
    try:
        session = requests.Session()
        session.headers.update({**_HEADERS, "Referer": "https://www.law.go.kr/admRulInfoP.do"})
        resp = session.get(url, timeout=timeout)
        resp.raise_for_status()
        resp.encoding = "utf-8"
    except requests.exceptions.RequestException as exc:
        logger.debug("[스크래퍼 HstList] 요청 실패: %s", exc)
        return []

    soup = BeautifulSoup(resp.text, "html.parser")
    results = []
    date_re = re.compile(r"(\d{4})\.\s*(\d{1,2})\.\s*(\d{1,2})")

    for li in soup.find_all("li"):
        a = li.find("a")
        if not a:
            continue
        onclick = a.get("onclick", "")
        m = re.search(r"admRulViewHst\([^,]+,\s*'(\d+)'\)", onclick)
        if not m:
            continue
        seq_id = m.group(1)

        text = li.get_text(" ", strip=True)
        dates = date_re.findall(text)
        if not dates:
            continue

        # 텍스트 패턴: "[시행 YYYY.M.D.] [기관명 제NNN호, YYYY.M.D., 일부개정]"
        # dates[0] = 시행일, dates[1] = 공포일 (없으면 시행일과 동일)
        enforce_y, enforce_m, enforce_d = dates[0]
        enforce_date = f"{enforce_y}{int(enforce_m):02d}{int(enforce_d):02d}"

        if len(dates) >= 2:
            ann_y, ann_m, ann_d = dates[1]
            announce_date = f"{ann_y}{int(ann_m):02d}{int(ann_d):02d}"
        else:
            announce_date = enforce_date

        num_m = re.search(r"제\s*([\d\-]+)\s*호", text)
        announce_num = num_m.group(1) if num_m else ""

        results.append({
            "법령ID":    seq_id,
            "법령MST":   mst,
            "시행일자":   enforce_date,
            "공포일자":   announce_date,
            "공포번호":   announce_num,
            "법령명한글": a.get_text(strip=True),
        })

    if results:
        logger.info("[스크래퍼 HstList] %d건 파싱 성공 (mst=%s)", len(results), mst)
    return results


# ── 2단계: requests + BeautifulSoup (기존 URL 패턴) ──────────────────────────

def _scrape_static(mst: str, timeout: int) -> list[dict]:
    """requests로 HTML을 받아 BeautifulSoup으로 파싱.

    UTF-8 → EUC-KR 순으로 인코딩 시도.
    """
    BeautifulSoup = _try_bs4()
    if BeautifulSoup is None:
        logger.debug("[스크래퍼] beautifulsoup4 미설치 — 정적 파싱 스킵")
        return []

    session = requests.Session()
    session.headers.update(_HEADERS)

    for url_template in _HISTORY_URL_PATTERNS:
        url = url_template.format(mst=mst)
        try:
            resp = session.get(url, timeout=timeout)
            resp.raise_for_status()
            time.sleep(0.3)

            for encoding in ("utf-8", "euc-kr"):
                try:
                    resp.encoding = encoding
                    soup = BeautifulSoup(resp.text, "html.parser")
                    results = _parse_table(soup, mst)
                    if not results:
                        results = _parse_list(soup, mst)
                    if results:
                        logger.info(
                            "[스크래퍼 BS4] %d건 파싱 성공: %s (인코딩: %s)",
                            len(results), url, encoding,
                        )
                        return results
                except Exception as exc:
                    logger.debug("[스크래퍼 BS4] 파싱 오류 (%s, %s): %s", url, encoding, exc)

        except requests.exceptions.RequestException as exc:
            logger.debug("[스크래퍼 BS4] 요청 실패 (%s): %s", url, exc)

    return []


# ── HTML 테이블 파싱 (공통) ───────────────────────────────────────────────────

def _parse_table(soup, mst: str) -> list[dict]:
    """BeautifulSoup soup에서 연혁 테이블 파싱."""
    results = []

    for table in soup.find_all("table"):
        rows = table.find_all("tr")
        if len(rows) < 2:
            continue

        headers = [th.get_text(strip=True) for th in rows[0].find_all(["th", "td"])]

        col_enforce = _find_col(headers, [
            "시행일", "시행일자", "효력발생일", "발효일",
        ])
        col_announce = _find_col(headers, [
            "공포번호", "제정·개정번호", "제정개정번호", "번호", "고시번호", "훈령번호",
        ])
        col_name = _find_col(headers, [
            "행정규칙명", "법령명", "규칙명", "훈령명", "고시명", "예규명", "규정명",
        ])

        if col_enforce is None:
            continue

        for row in rows[1:]:
            cells = row.find_all(["td", "th"])
            if not cells:
                continue

            max_col = max(
                c for c in [col_enforce, col_announce, col_name] if c is not None
            )
            if len(cells) <= max_col:
                continue

            enforce_raw = cells[col_enforce].get_text(strip=True)
            enforce_date = _parse_date_str(enforce_raw)
            if not enforce_date:
                continue

            results.append({
                "법령ID":    _extract_id_from_row(row) or "",
                "법령MST":   mst,
                "시행일자":   enforce_date,
                "공포번호":   cells[col_announce].get_text(strip=True) if col_announce is not None else "",
                "법령명한글": cells[col_name].get_text(strip=True) if col_name is not None else "",
            })

    return results


def _parse_list(soup, mst: str) -> list[dict]:
    """테이블 구조가 없는 페이지에서 <ul>/<li> 또는 링크 기반으로 연혁 항목 추출."""
    results = []
    date_re = re.compile(r"(\d{4})[.\-](\d{1,2})[.\-](\d{1,2})")

    for a in soup.find_all("a", href=True):
        href = a["href"]
        id_m = re.search(r'(?:admRulSeq|lsiSeq|ID)=(\d+)', href)
        if not id_m:
            id_m = re.search(r"(?:fn_\w+|javascript:\w+)\s*\(\s*'(\d+)'", href)
        if not id_m:
            continue

        law_id = id_m.group(1)

        parent = a.find_parent(["li", "tr", "div", "td"])
        text_src = (parent or a).get_text(" ", strip=True)

        dates = date_re.findall(text_src)
        if not dates:
            continue

        last = dates[-1]
        enforce_date = f"{last[0]}{int(last[1]):02d}{int(last[2]):02d}"

        num_m = re.search(r"제?\s*(\d[\d\-]+\d)\s*호", text_src)
        announce_num = num_m.group(1) if num_m else ""

        results.append({
            "법령ID":    law_id,
            "법령MST":   mst,
            "시행일자":   enforce_date,
            "공포번호":   announce_num,
            "법령명한글": a.get_text(strip=True),
        })

    return results


def _find_col(headers: list[str], keywords: list[str]) -> Optional[int]:
    """헤더 목록에서 키워드가 포함된 첫 번째 컬럼 인덱스 반환."""
    for kw in keywords:
        for i, h in enumerate(headers):
            if kw in h:
                return i
    return None


def _extract_id_from_row(row) -> Optional[str]:
    """테이블 행의 링크에서 법령ID 추출."""
    for a in row.find_all("a", href=True):
        href = a["href"]
        m = re.search(r'(?:admRulSeq|lsiSeq|ID)=(\d+)', href)
        if m:
            return m.group(1)
        m = re.search(r"fn_\w+\('(\d+)'", href)
        if m:
            return m.group(1)
        m = re.search(r"javascript:\w+\('(\d+)'", href)
        if m:
            return m.group(1)
    return None


# ── 법령 연혁 스크래핑 ────────────────────────────────────────────────────────

def scrape_law_history(mst: str, timeout: int = 15) -> list[dict]:
    """법령 연혁 목록 스크래핑 (API lawHistory.do 미지원 시 fallback).

    Returns:
        [{"법령ID": str, "법령MST": str, "시행일자": str,
          "공포번호": str, "법령명한글": str}, ...]
        시행일자 기준 오름차순 정렬, 동일 시행일 중복 제거
    """
    results = _scrape_law_static(mst, timeout)
    if results:
        return _dedup(results)

    logger.debug("[스크래퍼 법령] 모든 방법 실패 (mst=%s)", mst)
    return []


def _scrape_law_static(mst: str, timeout: int) -> list[dict]:
    """requests + BeautifulSoup으로 법령 연혁 HTML 파싱."""
    BeautifulSoup = _try_bs4()
    if BeautifulSoup is None:
        logger.debug("[스크래퍼 법령] beautifulsoup4 미설치 — 정적 파싱 스킵")
        return []

    session = requests.Session()
    session.headers.update(_HEADERS)

    for url_template in _LAW_HISTORY_URL_PATTERNS:
        url = url_template.format(mst=mst)
        try:
            resp = session.get(url, timeout=timeout)
            resp.raise_for_status()
            time.sleep(0.3)

            for encoding in ("utf-8", "euc-kr"):
                try:
                    resp.encoding = encoding
                    soup = BeautifulSoup(resp.text, "html.parser")
                    results = _parse_law_table(soup, mst)
                    if results:
                        logger.info(
                            "[스크래퍼 법령 BS4] %d건 파싱 성공: %s (인코딩: %s)",
                            len(results), url, encoding,
                        )
                        return results
                except Exception as exc:
                    logger.debug("[스크래퍼 법령 BS4] 파싱 오류 (%s, %s): %s", url, encoding, exc)

        except requests.exceptions.RequestException as exc:
            logger.debug("[스크래퍼 법령 BS4] 요청 실패 (%s): %s", url, exc)

    return []


def _parse_law_table(soup, mst: str) -> list[dict]:
    """BeautifulSoup soup에서 법령 연혁 테이블 파싱."""
    results = []

    for table in soup.find_all("table"):
        rows = table.find_all("tr")
        if len(rows) < 2:
            continue

        headers = [th.get_text(strip=True) for th in rows[0].find_all(["th", "td"])]

        col_enforce = _find_col(headers, ["시행일", "시행일자", "효력발생일", "발효일"])
        col_announce = _find_col(headers, [
            "공포번호", "제정·개정번호", "제정개정번호", "번호",
        ])
        col_name = _find_col(headers, ["법령명", "법령명한글", "명칭"])

        if col_enforce is None:
            continue

        for row in rows[1:]:
            cells = row.find_all(["td", "th"])
            if not cells:
                continue

            max_col = max(
                c for c in [col_enforce, col_announce, col_name] if c is not None
            )
            if len(cells) <= max_col:
                continue

            enforce_raw = cells[col_enforce].get_text(strip=True)
            enforce_date = _parse_date_str(enforce_raw)
            if not enforce_date:
                continue

            results.append({
                "법령ID":    _extract_id_from_row(row) or "",
                "법령MST":   mst,
                "시행일자":   enforce_date,
                "공포번호":   cells[col_announce].get_text(strip=True) if col_announce is not None else "",
                "법령명한글": cells[col_name].get_text(strip=True) if col_name is not None else "",
            })

    return results
