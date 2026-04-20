"""법령·행정규칙 연혁 웹 스크래퍼 — 법제처 API 미지원 시 fallback

2단계 전략:
  1단계 — requests + BeautifulSoup (정적 HTML 파싱, 빠름)
           - UTF-8 / EUC-KR 인코딩 자동 판별
  2단계 — Selenium headless Chrome (JS 렌더링 페이지 대응)
           - time.sleep 고정 대기 → WebDriverWait + EC 동적 대기
           - Selenium 4.6+ : Selenium Manager가 chromedriver 자동 다운로드

공개 함수:
  scrape_admrul_history(mst)  — 행정규칙 연혁 (admRulBylInfoR 등)
  scrape_law_history(mst)     — 법령 연혁 (lsBylInfoR / lsInfoP)

행정규칙 대상 URL:
  https://www.law.go.kr/admRulBylInfoR.do?admRulSeq={mst}
  https://www.law.go.kr/admRulLstv2R.do?admRulSeq={mst}
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
    "https://www.law.go.kr/admRulBylInfoR.do?admRulSeq={mst}",  # 행정규칙 연혁정보
    "https://www.law.go.kr/admRulLstv2R.do?admRulSeq={mst}",    # 행정규칙 목록 v2
    "https://www.law.go.kr/admRulInfoR.do?admRulSeq={mst}",     # 행정규칙 상세정보
]

_LAW_HISTORY_URL_PATTERNS = [
    "https://www.law.go.kr/lsBylInfoR.do?lsiSeq={mst}",   # 법령 연혁 목록
    "https://www.law.go.kr/lsInfoP.do?lsiSeq={mst}",      # 법령 상세 (연혁 탭 포함)
]

# Selenium: 테이블 요소 출현까지 최대 대기 시간 (초)
_SELENIUM_TABLE_TIMEOUT = 10
# 테이블 발견 후 내용 완전 로드 여유 시간 (초) — DOM 업데이트 대비 최소 대기
_SELENIUM_RENDER_MARGIN = 0.5


# ── 라이브러리 동적 임포트 ────────────────────────────────────────────────────

def _try_bs4():
    try:
        from bs4 import BeautifulSoup
        return BeautifulSoup
    except ImportError:
        return None


def _try_selenium():
    """Selenium + By 임포트 시도. 미설치 시 None 반환."""
    try:
        from selenium import webdriver
        from selenium.webdriver.chrome.options import Options
        from selenium.webdriver.common.by import By
        from selenium.webdriver.support.ui import WebDriverWait
        from selenium.webdriver.support import expected_conditions as EC
        from selenium.common.exceptions import TimeoutException
        return webdriver, Options, By, WebDriverWait, EC, TimeoutException
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

    # 한국어 형식: 2024년 1월 1일
    ko = re.match(r"(\d{4})\s*년\s*(\d{1,2})\s*월\s*(\d{1,2})\s*일", s)
    if ko:
        y, m, d = ko.group(1), ko.group(2).zfill(2), ko.group(3).zfill(2)
        return f"{y}{m}{d}"

    # 구분자 제거 후 8자리 숫자
    normalized = s.replace(".", "").replace("-", "").replace(" ", "")
    return normalized if (len(normalized) == 8 and normalized.isdigit()) else None


# ── 공개 진입점 ───────────────────────────────────────────────────────────────

def scrape_admrul_history(mst: str, timeout: int = 15) -> list[dict]:
    """행정규칙 연혁 목록 스크래핑.

    1단계(BS4) → 2단계(Selenium) 순으로 시도.
    두 단계 모두 실패하면 빈 목록 반환.

    Returns:
        [{"법령ID": str, "법령MST": str, "시행일자": str,
          "공포번호": str, "법령명한글": str}, ...]
        시행일자 기준 오름차순 정렬, 동일 시행일 중복 제거
    """
    # 1단계: 정적 HTML 파싱
    results = _scrape_static(mst, timeout)
    if results:
        return _dedup(results)

    # 2단계: Selenium headless (JS 렌더링 대응)
    results = _scrape_selenium(mst, timeout)
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


# ── 1단계: requests + BeautifulSoup ──────────────────────────────────────────

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

            # 요청 1회 후 인코딩만 바꿔 파싱 시도 (UTF-8 → EUC-KR 순)
            for encoding in ("utf-8", "euc-kr"):
                try:
                    resp.encoding = encoding
                    soup = BeautifulSoup(resp.text, "html.parser")
                    results = _parse_table(soup, mst)
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


# ── 2단계: Selenium headless Chrome ──────────────────────────────────────────

def _scrape_selenium(mst: str, timeout: int) -> list[dict]:
    """Selenium으로 JS 렌더링 후 파싱.

    테이블 요소 출현을 WebDriverWait로 동적 감지 — time.sleep 고정 대기 제거.
    테이블 미출현 시 TimeoutException을 잡아 다음 URL로 이동.
    """
    selenium_pkg = _try_selenium()
    if selenium_pkg is None:
        logger.debug("[스크래퍼 Selenium] selenium 미설치 — 스킵")
        return []

    webdriver, Options, By, WebDriverWait, EC, TimeoutException = selenium_pkg

    BeautifulSoup = _try_bs4()
    if BeautifulSoup is None:
        logger.debug("[스크래퍼 Selenium] beautifulsoup4 미설치 — 스킵")
        return []

    options = Options()
    options.add_argument("--headless=new")
    options.add_argument("--no-sandbox")
    options.add_argument("--disable-dev-shm-usage")
    options.add_argument("--disable-gpu")
    options.add_argument("--window-size=1280,900")
    options.add_argument(f"user-agent={_HEADERS['User-Agent']}")
    options.add_experimental_option("excludeSwitches", ["enable-logging"])

    driver = None
    try:
        driver = webdriver.Chrome(options=options)
        driver.set_page_load_timeout(timeout)

        for url_template in _HISTORY_URL_PATTERNS:
            url = url_template.format(mst=mst)
            try:
                driver.get(url)

                # 테이블 요소가 DOM에 나타날 때까지 동적 대기
                # TimeoutException 발생 시 해당 URL에 테이블 없음 → 다음으로
                try:
                    WebDriverWait(driver, _SELENIUM_TABLE_TIMEOUT).until(
                        EC.presence_of_element_located((By.CSS_SELECTOR, "table tr"))
                    )
                    # 테이블 발견 후 내용 완전 로드 여유
                    time.sleep(_SELENIUM_RENDER_MARGIN)
                except TimeoutException:
                    logger.debug(
                        "[스크래퍼 Selenium] 테이블 미출현 (timeout=%ds): %s",
                        _SELENIUM_TABLE_TIMEOUT, url,
                    )
                    continue

                soup = BeautifulSoup(driver.page_source, "html.parser")
                results = _parse_table(soup, mst)
                if results:
                    logger.info(
                        "[스크래퍼 Selenium] %d건 파싱 성공: %s", len(results), url
                    )
                    return results

            except Exception as exc:
                logger.debug("[스크래퍼 Selenium] 페이지 오류 (%s): %s", url, exc)

    except Exception as exc:
        logger.debug("[스크래퍼 Selenium] 드라이버 초기화 실패: %s", exc)
    finally:
        if driver:
            try:
                driver.quit()
            except Exception:
                pass

    return []


# ── HTML 테이블 파싱 (공통) ───────────────────────────────────────────────────

def _parse_table(soup, mst: str) -> list[dict]:
    """BeautifulSoup soup에서 연혁 테이블 파싱.

    헤더 키워드를 폭넓게 인식하여 페이지 구조 변형에 대응.
    """
    results = []

    for table in soup.find_all("table"):
        rows = table.find_all("tr")
        if len(rows) < 2:
            continue

        headers = [th.get_text(strip=True) for th in rows[0].find_all(["th", "td"])]

        # 시행일 컬럼 — 다양한 표기 대응
        col_enforce = _find_col(headers, [
            "시행일", "시행일자", "효력발생일", "발효일",
        ])
        # 공포번호 컬럼
        col_announce = _find_col(headers, [
            "공포번호", "제정·개정번호", "제정개정번호", "번호", "고시번호", "훈령번호",
        ])
        # 법령명 컬럼
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
        # admRulSeq, lsiSeq, ID 파라미터
        m = re.search(r'(?:admRulSeq|lsiSeq|ID)=(\d+)', href)
        if m:
            return m.group(1)
        # fn_xxx('12345') 형태 JS 호출
        m = re.search(r"fn_\w+\('(\d+)'", href)
        if m:
            return m.group(1)
        # javascript:fn_xxx('12345') 형태
        m = re.search(r"javascript:\w+\('(\d+)'", href)
        if m:
            return m.group(1)
    return None


# ── 법령 연혁 스크래핑 ────────────────────────────────────────────────────────

def scrape_law_history(mst: str, timeout: int = 15) -> list[dict]:
    """법령 연혁 목록 스크래핑 (API lawHistory.do 미지원 시 fallback).

    1단계(BS4) → 2단계(Selenium) 순으로 시도.

    Returns:
        [{"법령ID": str, "법령MST": str, "시행일자": str,
          "공포번호": str, "법령명한글": str}, ...]
        시행일자 기준 오름차순 정렬, 동일 시행일 중복 제거
    """
    results = _scrape_law_static(mst, timeout)
    if results:
        return _dedup(results)

    results = _scrape_law_selenium(mst, timeout)
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


def _scrape_law_selenium(mst: str, timeout: int) -> list[dict]:
    """Selenium headless Chrome으로 법령 연혁 JS 렌더링 후 파싱."""
    selenium_pkg = _try_selenium()
    if selenium_pkg is None:
        logger.debug("[스크래퍼 법령 Selenium] selenium 미설치 — 스킵")
        return []

    BeautifulSoup = _try_bs4()
    if BeautifulSoup is None:
        logger.debug("[스크래퍼 법령 Selenium] beautifulsoup4 미설치 — 스킵")
        return []

    webdriver, Options, By, WebDriverWait, EC, TimeoutException = selenium_pkg

    options = Options()
    options.add_argument("--headless=new")
    options.add_argument("--no-sandbox")
    options.add_argument("--disable-dev-shm-usage")
    options.add_argument("--disable-gpu")
    options.add_argument("--window-size=1280,900")
    options.add_argument(f"user-agent={_HEADERS['User-Agent']}")
    options.add_experimental_option("excludeSwitches", ["enable-logging"])

    driver = None
    try:
        driver = webdriver.Chrome(options=options)
        driver.set_page_load_timeout(timeout)

        for url_template in _LAW_HISTORY_URL_PATTERNS:
            url = url_template.format(mst=mst)
            try:
                driver.get(url)
                try:
                    WebDriverWait(driver, _SELENIUM_TABLE_TIMEOUT).until(
                        EC.presence_of_element_located((By.CSS_SELECTOR, "table tr"))
                    )
                    time.sleep(_SELENIUM_RENDER_MARGIN)
                except TimeoutException:
                    logger.debug(
                        "[스크래퍼 법령 Selenium] 테이블 미출현 (timeout=%ds): %s",
                        _SELENIUM_TABLE_TIMEOUT, url,
                    )
                    continue

                soup = BeautifulSoup(driver.page_source, "html.parser")
                results = _parse_law_table(soup, mst)
                if results:
                    logger.info(
                        "[스크래퍼 법령 Selenium] %d건 파싱 성공: %s", len(results), url
                    )
                    return results

            except Exception as exc:
                logger.debug("[스크래퍼 법령 Selenium] 페이지 오류 (%s): %s", url, exc)

    except Exception as exc:
        logger.debug("[스크래퍼 법령 Selenium] 드라이버 초기화 실패: %s", exc)
    finally:
        if driver:
            try:
                driver.quit()
            except Exception:
                pass

    return []


def _parse_law_table(soup, mst: str) -> list[dict]:
    """BeautifulSoup soup에서 법령 연혁 테이블 파싱.

    반환 키는 _raw_to_version(target="law")가 기대하는 필드명과 일치한다.
    """
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
