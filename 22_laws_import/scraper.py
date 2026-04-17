"""행정규칙 연혁 웹 스크래퍼 — 법제처 API 미지원 시 fallback

2단계 전략:
  1단계 — requests + BeautifulSoup (정적 HTML 파싱, 빠름)
  2단계 — Selenium headless Chrome (JS 렌더링 페이지 대응)
           Selenium 4.6+ : Selenium Manager가 chromedriver 자동 다운로드
           Chrome 브라우저만 설치되어 있으면 별도 드라이버 불필요

대상 URL:
  https://www.law.go.kr/admRulBylInfoR.do?admRulSeq={mst}
  https://www.law.go.kr/admRulLstv2R.do?admRulSeq={mst}

주의:
  - law.go.kr 페이지 구조 변경 시 selector 수정 필요
  - Selenium 사용 시 첫 실행에 chromedriver 자동 다운로드 (인터넷 필요)
  - Chrome 미설치 시 Selenium 단계 조용히 스킵
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
    "https://www.law.go.kr/admRulLstv2R.do?admRulSeq={mst}",
]

# JS 렌더링 완료 대기 시간 (초)
_SELENIUM_WAIT = 3


# ── 라이브러리 동적 임포트 ────────────────────────────────────────────────────

def _try_bs4():
    try:
        from bs4 import BeautifulSoup
        return BeautifulSoup
    except ImportError:
        return None


def _try_selenium():
    """Selenium webdriver 임포트 시도. 미설치 시 None 반환."""
    try:
        from selenium import webdriver
        from selenium.webdriver.chrome.options import Options
        from selenium.webdriver.support.ui import WebDriverWait
        from selenium.webdriver.support import expected_conditions as EC
        return webdriver, Options, WebDriverWait, EC
    except ImportError:
        return None


# ── 날짜 파싱 ─────────────────────────────────────────────────────────────────

def _parse_date_str(s: str) -> Optional[str]:
    """'2024.01.01' 또는 '20240101' → '20240101'"""
    s = s.strip().replace(".", "").replace("-", "").replace(" ", "")
    return s if (len(s) == 8 and s.isdigit()) else None


# ── 공개 진입점 ───────────────────────────────────────────────────────────────

def scrape_admrul_history(mst: str, timeout: int = 15) -> list[dict]:
    """행정규칙 연혁 목록 스크래핑.

    1단계(BS4) → 2단계(Selenium) 순으로 시도.
    두 단계 모두 실패하면 빈 목록 반환.

    Returns:
        [{"법령ID": str, "시행일자": str, "공포번호": str, "법령명한글": str}, ...]
    """
    # 1단계: 정적 HTML 파싱
    results = _scrape_static(mst, timeout)
    if results:
        return results

    # 2단계: Selenium headless (JS 렌더링 대응)
    results = _scrape_selenium(mst, timeout)
    if results:
        return results

    logger.debug("[스크래퍼] 모든 방법 실패 (mst=%s)", mst)
    return []


# ── 1단계: requests + BeautifulSoup ──────────────────────────────────────────

def _scrape_static(mst: str, timeout: int) -> list[dict]:
    """requests로 HTML을 받아 BeautifulSoup으로 파싱."""
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
            resp.encoding = "utf-8"
            time.sleep(0.5)

            soup = BeautifulSoup(resp.text, "html.parser")
            results = _parse_table(soup, mst)
            if results:
                logger.info("[스크래퍼 BS4] %d건 파싱 성공: %s", len(results), url)
                return results

        except requests.exceptions.RequestException as exc:
            logger.debug("[스크래퍼 BS4] 요청 실패 (%s): %s", url, exc)
        except Exception as exc:
            logger.debug("[스크래퍼 BS4] 파싱 오류 (%s): %s", url, exc)

    return []


# ── 2단계: Selenium headless Chrome ──────────────────────────────────────────

def _scrape_selenium(mst: str, timeout: int) -> list[dict]:
    """Selenium으로 JS 렌더링 후 파싱.

    Selenium 4.6+ Selenium Manager가 chromedriver를 자동 관리.
    Chrome 미설치 또는 Selenium 미설치 시 조용히 스킵.
    """
    selenium_pkg = _try_selenium()
    if selenium_pkg is None:
        logger.debug("[스크래퍼 Selenium] selenium 미설치 — 스킵")
        return []

    webdriver, Options, WebDriverWait, EC = selenium_pkg
    BeautifulSoup = _try_bs4()
    if BeautifulSoup is None:
        logger.debug("[스크래퍼 Selenium] beautifulsoup4 미설치 — 스킵")
        return []

    options = Options()
    options.add_argument("--headless=new")       # Chrome 112+ headless 모드
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
                # JS 렌더링 완료 대기
                time.sleep(_SELENIUM_WAIT)

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
        # Chrome 미설치, 드라이버 오류 등
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
    """BeautifulSoup soup에서 연혁 테이블 파싱."""
    results = []

    for table in soup.find_all("table"):
        rows = table.find_all("tr")
        if len(rows) < 2:
            continue

        headers = [th.get_text(strip=True) for th in rows[0].find_all(["th", "td"])]
        col_enforce = _find_col(headers, ["시행일", "시행일자"])
        col_announce = _find_col(headers, ["공포번호", "제정·개정번호", "번호"])
        col_name = _find_col(headers, ["행정규칙명", "법령명", "규칙명"])

        if col_enforce is None:
            continue

        for row in rows[1:]:
            cells = row.find_all(["td", "th"])
            max_col = max(
                c for c in [col_enforce, col_announce, col_name] if c is not None
            )
            if len(cells) <= max_col:
                continue

            enforce_date = _parse_date_str(
                cells[col_enforce].get_text(strip=True)
            )
            if not enforce_date:
                continue

            results.append({
                "법령ID": _extract_id_from_row(row) or "",
                "법령MST": mst,
                "시행일자": enforce_date,
                "공포번호": cells[col_announce].get_text(strip=True) if col_announce is not None else "",
                "법령명한글": cells[col_name].get_text(strip=True) if col_name is not None else "",
            })

    return results


def _find_col(headers: list[str], keywords: list[str]) -> Optional[int]:
    for kw in keywords:
        for i, h in enumerate(headers):
            if kw in h:
                return i
    return None


def _extract_id_from_row(row) -> Optional[str]:
    for a in row.find_all("a", href=True):
        href = a["href"]
        m = re.search(r'(?:admRulSeq|lsiSeq|ID)=(\d+)', href)
        if m:
            return m.group(1)
        m = re.search(r"fn_\w+\('(\d+)'", href)
        if m:
            return m.group(1)
    return None
