"""법제처 국가법령정보 공동활용 API 클라이언트 (경로 A)

기준 URL: http://www.law.go.kr/DRF/
인증 방식: OC (법제처 가입 이메일 ID)

엔드포인트:
  - lawSearch.do  : 법령명 검색 → 법령ID·MST 목록
  - lawHistory.do : MST 기준 연혁 버전 목록 (공포번호·시행일 포함)
  - lawService.do : 특정 법령 본문 (조문·부칙 포함)
"""
import logging
import time
from typing import Optional

import requests
import xmltodict

import config

logger = logging.getLogger(__name__)


class LawAPIClient:
    """법제처 API 호출 — 검색·연혁·본문·행정규칙"""

    BASE = config.LAW_API_BASE

    def __init__(self, oc: str = ""):
        self.oc = oc or config.LAW_API_OC
        if not self.oc:
            raise ValueError(
                "법제처 기관코드(OC)가 없습니다.\n"
                ".env 파일에 LAW_API_OC=이메일 형식으로 입력하거나 "
                "GUI에서 직접 입력하세요."
            )
        self.session = requests.Session()

    # ── 내부 헬퍼 ─────────────────────────────────────────────────────────────

    def _get(self, endpoint: str, params: dict) -> dict:
        """공통 GET 요청 — XML 응답을 dict으로 변환, 재시도 포함"""
        url = f"{self.BASE}/{endpoint}"
        full_params = {"OC": self.oc, "type": "XML", **params}

        for attempt in range(1, config.API_RETRY + 1):
            try:
                logger.debug("API 요청: %s params=%s", endpoint, params)
                resp = self.session.get(
                    url, params=full_params, timeout=config.API_TIMEOUT
                )
                resp.raise_for_status()
                resp.encoding = "utf-8"
                return xmltodict.parse(resp.text)
            except requests.exceptions.Timeout:
                logger.warning("타임아웃 (%d/%d): %s", attempt, config.API_RETRY, endpoint)
            except requests.exceptions.HTTPError as exc:
                # 404는 엔드포인트 자체가 없는 것 — 재시도해도 소용없으므로 즉시 중단
                if exc.response is not None and exc.response.status_code == 404:
                    raise
                logger.warning("HTTP 오류 (%d/%d): %s — %s", attempt, config.API_RETRY, endpoint, exc)
            except requests.exceptions.RequestException as exc:
                logger.warning("요청 실패 (%d/%d): %s — %s", attempt, config.API_RETRY, endpoint, exc)
            except Exception as exc:
                # API 서버가 200이지만 비XML 응답(HTML 오류 페이지 등) 반환 시 XML 파싱 실패
                logger.warning("응답 파싱 실패 (%d/%d): %s — %s", attempt, config.API_RETRY, endpoint, exc)

            if attempt < config.API_RETRY:
                time.sleep(config.API_RETRY_DELAY)

        raise ConnectionError(
            f"API 호출 최대 재시도 초과: {endpoint}\n"
            "네트워크 연결과 기관코드(OC)를 확인하세요."
        )

    # ── 공개 메서드 ───────────────────────────────────────────────────────────

    def search_law(
        self, query: str, target: str = "law", display: int = 10
    ) -> list[dict]:
        """법령명으로 검색 → 법령 목록 반환

        Args:
            query:   검색어 (법령명 전체 또는 일부)
            target:  "law"(법령) 또는 "admrul"(행정규칙)
            display: 검색 결과 수 (최대 100)

        Returns:
            법령 dict 목록. 각 항목에 법령ID, 법령MST, 법령명한글,
            공포번호, 공포일자, 시행일자 포함.
        """
        data = self._get(
            "lawSearch.do",
            {"target": target, "query": query, "display": display},
        )
        # 행정규칙은 루트 요소와 항목 키가 법령과 다름
        if target == "admrul":
            result = data.get("AdmRulSearch", {}) or {}
            items = result.get("admrul", [])
        else:
            result = data.get("LawSearch", {}) or {}
            items = result.get("law", [])
        if isinstance(items, dict):
            items = [items]
        return items or []

    def get_law_history(self, mst: str = "", query: str = "", target: str = "law") -> list[dict]:
        """연혁 법령 전체 목록 조회 (공포번호·공포일·시행일 포함)

        Args:
            mst:    법령 MST 번호 — admrul 전용 (lawHistory.do 파라미터)
            query:  법령명 검색어  — law 전용 (lsHistory 파라미터)
            target: "law" 또는 "admrul"

        Returns:
            연혁 법령 dict 목록 (시행일 오름차순 정렬되지 않을 수 있음)

        Note:
            law   타입: lawSearch.do?target=lsHistory 사용 (페이지네이션 포함)
            admrul타입: lawHistory.do?MST=mst 사용 (기존 방식 유지)
        """
        if target == "admrul":
            data = self._get("lawHistory.do", {"target": target, "MST": mst})
            history_root = data.get("LawHistory", {}) or {}
            history_section = history_root.get("연혁", {}) or {}
            items = history_section.get("연혁법령", [])
            if isinstance(items, dict):
                items = [items]
            return items or []

        # law 타입: lsHistory 엔드포인트 (HTML 파싱 + 페이지네이션)
        # ※ lsHistory는 type=XML 미지원 — HTML 응답을 직접 파싱
        import re
        from bs4 import BeautifulSoup

        def _norm_date(s: str) -> str:
            """'2026.3.10' → '20260310'"""
            parts = s.strip().split(".")
            if len(parts) == 3:
                return f"{parts[0]}{int(parts[1]):02d}{int(parts[2]):02d}"
            return s.replace(".", "")

        all_items: list[dict] = []
        page = 1
        while True:
            try:
                resp = self.session.get(
                    f"{self.BASE}/lawSearch.do",
                    params={
                        "OC": self.oc, "type": "HTML", "target": "lsHistory",
                        "query": query, "display": 100, "page": page,
                    },
                    timeout=config.API_TIMEOUT,
                )
                resp.raise_for_status()
                resp.encoding = "utf-8"
            except Exception as exc:
                logger.warning("lsHistory HTML 요청 실패 (page=%d): %s", page, exc)
                break

            soup = BeautifulSoup(resp.text, "html.parser")
            tables = soup.find_all("table")
            if not tables:
                break

            # 링크에서 법령일련번호(MST) 추출
            links = [a for a in soup.find_all("a", href=True) if "lsHistory" in a["href"]]
            mst_list = []
            for a in links:
                m = re.search(r"MST=(\d+)", a["href"])
                mst_list.append(m.group(1) if m else "")

            rows = tables[0].find_all("tr")[1:]  # 헤더 제외
            for i, row in enumerate(rows):
                cells = [td.get_text(strip=True) for td in row.find_all("td")]
                if len(cells) < 8:
                    continue
                lsi = mst_list[i] if i < len(mst_list) else ""
                announce_num = re.sub(r"[^0-9\-]", "", cells[5])  # "제 21418호" → "21418"
                item = {
                    "법령명한글": cells[1],
                    "공포번호": announce_num,
                    "공포일자": _norm_date(cells[6]),
                    "시행일자": _norm_date(cells[7]),
                    "법령일련번호": lsi,
                    "법령ID": lsi,   # 본문 조회(lawService.do?ID=) 에 동일 값 사용
                    "현행연혁코드": "현행" if "현행" in cells[-1] else "연혁",
                }
                all_items.append(item)

            if len(rows) < 100:
                break
            page += 1

        return all_items

    def get_law_text(self, law_id: str, target: str = "law") -> dict:
        """법령 본문 조회 — 조문 + 부칙 포함

        Args:
            law_id: 법령ID (버전별 고유 번호, lawSearch/lawHistory에서 획득)
            target: "law" 또는 "admrul"

        Returns:
            law   → '법령' 루트 dict (기본정보 / 조문단위 / 부칙단위 구조)
            admrul→ 'AdmRulService' 루트 dict (행정규칙기본정보 / 조문내용 배열)
        """
        data = self._get("lawService.do", {"target": target, "ID": law_id})
        if target == "admrul":
            return data.get("AdmRulService", {}) or {}
        return data.get("법령", {}) or {}

    def get_admrul_text(self, law_id: str) -> dict:
        """행정규칙 본문 조회 (target=admrul 단축 메서드)"""
        return self.get_law_text(law_id, target="admrul")
