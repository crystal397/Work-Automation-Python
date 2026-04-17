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
                logger.warning("HTTP 오류 (%d/%d): %s — %s", attempt, config.API_RETRY, endpoint, exc)
            except requests.exceptions.RequestException as exc:
                logger.warning("요청 실패 (%d/%d): %s — %s", attempt, config.API_RETRY, endpoint, exc)

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
        result = data.get("LawSearch", {}) or {}
        laws = result.get("law", [])
        if isinstance(laws, dict):
            laws = [laws]
        return laws or []

    def get_law_history(self, mst: str, target: str = "law") -> list[dict]:
        """연혁 법령 전체 목록 조회 (공포번호·공포일·시행일 포함)

        Args:
            mst:    법령 MST 번호 (법령 계열 식별자)
            target: "law" 또는 "admrul"

        Returns:
            연혁 법령 dict 목록 (시행일 오름차순 정렬되지 않을 수 있음)
        """
        data = self._get("lawHistory.do", {"target": target, "MST": mst})
        history_root = data.get("LawHistory", {}) or {}
        history_section = history_root.get("연혁", {}) or {}
        items = history_section.get("연혁법령", [])
        if isinstance(items, dict):
            items = [items]
        return items or []

    def get_law_text(self, law_id: str, target: str = "law") -> dict:
        """법령 본문 조회 — 조문 + 부칙 포함

        Args:
            law_id: 법령ID (버전별 고유 번호, lawSearch/lawHistory에서 획득)
            target: "law" 또는 "admrul"

        Returns:
            '법령' 루트 dict. 기본정보 / 조문 / 부칙 하위 구조 포함.
        """
        data = self._get("lawService.do", {"target": target, "ID": law_id})
        return data.get("법령", {}) or {}

    def get_admrul_text(self, law_id: str) -> dict:
        """행정규칙 본문 조회 (target=admrul 단축 메서드)"""
        return self.get_law_text(law_id, target="admrul")
