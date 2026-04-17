"""핵심 매칭 엔진 — 입찰공고일 기준 6단계 시행일 판단 로직

6단계:
  1. 연혁 조회     — 해당 법령 전체 버전(공포번호·시행일) 수집
  2. 1차 후보 선정 — 시행일 ≤ 입찰공고일 중 가장 최근 버전
  3. 부칙 경과규정 — 1차 후보 부칙 파싱, 경과규정 패턴 탐지
  4. 2차 후보 선정 — 경과규정 존재 시 직전 버전을 병렬 후보로 제시
  5. 상위·하위법 정합성 — (결과 표시 시 law/시행령/시행규칙 시행일 비교 경고)
  6. 사용자 검토   — 자동 확정 불가 시 두 후보 모두 출력 + 플래그
"""
import logging
import re
from dataclasses import dataclass, field
from datetime import date
from typing import Callable, Optional

from api_client import LawAPIClient
from cache import LawCache
import config

logger = logging.getLogger(__name__)

# ── 부칙 경과규정 탐지 정규표현식 ────────────────────────────────────────────
_TRANSITIONAL_PATTERNS: list[re.Pattern] = [
    re.compile(r"시행\s*전에?\s*(이미\s*)?(입찰\s*공고|공고)"),
    re.compile(r"종전의?\s*규정에\s*따른다"),
    re.compile(r"최초로\s*입찰\s*공고하는?\s*분부터"),
    re.compile(r"시행\s*당시\s*종전"),
    re.compile(r"이\s*(법|영|규칙)\s*시행\s*전.*?(입찰|공고|계약)"),
    re.compile(r"개정규정은\s*이\s*(법|영|규칙)\s*시행\s*후\s*최초"),
    re.compile(r"공고된\s*계약.*?종전"),
]


# ── 데이터 클래스 ─────────────────────────────────────────────────────────────

@dataclass
class LawVersion:
    """단일 법령 버전 정보"""
    law_id: str           # 법령ID (버전 고유값)
    mst: str              # 법령MST (법령 계열 식별자)
    name: str             # 법령명한글
    target: str           # "law" | "admrul"
    announce_num: str     # 공포번호
    announce_date: date   # 공포일자
    enforce_date: date    # 시행일자
    text: dict = field(default_factory=dict, repr=False)  # 법령 본문 (lazy load)

    @property
    def source_url(self) -> str:
        return f"https://www.law.go.kr/lsInfoP.do?lsiSeq={self.law_id}"


@dataclass
class MatchResult:
    """단일 법령 매칭 최종 결과"""
    display_name: str                     # 표시명 (config.TARGET_LAWS 기준)
    selected: Optional[LawVersion]        # 1차 선택 버전
    prev_version: Optional[LawVersion]    # 직전 버전 (부칙 해당 시 병렬 제시)
    transitional_flag: bool               # 부칙 경과규정 탐지 여부
    transitional_text: str                # 탐지된 경과규정 발췌문
    relevant_articles: list[dict] = field(default_factory=list)
    needs_user_review: bool = False       # 사용자 확인 필요
    warning: str = ""                     # 경고 메시지


# ── 유틸 ──────────────────────────────────────────────────────────────────────

def _parse_date(s: str) -> date:
    s = str(s).strip().replace("-", "")
    if len(s) == 8 and s.isdigit():
        return date(int(s[:4]), int(s[4:6]), int(s[6:8]))
    raise ValueError(f"날짜 파싱 실패: {s!r}")


def _raw_to_version(raw: dict, mst: str, target: str) -> Optional[LawVersion]:
    """API 응답 raw dict → LawVersion. 실패 시 None."""
    try:
        enforce_str = str(raw.get("시행일자") or "").strip()
        if not enforce_str:
            return None
        announce_str = str(raw.get("공포일자") or "").strip()
        return LawVersion(
            law_id=str(raw.get("법령ID") or raw.get("ID") or ""),
            mst=str(raw.get("법령MST") or mst),
            name=str(raw.get("법령명한글") or raw.get("법령명") or ""),
            target=target,
            announce_num=str(raw.get("공포번호") or ""),
            announce_date=_parse_date(announce_str) if announce_str else date(1900, 1, 1),
            enforce_date=_parse_date(enforce_str),
        )
    except Exception as exc:
        logger.debug("버전 파싱 스킵: %s — %s", raw, exc)
        return None


# ── 메인 엔진 ─────────────────────────────────────────────────────────────────

class LawMatcher:
    """입찰공고일 기준 법령 버전 자동 매칭기"""

    def __init__(self, oc: str = ""):
        self.client = LawAPIClient(oc)
        self.cache = LawCache()

    # ── 내부 — 캐시 래핑 ──────────────────────────────────────────────────────

    def _search_mst(self, query: str, target: str) -> Optional[str]:
        """법령명으로 MST 번호 조회 (캐시 우선)"""
        cache_key = f"search:{target}:{query}"
        cached = self.cache.get(cache_key)
        if cached:
            return cached

        laws = self.client.search_law(query, target=target, display=10)
        for law in laws:
            name = str(law.get("법령명한글") or law.get("법령명") or "")
            mst = str(law.get("법령MST") or law.get("MST") or "")
            # 정확히 일치하는 것 우선, 없으면 부분 일치 허용
            if name == query and mst:
                self.cache.set(cache_key, mst)
                logger.info("MST 확인(정확): %s → %s", query, mst)
                return mst
        for law in laws:
            name = str(law.get("법령명한글") or law.get("법령명") or "")
            mst = str(law.get("법령MST") or law.get("MST") or "")
            if query in name and mst:
                self.cache.set(cache_key, mst)
                logger.info("MST 확인(부분): %s → %s (name=%s)", query, mst, name)
                return mst

        logger.warning("MST 조회 실패: %s", query)
        return None

    def _get_history(self, mst: str, target: str) -> list[LawVersion]:
        """연혁 법령 목록 조회 (캐시 우선)"""
        cache_key = f"history:{target}:{mst}"
        cached: Optional[list] = self.cache.get(cache_key)

        raw_list: list[dict]
        if cached is not None:
            raw_list = cached
        else:
            raw_list = self.client.get_law_history(mst, target=target)
            if raw_list:
                self.cache.set(cache_key, raw_list)

        versions = [v for raw in raw_list if (v := _raw_to_version(raw, mst, target))]
        versions.sort(key=lambda v: v.enforce_date)
        return versions

    def _get_text(self, version: LawVersion) -> dict:
        """법령 본문 조회 (캐시 우선). 실패 시 빈 dict."""
        if version.text:
            return version.text

        cache_key = f"text:{version.target}:{version.law_id}"
        cached = self.cache.get(cache_key)
        if cached:
            version.text = cached
            return cached

        try:
            text = self.client.get_law_text(version.law_id, target=version.target)
            if text:
                version.text = text
                self.cache.set(cache_key, text)
            return text
        except Exception as exc:
            logger.error("본문 조회 실패 (ID=%s): %s", version.law_id, exc)
            return {}

    # ── 내부 — 분석 ───────────────────────────────────────────────────────────

    def _detect_transitional(self, text: dict) -> tuple[bool, str]:
        """부칙에서 경과규정 패턴 탐지 → (발견 여부, 발췌문)"""
        sections = text.get("부칙") or {}
        units = sections.get("부칙단위", [])
        if isinstance(units, dict):
            units = [units]

        full_text = ""
        for unit in units:
            content = unit.get("부칙내용") or ""
            if isinstance(content, list):
                content = " ".join(str(c) for c in content)
            full_text += str(content) + " "

        for pattern in _TRANSITIONAL_PATTERNS:
            m = pattern.search(full_text)
            if m:
                start = max(0, m.start() - 20)
                end = min(len(full_text), m.end() + 60)
                excerpt = full_text[start:end].strip()
                logger.warning("경과규정 탐지 — 패턴: %s | 발췌: %s", pattern.pattern, excerpt[:80])
                return True, excerpt

        return False, ""

    def _filter_articles(self, text: dict) -> list[dict]:
        """공기연장 키워드가 포함된 조문 필터링"""
        articles = []
        article_units = (text.get("조문") or {}).get("조문단위", [])
        if isinstance(article_units, dict):
            article_units = [article_units]

        for unit in article_units:
            title = str(unit.get("조제목") or "")
            content = str(unit.get("조문내용") or "")

            paragraphs = unit.get("항") or []
            if isinstance(paragraphs, dict):
                paragraphs = [paragraphs]
            para_text = " ".join(str(p.get("항내용") or "") for p in paragraphs)

            combined = title + content + para_text
            if any(kw in combined for kw in config.EXTENSION_KEYWORDS):
                articles.append({
                    "조번호": str(unit.get("조번호") or ""),
                    "조제목": title,
                    "조문내용": content,
                    "항": paragraphs,
                })

        return articles

    # ── 6단계 매칭 로직 ────────────────────────────────────────────────────────

    def _match_admrul(
        self,
        display_name: str,
        query: str,
        bid_date: date,
    ) -> MatchResult:
        """행정규칙 전용 매칭 — 법제처 API는 admrul 연혁 조회를 지원하지 않으므로
        현행(최신) 버전만 조회하고 사용자에게 수동 확인을 요청한다."""
        logger.warning(
            "[행정규칙] '%s' — 연혁 조회 불가, 현행 버전만 조회", display_name
        )

        # search_law 한 번으로 MST + 버전 정보를 동시에 확보 (이중 호출 방지)
        laws = self.client.search_law(query, target="admrul", display=10)
        if not laws:
            return MatchResult(
                display_name=display_name,
                selected=None, prev_version=None,
                transitional_flag=False, transitional_text="",
                warning=f"행정규칙 조회 실패: '{query}'",
                needs_user_review=True,
            )

        # 정확 일치 우선, 없으면 부분 일치
        raw = next(
            (l for l in laws if str(l.get("법령명한글") or l.get("법령명") or "") == query),
            None,
        ) or laws[0]

        mst = str(raw.get("법령MST") or raw.get("MST") or "")
        version = _raw_to_version(raw, mst, "admrul")
        if not version:
            return MatchResult(
                display_name=display_name,
                selected=None, prev_version=None,
                transitional_flag=False, transitional_text="",
                warning=f"행정규칙 버전 파싱 실패: '{query}'",
                needs_user_review=True,
            )

        text = self._get_text(version)
        relevant_articles = self._filter_articles(text)

        return MatchResult(
            display_name=display_name,
            selected=version,
            prev_version=None,
            transitional_flag=False,
            transitional_text="",
            relevant_articles=relevant_articles,
            needs_user_review=True,
            warning=(
                "⚠ 행정규칙은 연혁 조회 불가 — 현행 버전 기준 표시. "
                f"입찰공고일({bid_date}) 기준 실제 시행 버전은 수동 확인 필요"
            ),
        )

    def match_one(
        self,
        display_name: str,
        query: str,
        target: str,
        bid_date: date,
    ) -> MatchResult:
        """단일 법령에 대한 6단계 시행일 매칭 실행"""
        logger.info("─── [매칭 시작] %s | 입찰공고일: %s ───", display_name, bid_date)

        # 행정규칙은 연혁 API 미지원 → 전용 메서드로 분기
        if target == "admrul":
            return self._match_admrul(display_name, query, bid_date)

        # 1단계 — 연혁 조회
        mst = self._search_mst(query, target)
        if not mst:
            return MatchResult(
                display_name=display_name,
                selected=None, prev_version=None,
                transitional_flag=False, transitional_text="",
                warning=f"법령 MST 조회 실패: '{query}'",
                needs_user_review=True,
            )

        all_versions = self._get_history(mst, target)
        if not all_versions:
            return MatchResult(
                display_name=display_name,
                selected=None, prev_version=None,
                transitional_flag=False, transitional_text="",
                warning=f"연혁 법령 없음: '{query}'",
                needs_user_review=True,
            )

        logger.info("연혁 버전 %d개 확인", len(all_versions))

        # 2단계 — 1차 후보 선정
        candidates = [v for v in all_versions if v.enforce_date <= bid_date]
        if not candidates:
            # 입찰공고일 이전 시행 버전이 없으면 가장 오래된 버전 사용
            primary = all_versions[0]
            logger.warning("시행일 이전 버전 없음 → 최고 오래된 버전 사용: %s", primary.enforce_date)
            return MatchResult(
                display_name=display_name,
                selected=primary, prev_version=None,
                transitional_flag=False, transitional_text="",
                warning="입찰공고일 이전 시행 버전 없음 — 가장 오래된 버전 표시",
                needs_user_review=True,
            )

        primary = candidates[-1]
        prev_version = candidates[-2] if len(candidates) >= 2 else None
        logger.info(
            "1차 후보: 공포번호=%s | 시행일=%s", primary.announce_num, primary.enforce_date
        )

        # 3단계 — 부칙 경과규정 확인
        text = self._get_text(primary)
        transitional, excerpt = self._detect_transitional(text)

        # 4단계 — 경과규정 존재 시 경고 + 직전 버전 병렬 제시
        if transitional:
            logger.warning(
                "부칙 경과규정 탐지 → 직전 버전(%s) 병렬 제시, 사용자 확인 필요",
                prev_version.announce_num if prev_version else "없음",
            )

        # 5단계 — 조문 필터링
        relevant_articles = self._filter_articles(text)
        logger.info("공기연장 관련 조문: %d개", len(relevant_articles))

        # 6단계 — 사용자 검토 여부
        needs_review = transitional

        logger.info("─── [매칭 완료] %s ───", display_name)

        return MatchResult(
            display_name=display_name,
            selected=primary,
            prev_version=prev_version,
            transitional_flag=transitional,
            transitional_text=excerpt,
            relevant_articles=relevant_articles,
            needs_user_review=needs_review,
        )

    def match_all(
        self,
        bid_date: date,
        laws: list[tuple[str, str, str]],
        progress_callback: Optional[Callable[[int, int, str], None]] = None,
    ) -> list[MatchResult]:
        """대상 법령 전체 일괄 매칭

        Args:
            bid_date:          입찰공고일
            laws:              [(display_name, query, target), ...]
            progress_callback: (현재 인덱스, 전체 수, 현재 법령명) 콜백

        Returns:
            MatchResult 목록 (laws 순서 동일)
        """
        results: list[MatchResult] = []
        total = len(laws)

        for i, (display_name, query, target) in enumerate(laws):
            if progress_callback:
                progress_callback(i, total, display_name)
            result = self.match_one(display_name, query, target, bid_date)
            results.append(result)

        if progress_callback:
            progress_callback(total, total, "완료")

        return results
