"""핵심 매칭 엔진 — 입찰공고일 기준 6단계 시행일 판단 로직

6단계:
  1. 연혁 조회     — 해당 법령 전체 버전(공포번호·시행일) 수집
  2. 1차 후보 선정 — 시행일 ≤ 입찰공고일 중 가장 최근 버전
  3. 부칙 경과규정 — 1차 후보 부칙 파싱, 경과규정 패턴 탐지 (유형 A/B 구분)
  4. 2차 후보 선정 — 경과규정 존재 시 직전 버전을 병렬 후보로 제시
  5. 상위·하위법 정합성 — law/시행령/시행규칙 시행일 비교 경고
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
# 유형 A: 법령 전체 적용 경과규정 ("이 법 시행 전에 공고된 입찰 → 종전 규정 적용")
_TRANSITIONAL_A_PATTERNS: list[re.Pattern] = [
    re.compile(r"시행\s*전에?\s*(이미\s*)?(입찰\s*공고|공고)"),
    re.compile(r"종전의?\s*규정에\s*따른다"),
    re.compile(r"시행\s*당시\s*종전"),
    re.compile(r"이\s*(법|영|규칙)\s*시행\s*전.*?(입찰|공고|계약)"),
    re.compile(r"공고된\s*계약.*?종전"),
]

# 유형 B: 특정 조문 단위 경과규정 ("제○조 개정규정은 시행 후 최초 공고 분부터")
_TRANSITIONAL_B_PATTERNS: list[re.Pattern] = [
    re.compile(r"개정규정은\s*이\s*(법|영|규칙)\s*시행\s*후\s*최초"),
    re.compile(r"최초로\s*입찰\s*공고하는?\s*분부터"),
]

# 유형 B 경과규정에서 영향받는 조번호 추출
_ARTICLE_NUM_PATTERN = re.compile(r"제\s*(\d+)\s*조")

# 상위·하위법 정합성 확인 그룹 (display_name 기준 접두사)
_LAW_FAMILY_GROUPS: list[list[str]] = [
    ["국가계약법", "국가계약법 시행령", "국가계약법 시행규칙"],
    ["지방계약법", "지방계약법 시행령", "지방계약법 시행규칙"],
    ["하도급법", "하도급법 시행령"],
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
    transitional_type: str = ""           # 경과규정 유형: "A" | "B" | "" (미탐지)
    transitional_articles: list[str] = field(default_factory=list)  # 유형 B: 영향 조번호
    relevant_articles: list[dict] = field(default_factory=list)
    needs_user_review: bool = False       # 사용자 확인 필요
    warning: str = ""                     # 경고 메시지
    consistency_warning: str = ""         # 5단계: 상위·하위법 시행일 불일치 경고


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

    def _detect_transitional(self, text: dict) -> tuple[bool, str, str, list[str]]:
        """부칙에서 경과규정 패턴 탐지 → (발견 여부, 발췌문, 유형("A"|"B"|""), 영향 조번호 목록)

        유형 A: 법령 전체 적용 — "이 법 시행 전 공고 → 종전 규정 적용"
        유형 B: 조문 단위 적용 — "제○조 개정규정은 시행 후 최초 공고 분부터"
        """
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

        def _excerpt(m: re.Match) -> str:
            start = max(0, m.start() - 20)
            end = min(len(full_text), m.end() + 80)
            return full_text[start:end].strip()

        # 유형 B 먼저 확인 (더 구체적인 패턴)
        for pattern in _TRANSITIONAL_B_PATTERNS:
            m = pattern.search(full_text)
            if m:
                excerpt = _excerpt(m)
                # 발췌문에서 영향받는 조번호 추출
                # 발췌 앞 문맥(최대 100자)까지 포함해 조번호 탐색
                context_start = max(0, m.start() - 100)
                context = full_text[context_start: m.end() + 20]
                art_nums = _ARTICLE_NUM_PATTERN.findall(context)
                logger.warning(
                    "경과규정 유형 B 탐지 — 패턴: %s | 영향 조: %s | 발췌: %s",
                    pattern.pattern, art_nums, excerpt[:80],
                )
                return True, excerpt, "B", art_nums

        # 유형 A 확인
        for pattern in _TRANSITIONAL_A_PATTERNS:
            m = pattern.search(full_text)
            if m:
                excerpt = _excerpt(m)
                logger.warning(
                    "경과규정 유형 A 탐지 — 패턴: %s | 발췌: %s", pattern.pattern, excerpt[:80]
                )
                return True, excerpt, "A", []

        return False, "", "", []

    def _filter_articles(self, text: dict) -> list[dict]:
        """공기연장 키워드가 포함된 조문 필터링 (호(號) 단위까지 추출)"""
        articles = []
        article_units = (text.get("조문") or {}).get("조문단위", [])
        if isinstance(article_units, dict):
            article_units = [article_units]

        for unit in article_units:
            title = str(unit.get("조제목") or "")
            content = str(unit.get("조문내용") or "")

            paragraphs_raw = unit.get("항") or []
            if isinstance(paragraphs_raw, dict):
                paragraphs_raw = [paragraphs_raw]

            # 항 정규화 + 호(號) 추출
            paragraphs: list[dict] = []
            for para in paragraphs_raw:
                sub_items_raw = para.get("호") or []
                if isinstance(sub_items_raw, dict):
                    sub_items_raw = [sub_items_raw]
                sub_items = [
                    {
                        "호번호": str(s.get("호번호") or s.get("subNo") or ""),
                        "호내용": str(s.get("호내용") or s.get("subContent") or ""),
                    }
                    for s in sub_items_raw
                ]
                paragraphs.append({
                    "항번호": str(para.get("항번호") or ""),
                    "항내용": str(para.get("항내용") or ""),
                    "호": sub_items,
                })

            para_text = " ".join(p["항내용"] for p in paragraphs)
            sub_text = " ".join(
                s["호내용"] for p in paragraphs for s in p["호"]
            )

            combined = title + content + para_text + sub_text
            if any(kw in combined for kw in config.EXTENSION_KEYWORDS):
                articles.append({
                    "조번호": str(unit.get("조번호") or ""),
                    "조제목": title,
                    "조문내용": content,
                    "항": paragraphs,
                })

        return articles

    # ── 6단계 매칭 로직 ────────────────────────────────────────────────────────

    def _admrul_history(self, mst: str, query: str) -> list[LawVersion]:
        """행정규칙 연혁 버전 목록 수집 — 3단계 fallback.

        1단계: lawHistory.do (admrul) — API가 지원하면 가장 정확
        2단계: search_law(display=100) — 다수 버전이 검색 결과에 포함될 수 있음
        3단계: 빈 목록 반환 (호출자가 현행 버전으로 fallback)
        """
        # 1단계: lawHistory.do admrul 시도
        try:
            versions = self._get_history(mst, "admrul")
            if versions:
                logger.info("[행정규칙 연혁] lawHistory.do 성공: %d건", len(versions))
                return versions
        except Exception as exc:
            logger.debug("[행정규칙 연혁] lawHistory.do 실패: %s", exc)

        # 2단계: search_law 광범위 검색으로 여러 버전 수집
        try:
            raw_list = self.client.search_law(query, target="admrul", display=100)
            versions = [
                v for raw in raw_list
                if (v := _raw_to_version(
                    raw,
                    str(raw.get("법령MST") or raw.get("MST") or mst),
                    "admrul",
                ))
            ]
            versions.sort(key=lambda v: v.enforce_date)
            if versions:
                logger.info(
                    "[행정규칙 연혁] search_law fallback 성공: %d건", len(versions)
                )
                return versions
        except Exception as exc:
            logger.debug("[행정규칙 연혁] search_law fallback 실패: %s", exc)

        return []

    def _match_admrul(
        self,
        display_name: str,
        query: str,
        bid_date: date,
    ) -> MatchResult:
        """행정규칙 전용 매칭.

        lawHistory.do → search_law 다중 결과 → 현행 버전 순으로 연혁을 확보하고
        입찰공고일 기준 가장 최근 시행 버전을 선정한다.
        연혁 확보에 실패한 경우 현행 버전을 표시하고 수동 확인을 요청한다.
        """
        # MST + 현행 버전 1건 우선 확보
        laws = self.client.search_law(query, target="admrul", display=10)
        if not laws:
            return MatchResult(
                display_name=display_name,
                selected=None, prev_version=None,
                transitional_flag=False, transitional_text="",
                warning=f"행정규칙 조회 실패: '{query}'",
                needs_user_review=True,
            )

        raw_current = next(
            (l for l in laws if str(l.get("법령명한글") or l.get("법령명") or "") == query),
            None,
        ) or laws[0]
        mst = str(raw_current.get("법령MST") or raw_current.get("MST") or "")

        # 연혁 수집 시도
        all_versions = self._admrul_history(mst, query)

        if all_versions:
            # 입찰공고일 이전 시행 버전 필터
            candidates = [v for v in all_versions if v.enforce_date <= bid_date]
            if candidates:
                primary = candidates[-1]
                prev_version = candidates[-2] if len(candidates) >= 2 else None
                history_source = "lawHistory" if len(all_versions) > 1 else "search"
                logger.info(
                    "[행정규칙] '%s' — 연혁 %d건 확보(%s), 선정: %s (시행일 %s)",
                    display_name, len(all_versions), history_source,
                    primary.announce_num, primary.enforce_date,
                )

                text = self._get_text(primary)
                transitional, excerpt, trans_type, trans_articles = self._detect_transitional(text)
                relevant_articles = self._filter_articles(text)

                return MatchResult(
                    display_name=display_name,
                    selected=primary,
                    prev_version=prev_version,
                    transitional_flag=transitional,
                    transitional_text=excerpt,
                    transitional_type=trans_type,
                    transitional_articles=trans_articles,
                    relevant_articles=relevant_articles,
                    needs_user_review=transitional,
                    warning=(
                        "※ 행정규칙 연혁은 부분적으로만 제공될 수 있습니다. "
                        "실제 시행 버전을 추가 확인하세요."
                    ) if history_source == "search" else "",
                )

        # 연혁 확보 실패 — 현행 버전으로 fallback
        logger.warning("[행정규칙] '%s' — 연혁 확보 실패, 현행 버전 표시", display_name)
        version = _raw_to_version(raw_current, mst, "admrul")
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
                "⚠ 행정규칙 연혁 조회 불가 — 현행 버전 기준 표시. "
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

        # 3단계 — 부칙 경과규정 확인 (유형 A/B 구분)
        text = self._get_text(primary)
        transitional, excerpt, trans_type, trans_articles = self._detect_transitional(text)

        # 4단계 — 경과규정 존재 시 경고 + 직전 버전 병렬 제시
        if transitional:
            type_label = f"유형 {trans_type}" if trans_type else "유형 불명"
            logger.warning(
                "부칙 경과규정 탐지(%s) → 직전 버전(%s) 병렬 제시, 사용자 확인 필요",
                type_label,
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
            transitional_type=trans_type,
            transitional_articles=trans_articles,
            relevant_articles=relevant_articles,
            needs_user_review=needs_review,
        )

    def _check_family_consistency(self, results: list[MatchResult]) -> None:
        """5단계: 상위·하위법(법률/시행령/시행규칙) 시행일 정합성 확인.

        같은 법령 계열(예: 국가계약법 / 시행령 / 시행규칙)의 시행일이
        서로 다른 경우 각 MatchResult에 consistency_warning을 추가한다.
        results 목록을 in-place 수정한다.
        """
        result_by_name = {r.display_name: r for r in results}

        for family in _LAW_FAMILY_GROUPS:
            members = [result_by_name[name] for name in family if name in result_by_name]
            # 시행일을 확인할 수 있는 것만
            dated = [(m, m.selected.enforce_date) for m in members if m.selected]
            if len(dated) < 2:
                continue

            dates = {d for _, d in dated}
            if len(dates) == 1:
                continue  # 모두 동일 — 정상

            # 불일치: 각 결과에 경고 추가
            date_summary = ", ".join(
                f"{m.display_name}={d}" for m, d in dated
            )
            warn_msg = f"⚠ 상위·하위법 시행일 불일치 — {date_summary}"
            logger.warning("정합성 경고: %s", warn_msg)
            for m, _ in dated:
                m.consistency_warning = warn_msg
                m.needs_user_review = True

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

        # 5단계 — 상위·하위법 정합성 확인 (전체 결과 수집 후 일괄 처리)
        self._check_family_consistency(results)

        if progress_callback:
            progress_callback(total, total, "완료")

        return results
