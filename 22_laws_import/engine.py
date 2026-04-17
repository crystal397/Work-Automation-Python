"""핵심 매칭 엔진 — 입찰공고일 기준 6단계 시행일 판단 로직

6단계:
  1. 연혁 조회     — 해당 법령 전체 버전(공포번호·시행일) 수집
  2. 1차 후보 선정 — 시행일 ≤ 입찰공고일 중 가장 최근 버전
  3. 부칙 경과규정 — 1차 후보 부칙 파싱, 경과규정 패턴 탐지 (유형 A/B 구분)
  4. 2차 후보 선정 — 경과규정 존재 시 직전 버전을 병렬 후보로 제시
  5. 상위·하위법 정합성 — law/시행령/시행규칙 시행일 비교 경고
  6. 사용자 검토   — 자동 확정 불가 시 두 후보 모두 출력 + 플래그
"""
import json
import logging
import re
import subprocess
from dataclasses import dataclass, field
from datetime import date
from typing import Callable, Optional

from api_client import LawAPIClient
from cache import LawCache
from scraper import scrape_admrul_history
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

# 상위·하위법 정합성 확인 그룹 (display_name 기준)
# law 타입끼리만 비교 — admrul은 시행일이 달라도 정상이므로 별도 처리
_LAW_FAMILY_GROUPS: list[list[str]] = [
    ["국가계약법", "국가계약법 시행령", "국가계약법 시행규칙"],
    ["지방계약법", "지방계약법 시행령", "지방계약법 시행규칙"],
    ["하도급법", "하도급법 시행령"],
    ["조달사업에 관한 법률", "조달사업에 관한 법률 시행령"],
]

# 행정규칙 ↔ 상위 법령 연결 (시행일 불일치 시 soft warning 용도)
_ADMRUL_PARENT_MAP: dict[str, str] = {
    "공사계약일반조건": "국가계약법",
    "예정가격 작성기준": "국가계약법",
    "정부 입찰·계약 집행기준": "국가계약법",
    "지방자치단체 입찰 및 계약집행기준": "지방계약법",
}


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

        유형 B는 부칙단위별로 개별 처리하여 조번호 오추출을 방지한다.
        """
        sections = text.get("부칙") or {}
        units = sections.get("부칙단위", [])
        if isinstance(units, dict):
            units = [units]

        # 부칙단위별 텍스트 분리 (③ 개선: 단위별 개별 파싱으로 조번호 혼재 방지)
        unit_texts: list[str] = []
        for unit in units:
            content = unit.get("부칙내용") or ""
            if isinstance(content, list):
                content = " ".join(str(c) for c in content)
            unit_texts.append(str(content))

        full_text = " ".join(unit_texts)

        def _excerpt(text_src: str, m: re.Match) -> str:
            start = max(0, m.start() - 20)
            end = min(len(text_src), m.end() + 80)
            return text_src[start:end].strip()

        # 유형 B: 부칙단위별 개별 탐색 → 해당 단위 내 조번호만 추출
        for unit_text in unit_texts:
            for pattern in _TRANSITIONAL_B_PATTERNS:
                m = pattern.search(unit_text)
                if m:
                    excerpt = _excerpt(unit_text, m)
                    # 이 부칙단위 텍스트 내에서만 조번호 추출 (타 단위 혼재 방지)
                    art_nums = _ARTICLE_NUM_PATTERN.findall(unit_text)
                    logger.warning(
                        "경과규정 유형 B 탐지 — 패턴: %s | 영향 조: %s | 발췌: %s",
                        pattern.pattern, art_nums, excerpt[:80],
                    )
                    return True, excerpt, "B", art_nums

        # 유형 A: 전체 텍스트 탐색 (조번호 추출 불필요)
        for pattern in _TRANSITIONAL_A_PATTERNS:
            m = pattern.search(full_text)
            if m:
                excerpt = _excerpt(full_text, m)
                logger.warning(
                    "경과규정 유형 A 탐지 — 패턴: %s | 발췌: %s", pattern.pattern, excerpt[:80]
                )
                return True, excerpt, "A", []

        # regex 미탐지 → Claude Code CLI fallback (② 비정형 패턴 보완)
        return self._detect_transitional_claude(full_text)

    def _detect_transitional_claude(
        self, full_text: str
    ) -> tuple[bool, str, str, list[str]]:
        """Claude Code CLI를 사용한 경과규정 탐지 (regex 미탐지 시 fallback).

        `claude -p "..."` 를 subprocess로 호출한다.
        claude CLI가 설치되지 않은 경우 조용히 스킵한다.
        """
        if not full_text.strip():
            return False, "", "", []

        prompt = (
            "다음은 법령 부칙 텍스트입니다. "
            "공기연장 입찰공고일 기준 경과규정이 있는지 분석해주세요.\n\n"
            "[부칙 텍스트]\n"
            f"{full_text[:3000]}\n\n"
            "다음 JSON 형식으로만 응답하세요 (마크다운·설명 없이 순수 JSON만):\n"
            '{"found": true 또는 false, '
            '"type": "A" 또는 "B" 또는 "", '
            '"excerpt": "발견된 경과규정 문장 (없으면 빈 문자열)", '
            '"articles": ["1", "2"]}\n\n'
            "유형 A: 법령 전체 경과규정 — 예) "
            '"이 법 시행 전에 공고된 입찰에 대해서는 종전의 규정에 따른다"\n'
            "유형 B: 조문 단위 경과규정 — 예) "
            '"제○조의 개정규정은 시행 후 최초로 입찰 공고하는 분부터 적용"\n'
            "articles: 유형 B일 때 영향받는 조번호(숫자만), 유형 A이면 빈 배열"
        )

        try:
            result = subprocess.run(
                ["claude", "-p", prompt],
                capture_output=True,
                text=True,
                timeout=60,
                encoding="utf-8",
            )
            if result.returncode != 0 or not result.stdout.strip():
                return False, "", "", []

            output = result.stdout.strip()
            # 응답에 JSON 외 텍스트가 섞여도 JSON 블록만 추출
            json_match = re.search(r"\{[^{}]*\}", output, re.DOTALL)
            if not json_match:
                return False, "", "", []

            data = json.loads(json_match.group())
            found = bool(data.get("found", False))
            if found:
                trans_type = str(data.get("type", ""))
                excerpt = str(data.get("excerpt", ""))
                articles = [str(a) for a in data.get("articles", [])]
                logger.warning(
                    "경과규정 Claude Code 탐지 — 유형: %s | 발췌: %s",
                    trans_type, excerpt[:80],
                )
                return True, excerpt, trans_type, articles

        except FileNotFoundError:
            logger.debug("claude CLI 없음 — 비정형 경과규정 탐지 스킵")
        except subprocess.TimeoutExpired:
            logger.debug("claude CLI 타임아웃 — 경과규정 탐지 스킵")
        except (json.JSONDecodeError, Exception) as exc:
            logger.debug("Claude Code 응답 파싱 실패: %s", exc)

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

            # 항 정규화 + 호(號) + 목(目) 추출
            paragraphs: list[dict] = []
            for para in paragraphs_raw:
                sub_items_raw = para.get("호") or []
                if isinstance(sub_items_raw, dict):
                    sub_items_raw = [sub_items_raw]

                sub_items = []
                for s in sub_items_raw:
                    # 목(目) — 호 하위 단계
                    sub_sub_raw = s.get("목") or []
                    if isinstance(sub_sub_raw, dict):
                        sub_sub_raw = [sub_sub_raw]
                    sub_subs = [
                        {
                            "목번호": str(ss.get("목번호") or ss.get("subsubNo") or ""),
                            "목내용": str(ss.get("목내용") or ss.get("subsubContent") or ""),
                        }
                        for ss in sub_sub_raw
                    ]
                    sub_items.append({
                        "호번호": str(s.get("호번호") or s.get("subNo") or ""),
                        "호내용": str(s.get("호내용") or s.get("subContent") or ""),
                        "목": sub_subs,
                    })

                paragraphs.append({
                    "항번호": str(para.get("항번호") or ""),
                    "항내용": str(para.get("항내용") or ""),
                    "호": sub_items,
                })

            para_text = " ".join(p["항내용"] for p in paragraphs)
            sub_text = " ".join(
                s["호내용"] for p in paragraphs for s in p["호"]
            )
            subsub_text = " ".join(
                ss["목내용"]
                for p in paragraphs
                for s in p["호"]
                for ss in s["목"]
            )

            combined = title + content + para_text + sub_text + subsub_text
            if any(kw in combined for kw in config.EXTENSION_KEYWORDS):
                articles.append({
                    "조번호": str(unit.get("조번호") or ""),
                    "조제목": title,
                    "조문내용": content,
                    "항": paragraphs,
                })

        return articles

    # ── 6단계 매칭 로직 ────────────────────────────────────────────────────────

    def _admrul_history(self, mst: str, query: str) -> tuple[list[LawVersion], str]:
        """행정규칙 연혁 버전 목록 수집 — 4단계 fallback.

        Returns:
            (versions, source) — source: "api" | "scraper" | "search" | ""
        """
        # 1단계: lawHistory.do admrul 시도
        try:
            versions = self._get_history(mst, "admrul")
            if versions:
                logger.info("[행정규칙 연혁] lawHistory.do 성공: %d건", len(versions))
                return versions, "api"
        except Exception as exc:
            logger.debug("[행정규칙 연혁] lawHistory.do 실패: %s", exc)

        # 2단계: 웹 스크래핑 (beautifulsoup4 필요)
        try:
            raw_list = scrape_admrul_history(mst)
            versions = [
                v for raw in raw_list
                if (v := _raw_to_version(raw, mst, "admrul"))
            ]
            versions.sort(key=lambda v: v.enforce_date)
            if versions:
                logger.info("[행정규칙 연혁] 스크래핑 성공: %d건", len(versions))
                return versions, "scraper"
        except Exception as exc:
            logger.debug("[행정규칙 연혁] 스크래핑 실패: %s", exc)

        # 3단계: search_law 광범위 검색으로 여러 버전 수집
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
                return versions, "search"
        except Exception as exc:
            logger.debug("[행정규칙 연혁] search_law fallback 실패: %s", exc)

        return [], ""

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

        # 연혁 수집 시도 (B1 수정: source 명시 추적)
        all_versions, history_source = self._admrul_history(mst, query)

        if all_versions:
            # 입찰공고일 이전 시행 버전 필터
            candidates = [v for v in all_versions if v.enforce_date <= bid_date]

            # B2 수정: 연혁은 있으나 입찰공고일 이전 버전이 없는 경우 명시
            if not candidates:
                logger.warning(
                    "[행정규칙] '%s' — 연혁 %d건 확보했으나 입찰공고일(%s) 이전 시행 버전 없음",
                    display_name, len(all_versions), bid_date,
                )
                earliest = all_versions[0]
                return MatchResult(
                    display_name=display_name,
                    selected=earliest, prev_version=None,
                    transitional_flag=False, transitional_text="",
                    warning=(
                        f"⚠ 행정규칙 연혁 {len(all_versions)}건 확보했으나 "
                        f"입찰공고일({bid_date}) 이전 시행 버전 없음 — "
                        f"가장 오래된 버전(시행일 {earliest.enforce_date}) 표시"
                    ),
                    needs_user_review=True,
                )

            primary = candidates[-1]
            prev_version = candidates[-2] if len(candidates) >= 2 else None
            logger.info(
                "[행정규칙] '%s' — 연혁 %d건 확보(소스: %s), 선정: %s (시행일 %s)",
                display_name, len(all_versions), history_source,
                primary.announce_num, primary.enforce_date,
            )

            text = self._get_text(primary)
            transitional, excerpt, trans_type, trans_articles = self._detect_transitional(text)
            relevant_articles = self._filter_articles(text)

            # 소스별 warning 메시지
            source_warning = {
                "api": "",
                "scraper": "※ 행정규칙 연혁은 웹 스크래핑으로 확보 — 실제 시행 버전 추가 확인 권장",
                "search": "※ 행정규칙 연혁은 검색 결과 기반 (일부 누락 가능) — 실제 시행 버전 추가 확인 권장",
            }.get(history_source, "")

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
                warning=source_warning,
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
        """5단계: 상위·하위법 시행일 정합성 확인.

        [law 계열] 법률/시행령/시행규칙 간 시행일 불일치 → 경고 + 사용자 확인 요청
        [admrul 연결] 행정규칙 시행일이 상위 법령보다 오래된 경우 → soft warning
        results 목록을 in-place 수정한다.
        """
        result_by_name = {r.display_name: r for r in results}

        # ── law 계열 정합성 ──────────────────────────────────────────────────────
        for family in _LAW_FAMILY_GROUPS:
            members = [result_by_name[name] for name in family if name in result_by_name]
            dated = [(m, m.selected.enforce_date) for m in members if m.selected]
            if len(dated) < 2:
                continue

            dates = {d for _, d in dated}
            if len(dates) == 1:
                continue  # 모두 동일 — 정상

            date_summary = ", ".join(f"{m.display_name}={d}" for m, d in dated)
            warn_msg = f"⚠ 상위·하위법 시행일 불일치 — {date_summary}"
            logger.warning("정합성 경고: %s", warn_msg)
            for m, _ in dated:
                m.consistency_warning = warn_msg
                m.needs_user_review = True

        # ── admrul ↔ 상위 법령 soft warning ────────────────────────────────────
        for admrul_name, parent_name in _ADMRUL_PARENT_MAP.items():
            admrul_r = result_by_name.get(admrul_name)
            parent_r = result_by_name.get(parent_name)
            if not admrul_r or not parent_r:
                continue
            if not admrul_r.selected or not parent_r.selected:
                continue

            admrul_date = admrul_r.selected.enforce_date
            parent_date = parent_r.selected.enforce_date

            # 행정규칙 시행일이 상위 법령보다 1년 이상 오래된 경우 경고
            days_diff = (parent_date - admrul_date).days
            if days_diff > 365:
                warn_msg = (
                    f"※ {admrul_name} 시행일({admrul_date})이 "
                    f"상위 법령 {parent_name} 시행일({parent_date})보다 "
                    f"{days_diff // 365}년 이상 앞섬 — 행정규칙 갱신 여부 확인 권장"
                )
                logger.warning("admrul soft warning: %s", warn_msg)
                if not admrul_r.consistency_warning:
                    admrul_r.consistency_warning = warn_msg

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
