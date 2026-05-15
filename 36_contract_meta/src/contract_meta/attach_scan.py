"""수신자료 폴더 스캔 → appendix.yaml 자동 생성.

파일명·폴더명 패턴 매칭으로 5.1 ~ 5.5 자동 분류. 사람이 검토·수정 후 `contract-meta attach` 실행.
"""

from __future__ import annotations

import re
from dataclasses import dataclass, field
from pathlib import Path


# (정규식, 섹션명, 우선순위) — 위에서부터 매칭, 우선순위 큰 게 이김
# 패턴 추가 시 A공구(도시철도)·B공구(고속도로)의 실제 파일명 패턴을 참고해 확장.
_RULES: list[tuple[re.Pattern, str, int]] = [
    # 5.2 — 계약·합의 문서 (가장 구체적)
    # `변경\s*\(YY\.MM\.DD\)` — C공구(철도)식 변경계약서 파일명 ("변경 (24.12.02) - 차수(제1차) 1회(공기연장).pdf")
    (
        re.compile(
            r"변경계약서|최초.*계약서|도급계약서|용역\s*계약서?|착공계|준공계|"
            r"합의각서|이행각서|확약서|공동\s*수급|"
            r"변경\s*\(\d{2,4}\.\d{1,2}\.\d{1,2}\)"
        ),
        "5.2. 계약문서",
        100,
    ),
    (
        re.compile(
            r"계약현황|용역\s*과업|특수조건|일반조건|입찰\s*공고서?|현장\s*설명서?|"
            r"입찰\s*안내서?|과업\s*지시서|공사\s*설명회"
        ),
        "5.2. 계약문서",
        90,
    ),
    # 5.3 — 수발신 공문류 (승인·통보·실정보고·요청 등)
    # 주의: 짧은 일반 명사(진정·보고 등)는 회사명·일반 단어에 substring 매칭되므로
    #       반드시 한정형(`진정서`, `보고서`)으로 사용.
    (
        re.compile(
            r"공문|수발신|문서번호|결의서|허가\s*알림|검토서|회신|승인\s*통보|실정\s*보고|"
            r"통보문?|협의서|회의록|업무\s*일지|보고서|품의서|기안문|민원|진정서|"
            r"이의\s*유보|연장\s*요청|연장\s*승인|"
            r"재정\s*집행|공정\s*회의"
        ),
        "5.3. 수발신문서",
        80,
    ),
    # 5.4 — 법령·예규·판례
    (
        re.compile(r"법령|예규|국가계약|지방계약|판례|판결서?|판결문|법률\s*라운지|입찰유의서|회계예규"),
        "5.4. 적용 법령",
        70,
    ),
    # 5.5 노무 — 인원 증빙·급여 (경력증명서·재직증명서·4대보험 등)
    (
        re.compile(
            r"노무비|급여|임금|인사카드|인원투입|지불조서|명세서|경력증명서?|재직증명서?|"
            r"퇴직|국민연금|건강보험|고용보험|산재보험|4대\s*보험|원천징수|발령|임명|"
            r"근로계약서|급여대장|급여지급|급여\s*명세|작업\s*일보"
        ),
        "5.5. 공기연장 간접비 증빙자료 (노무)",
        60,
    ),
    # 5.5 경비 — 비목별 증빙 (숙소·차량·임대차·이체확인증 등)
    (
        re.compile(
            r"경비|영수증|세금계산서|법인카드|전표|복리|소모|임차|수도|전력|통신|건강검진|"
            r"숙소|월세|차량|렌트|렌탈|관리비|이체\s*확인증|보증\s*수수료|임대차계약서?|"
            r"유류|주유|식대|회의비|광고비|운반비"
        ),
        "5.5. 공기연장 간접비 증빙자료 (경비)",
        50,
    ),
    # 5.1 — 산출·산정 근거 (구체)
    (
        re.compile(
            r"수량산출|단가산출|산출내역|원가계산|산정\s*기준|산정\s*근거|일수\s*산정|"
            r"적용\s*요율|원가\s*실적|단가\s*설명|공정표|견적서?|설계변경|설계도서|"
            r"암판정|장비\s*대기"
        ),
        "5.1. 공기연장 간접비 산정근거",
        40,
    ),
    # 5.1 — 일반 (요약·집계 — 가장 마지막)
    (re.compile(r"요약|집계"), "5.1. 공기연장 간접비 산정근거", 30),
]


@dataclass
class _Hit:
    path: Path
    section: str
    priority: int
    reason: str


def classify(path: Path) -> _Hit:
    name_only = path.name
    parents_only = " ".join(p.name for p in path.parents[:4])
    best: _Hit | None = None
    for pat, sec, prio in _RULES:
        m = pat.search(name_only)
        if not m:
            continue
        if best is None or prio > best.priority:
            best = _Hit(path=path, section=sec, priority=prio, reason=m.group(0))
    if best is not None:
        return best
    for pat, sec, prio in _RULES:
        m = pat.search(parents_only)
        if not m:
            continue
        if best is None or prio > best.priority:
            best = _Hit(path=path, section=sec, priority=prio, reason=m.group(0))
    if best is None:
        return _Hit(path=path, section="5.1. 공기연장 간접비 산정근거", priority=0, reason="(분류 안 됨 — 기본)")
    return best


def scan_to_appendix(
    root: Path,
    *,
    project_root: Path | None = None,
    recursive: bool = True,
    extensions: tuple[str, ...] = (".pdf", ".docx"),
    absolute_paths: bool = False,
) -> dict:
    """루트 폴더 스캔 → appendix.yaml 구조 dict 반환.

    absolute_paths=True 면 path 를 절대경로로 저장 (attach 호출 시 cwd 무관).
    기본 False — project_root 기준 상대경로 (이식성 우선).
    """
    root = root.resolve()
    project_root = (project_root or root).resolve()
    sections: dict[str, list[dict]] = {}
    other: list[Path] = []

    iterator = root.rglob("*") if recursive else root.iterdir()
    for p in iterator:
        if not p.is_file():
            continue
        if p.suffix.lower() not in extensions:
            continue
        hit = classify(p)
        if absolute_paths:
            path_str = str(p)
        else:
            try:
                rel = p.relative_to(project_root)
            except ValueError:
                rel = p
            path_str = str(rel)
        sections.setdefault(hit.section, []).append({
            "path": path_str.replace("\\", "/"),
            "_match": hit.reason,
        })

    spec = {"sections": []}
    for sec in sorted(sections.keys()):
        files = sorted(sections[sec], key=lambda x: x["path"])
        spec["sections"].append({"title": sec, "files": files})
    return spec
