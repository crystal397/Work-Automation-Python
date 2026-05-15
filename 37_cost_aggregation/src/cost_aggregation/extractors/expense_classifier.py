"""경비 11비목 자동 분류 (보고서 4.3.3.1 직접계상비목).

파일명·폴더명 휴리스틱 매칭. 영수증·전표 PDF 의 본문 OCR 까지 자동화하려면
[[reference_korean_ocr_models]] 와 결합 (별도 후속 작업).

비목 분류 우선순위는 큰 게 이김 — 더 구체적인 패턴이 일반 패턴을 덮어쓴다.
"""
from __future__ import annotations

import re
from dataclasses import dataclass
from pathlib import Path


# (정규식, 비목명, 우선순위) — priority 큰 게 이김
# 보고서 4.3.3.1 표준 11비목 + 추가 세부 분류
_EXPENSE_RULES: list[tuple[re.Pattern, str, int]] = [
    # 가장 구체적 (priority 100~) ─ 회사·계좌·이체 등 회계 부속 자료는 제외
    (re.compile(r"보증\s*수수료|보증료|이행보증|선급보증|하자보증|지급보증|대금\s*지급\s*보증"), "보증수수료", 110),
    (re.compile(r"산재보험|고용보험|근재보험|산업재해|건강보험|국민연금|4대\s*보험"), "보험료", 105),
    (re.compile(r"자동차종합|자동차\s*보험|손해\s*배상|영업\s*배상|배상책임"), "보험료(차량/배상)", 100),
    (re.compile(r"안전관리비|안전\s*시설|안전모|안전화|안전대|안전벨트|안전\s*교육|보호구|소방"), "안전관리비", 95),
    (re.compile(r"복리\s*후생|식대|식비|회식|체육|간식|건강\s*검진|체력단련|단합대회"), "복리후생비", 90),
    (re.compile(r"여비\s*교통|출장비?|교통비|주차료|톨게이트|유류대?|주유"), "여비교통비", 85),
    (re.compile(r"통신비?|전화비?|인터넷|핸드폰|휴대폰|모바일|팩스|우편"), "통신비", 80),
    (re.compile(r"광고\s*선전|광고비?|홍보비?|간판"), "광고선전비", 75),
    (re.compile(r"회의비?|회의\s*실|미팅"), "회의비", 70),
    (re.compile(r"세금\s*공과|면허세|등록세|취득세|재산세|공과금|관리비\s*(?:납부|영수증)"), "세금과공과", 65),
    # 운반·기계·외주 (현장 시공 관련)
    (re.compile(r"운반비|운송비?|배송|택배|화물|운임"), "운반비", 60),
    (re.compile(r"기계\s*경비|장비\s*비?|장비\s*임차|크레인|굴착기|덤프|레미콘|믹서"), "기계경비", 55),
    (re.compile(r"외주\s*가공|가공비?"), "외주가공비", 50),
    # 가설비 (현장사무소·숙소·식당·창고)
    (
        re.compile(r"현장\s*사무소?|숙소|식당|취사장|창고|컨테이너|가설\s*건물|가설\s*전기|가설\s*수도"),
        "가설비",
        45,
    ),
    # 임차료 (월세·렌트·임대) — 일반
    (re.compile(r"월세|임대차|임차료|렌트(?!카드)|리스(?!트)"), "지급임차료", 40),
    # 보관비
    (re.compile(r"보관비?|적치|저장"), "보관비", 35),
    # 전력·수도·광열 — 일반 (마지막 폴백 가까이)
    (re.compile(r"전력|전기료|수도료?|가스(?!비)|광열|난방|냉방|TV\s*수신"), "전력수도광열비", 30),
    # 소모품·도서·인쇄
    (re.compile(r"소모품|문구|도서|인쇄|복사|토너|용지"), "소모품비", 25),
    (re.compile(r"수선|보수|수리|정비"), "수선비", 20),
]


@dataclass
class ExpenseHit:
    path: Path
    item: str            # 비목 이름
    priority: int
    reason: str          # 매칭된 키워드


def classify_expense(path: Path) -> ExpenseHit:
    """파일/폴더명 → 경비 비목 매칭.

    매칭 없으면 ``item="(미분류)"`` 반환 → 사람 검토 필요.
    """
    haystack = " ".join([path.name, *[p.name for p in path.parents[:4]]])
    best: ExpenseHit | None = None
    for pat, item, prio in _EXPENSE_RULES:
        m = pat.search(haystack)
        if not m:
            continue
        if best is None or prio > best.priority:
            best = ExpenseHit(path=path, item=item, priority=prio, reason=m.group(0))
    if best is None:
        return ExpenseHit(path=path, item="(미분류)", priority=0, reason="(매칭 없음)")
    return best


def classify_directory(
    root: Path,
    *,
    extensions: tuple[str, ...] = (".pdf", ".jpg", ".jpeg", ".png", ".xlsx"),
) -> dict[str, list[Path]]:
    """폴더 안 영수증·전표·증빙 PDF 들을 비목별로 분류해 dict 반환.

    반환 예::

        {
          "전력수도광열비": [Path("..."), ...],
          "통신비":         [Path("..."), ...],
          "(미분류)":       [Path("..."), ...],
        }
    """
    by_item: dict[str, list[Path]] = {}
    for p in root.rglob("*"):
        if not p.is_file():
            continue
        if p.suffix.lower() not in extensions:
            continue
        hit = classify_expense(p)
        by_item.setdefault(hit.item, []).append(p)
    return by_item
