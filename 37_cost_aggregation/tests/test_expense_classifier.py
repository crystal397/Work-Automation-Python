"""expense_classifier.classify_expense 회귀 테스트.

A공구(도시철도) 의 경비 PDF 파일명 패턴을 기반으로 4.3.3.1 비목 분류 정확도 검증.

새 양식 발견 시:
1. CASES 에 (path, expected_item) 추가
2. 실패하면 expense_classifier._EXPENSE_RULES 보강
3. 회귀 테스트 통과 후 커밋
"""
from __future__ import annotations

import sys
from pathlib import Path

_ROOT = Path(__file__).resolve().parent.parent / "src"
if str(_ROOT) not in sys.path:
    sys.path.insert(0, str(_ROOT))

from cost_aggregation.extractors.expense_classifier import classify_expense


CASES = [
    # ── 보증수수료 ──
    ("samples/A공구(도시철도)/250131_하도급대금지급보증_영수증_2403.pdf", "보증수수료"),
    ("samples/A공구(도시철도)/이행보증료_2403.pdf", "보증수수료"),
    # ── 보험료 ──
    ("samples/A공구(도시철도)/근재보험료 전표_446,347.pdf", "보험료"),
    ("samples/A공구(도시철도)/산재보험료_2501.pdf", "보험료"),
    ("samples/A공구(도시철도)/고용보험료/2501_급여명세서.pdf", "보험료"),
    ("samples/A공구(도시철도)/자동차종합보험_차량_A.pdf", "보험료(차량/배상)"),
    # ── 안전관리비 ──
    ("samples/A공구(도시철도)/안전관리비/안전모 구매_240315.pdf", "안전관리비"),
    ("samples/A공구(도시철도)/안전교육비_2501.pdf", "안전관리비"),
    # ── 복리후생비 ──
    ("samples/A공구(도시철도)/복리후생비/식대 영수증_2502.pdf", "복리후생비"),
    ("samples/A공구(도시철도)/건강검진_2502.pdf", "복리후생비"),
    # ── 여비교통비 ──
    ("samples/A공구(도시철도)/여비교통비/출장비_2503.pdf", "여비교통비"),
    ("samples/A공구(도시철도)/주유비_2502.pdf", "여비교통비"),
    # ── 통신비 ──
    ("samples/A공구(도시철도)/250226_통신비_OOO,OOO.pdf", "통신비"),
    ("samples/A공구(도시철도)/휴대폰요금_2502.pdf", "통신비"),
    # ── 운반비 ──
    ("samples/A공구(도시철도)/운반비/화물운송_2503.pdf", "운반비"),
    # ── 가설비 ──
    ("samples/A공구(도시철도)/현장사무실/임대료_2502.pdf", "가설비"),
    ("samples/A공구(도시철도)/숙소/숙소_A 202호_월세계약서.pdf", "가설비"),
    # ── 지급임차료 ──
    ("samples/A공구(도시철도)/장비임차/굴착기.pdf", "기계경비"),  # "굴착기" 가 더 구체
    ("samples/A공구(도시철도)/오피스텔 월세_240213.pdf", "지급임차료"),
    # ── 전력수도광열비 ──
    ("samples/A공구(도시철도)/전력수도광열비/관계기관(전력)_2502.pdf", "전력수도광열비"),
    # ── 회의비 ──
    ("samples/A공구(도시철도)/회의비 영수증_2502.pdf", "회의비"),
]


def run_cases() -> tuple[int, int]:
    passed = failed = 0
    for path, expected in CASES:
        hit = classify_expense(Path(path))
        if hit.item == expected:
            passed += 1
        else:
            failed += 1
            print(f"  FAIL  expected={expected}  actual={hit.item} (reason={hit.reason!r})")
            print(f"        path={path}")
    return passed, failed


def test_expense_classifier_regression():
    p, f = run_cases()
    assert f == 0, f"{f} 케이스 실패 / {p} 통과"


if __name__ == "__main__":
    p, f = run_cases()
    print(f"\n총 {p+f}건 / 통과 {p} / 실패 {f}")
    sys.exit(0 if f == 0 else 1)
