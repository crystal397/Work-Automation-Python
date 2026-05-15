"""renderer.kdate / renderer.kperiod 필터 회귀 테스트."""
from __future__ import annotations

import sys
from datetime import date, datetime
from pathlib import Path

_ROOT = Path(__file__).resolve().parent.parent / "src"
if str(_ROOT) not in sys.path:
    sys.path.insert(0, str(_ROOT))

from contract_meta.body.renderer import kdate, kperiod


KDATE_CASES = [
    # (입력, 스타일, 기대)
    ("2023-12-22", "dot", "2023. 12. 22."),
    ("2023-12-22", "long", "2023년 12월 22일"),
    ("2023-12-22", "compact", "2023.12.22"),
    ("2023.12.22", "dot", "2023. 12. 22."),
    ("2023/12/22", "dot", "2023. 12. 22."),
    ("20231222", "dot", "2023. 12. 22."),
    (date(2024, 8, 31), "dot", "2024. 8. 31."),
    (datetime(2024, 8, 31, 14, 30), "dot", "2024. 8. 31."),
    # 폴백 — 파싱 실패 시 원본 반환
    ("not-a-date", "dot", "not-a-date"),
    (None, "dot", ""),
    ("", "dot", ""),
]


KPERIOD_CASES = [
    # dot 스타일은 zero-pad 안 함: '2024. 8. 31.' (12월은 그냥 12)
    (("2023-12-22", "2024-08-31"), "2023. 12. 22. ~ 2024. 8. 31."),
    (("2023-12-22", None), "2023. 12. 22. ~"),
    ((None, "2024-08-31"), "~ 2024. 8. 31."),
    ((None, None), ""),
]


def test_kdate():
    failed = 0
    for value, style, expected in KDATE_CASES:
        actual = kdate(value, style)
        if actual != expected:
            print(f"  FAIL kdate({value!r}, {style!r}) = {actual!r} (expected {expected!r})")
            failed += 1
    assert failed == 0, f"{failed} kdate 실패"


def test_kperiod():
    failed = 0
    for (start, end), expected in KPERIOD_CASES:
        actual = kperiod(start, end)
        if actual != expected:
            print(f"  FAIL kperiod({start!r}, {end!r}) = {actual!r} (expected {expected!r})")
            failed += 1
    assert failed == 0, f"{failed} kperiod 실패"


if __name__ == "__main__":
    try:
        test_kdate()
        print("  PASS kdate")
    except AssertionError as e:
        print(f"  {e}")
    try:
        test_kperiod()
        print("  PASS kperiod")
    except AssertionError as e:
        print(f"  {e}")
