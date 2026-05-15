"""calc.py 인원별 ⑤⑥⑦ 산식 회귀 테스트.

보고서 4.3.2.2 의 6개 컬럼 (① 급여 / ② 일수 / ③ 퇴직 / ④ 소계 / ⑤ 추정일수 /
⑥ 1일평균 / ⑦ 추정소계 / C 합계) 자동 산식 검증.
"""
from __future__ import annotations

import sys
from pathlib import Path

_ROOT = Path(__file__).resolve().parent.parent / "src"
if str(_ROOT) not in sys.path:
    sys.path.insert(0, str(_ROOT))

from cost_aggregation import calc


def test_severance():
    # ③ 퇴직급여충당금 = 급여 / 12 (반올림)
    cv = calc.calc_severance(12_000_000)
    assert cv.value == 1_000_000, f"퇴직={cv.value}"
    # 12 로 나누어 떨어지지 않을 때
    cv = calc.calc_severance(10_000_000)
    assert cv.value == 833_333, f"퇴직={cv.value}"


def test_subtotal_actual():
    # ④ 실비 소계 = ① + ③
    cv = calc.calc_subtotal_actual(salary_actual=12_000_000, severance=1_000_000)
    assert cv.value == 13_000_000


def test_daily_rate():
    # ⑥ 1일평균 = ④ / ②
    cv = calc.calc_daily_rate(subtotal_actual=13_000_000, work_days_actual=30)
    assert cv.value == 433_333, f"1일평균={cv.value}"

    # ② == 0 보호
    cv = calc.calc_daily_rate(subtotal_actual=13_000_000, work_days_actual=0)
    assert cv.value == 0


def test_estimate_subtotal():
    # ⑦ 추정 소계 = ⑤ × ⑥
    cv = calc.calc_estimate_subtotal(days_estimate=95, daily_rate=433_333)
    assert cv.value == 95 * 433_333


def test_total_salary():
    # C 합계 = ④ + ⑦
    cv = calc.calc_total_salary(subtotal_actual=13_000_000, subtotal_estimate=41_166_635)
    assert cv.value == 54_166_635


def test_chain_full_row():
    """4.3.2.2 한 행 전체 체인 — 보고서의 ①~C 흐름 검증.

    예시: 우동석 소장 (C공구(철도) 케이스 모사)
    - ① 급여 = 55,850,000  → 큰 단위
    - ② 일수 = 30
    - ⑤ 추정일수 = 95
    """
    salary_actual = 55_850_000
    work_days = 30
    days_est = 95

    severance = calc.calc_severance(salary_actual)
    assert severance.value == 4_654_167  # 55,850,000 / 12 반올림

    subtotal_actual = calc.calc_subtotal_actual(salary_actual, severance.value)
    assert subtotal_actual.value == 60_504_167

    daily_rate = calc.calc_daily_rate(subtotal_actual.value, work_days)
    assert daily_rate.value == 2_016_806  # 60,504,167 / 30 반올림

    estimate_sub = calc.calc_estimate_subtotal(days_est, daily_rate.value)
    assert estimate_sub.value == 95 * 2_016_806

    total = calc.calc_total_salary(subtotal_actual.value, estimate_sub.value)
    assert total.value == subtotal_actual.value + estimate_sub.value


if __name__ == "__main__":
    funcs = [test_severance, test_subtotal_actual, test_daily_rate,
             test_estimate_subtotal, test_total_salary, test_chain_full_row]
    failed = 0
    for f in funcs:
        try:
            f()
            print(f"  PASS {f.__name__}")
        except AssertionError as e:
            print(f"  FAIL {f.__name__}: {e}")
            failed += 1
    print(f"\n{len(funcs) - failed}/{len(funcs)} passed")
    sys.exit(0 if failed == 0 else 1)
