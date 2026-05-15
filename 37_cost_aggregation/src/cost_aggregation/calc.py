"""산식 계산 — 예정가격 작성기준에 따른 보험료·일반관리비·이윤·총원가·총액.

산식 (보고서 4.3 / 4.4 / 4.5 동일)
- 산재보험료    = 간접노무비 × industrial_accident_insurance_percent
- 고용보험료    = 간접노무비 × employment_insurance_percent
- 경비 합계     = 직접계상비목 + 산재보험료 + 고용보험료
- 소계          = 간접노무비 + 경비 합계
- 일반관리비    = 소계 × general_admin_percent
- 이윤          = (소계 + 일반관리비) × profit_percent
- 총원가        = 소계 + 일반관리비 + 이윤
- 총액(천원절사) = floor(총원가 / 1000) × 1000  (영세율인 경우 부가세 = 0)
"""

from __future__ import annotations

from contract_meta.models import ContractMeta
from cost_aggregation.models import ComputedValue


def calc_rate_based_insurance(labor_total: int, industrial_pct: float, employment_pct: float) -> tuple[ComputedValue, ComputedValue]:
    """산재·고용보험료 산정 (보고서 4.3.3.2 / 4.4.3.2)."""
    ind = int(round(labor_total * industrial_pct / 100))
    emp = int(round(labor_total * employment_pct / 100))
    return (
        ComputedValue(
            value=ind,
            formula=f"간접노무비×{industrial_pct}%",
            inputs=[("간접노무비", labor_total)],
            rate_percent=industrial_pct,
        ),
        ComputedValue(
            value=emp,
            formula=f"간접노무비×{employment_pct}%",
            inputs=[("간접노무비", labor_total)],
            rate_percent=employment_pct,
        ),
    )


def calc_expense_total(direct_total: int, industrial: int, employment: int) -> ComputedValue:
    return ComputedValue(
        value=direct_total + industrial + employment,
        formula="직접계상비목 + 산재보험료 + 고용보험료",
        inputs=[
            ("직접계상비목", direct_total),
            ("산재보험료", industrial),
            ("고용보험료", employment),
        ],
    )


def calc_subtotal(labor_total: int, expense_total: int) -> ComputedValue:
    return ComputedValue(
        value=labor_total + expense_total,
        formula="간접노무비 + 경비",
        inputs=[("간접노무비", labor_total), ("경비", expense_total)],
    )


def calc_general_admin(subtotal: int, rate_pct: float) -> ComputedValue:
    return ComputedValue(
        value=int(round(subtotal * rate_pct / 100)),
        formula=f"소계×{rate_pct}%",
        inputs=[("소계", subtotal)],
        rate_percent=rate_pct,
    )


def calc_profit(subtotal: int, general_admin: int, rate_pct: float) -> ComputedValue:
    base = subtotal + general_admin
    return ComputedValue(
        value=int(round(base * rate_pct / 100)),
        formula=f"(소계+일반관리비)×{rate_pct}%",
        inputs=[("소계", subtotal), ("일반관리비", general_admin)],
        rate_percent=rate_pct,
    )


def calc_gross(subtotal: int, general_admin: int, profit: int) -> ComputedValue:
    return ComputedValue(
        value=subtotal + general_admin + profit,
        formula="소계+일반관리비+이윤",
        inputs=[("소계", subtotal), ("일반관리비", general_admin), ("이윤", profit)],
    )


def calc_grand_total(gross: int, vat: int = 0, *, round_to: int = 1000) -> ComputedValue:
    raw = gross + vat
    rounded = (raw // round_to) * round_to
    return ComputedValue(
        value=rounded,
        formula=f"(총원가+부가세) 천원단위절사" if vat else "총원가 천원단위절사",
        inputs=[("총원가", gross), ("부가세", vat)],
    )


def calc_severance(salary_actual: int) -> ComputedValue:
    """③ 퇴직급여충당금 = 급여 / 12 (보고서 4.3.2.2)."""
    return ComputedValue(
        value=int(round(salary_actual / 12)),
        formula="급여(①) ÷ 12",
        inputs=[("급여(①)", salary_actual)],
    )


def calc_subtotal_actual(salary_actual: int, severance: int) -> ComputedValue:
    """④ 실비 소계 = ① + ③."""
    return ComputedValue(
        value=salary_actual + severance,
        formula="급여(①) + 퇴직급여충당금(③)",
        inputs=[("급여(①)", salary_actual), ("퇴직급여(③)", severance)],
    )


def calc_daily_rate(subtotal_actual: int, work_days_actual: int) -> ComputedValue:
    """⑥ 1일평균 노무비 = ④ / ② (소수점 반올림, 0일 보호).

    work_days_actual == 0 이면 ZeroDivisionError 대신 value=0 + 경고용 formula 반환.
    """
    if work_days_actual <= 0:
        return ComputedValue(
            value=0,
            formula="④ ÷ ② (② == 0 — 산정 불가, 입력 확인 필요)",
            inputs=[("실비소계(④)", subtotal_actual), ("실비일수(②)", work_days_actual)],
        )
    return ComputedValue(
        value=int(round(subtotal_actual / work_days_actual)),
        formula="실비소계(④) ÷ 실비일수(②)",
        inputs=[("실비소계(④)", subtotal_actual), ("실비일수(②)", work_days_actual)],
    )


def calc_estimate_subtotal(days_estimate: int, daily_rate: int) -> ComputedValue:
    """⑦ 추정 소계 = ⑤ × ⑥."""
    return ComputedValue(
        value=days_estimate * daily_rate,
        formula="추정일수(⑤) × 1일평균(⑥)",
        inputs=[("추정일수(⑤)", days_estimate), ("1일평균(⑥)", daily_rate)],
    )


def calc_total_salary(subtotal_actual: int, subtotal_estimate: int) -> ComputedValue:
    """C 인원별 합계 = ④ + ⑦."""
    return ComputedValue(
        value=subtotal_actual + subtotal_estimate,
        formula="실비소계(④) + 추정소계(⑦)",
        inputs=[("실비소계(④)", subtotal_actual), ("추정소계(⑦)", subtotal_estimate)],
    )


def rates_from_contract_meta(meta: ContractMeta) -> dict[str, float]:
    """contract_meta 의 rates 4종을 dict 로 추출. 없으면 KeyError."""
    if meta.rates is None:
        raise ValueError("contract_meta.rates 누락")
    r = meta.rates

    def _need(label, sourced):
        if sourced is None:
            raise ValueError(f"contract_meta.rates.{label} 누락")
        return sourced.value

    return {
        "general_admin": _need("general_admin_percent", r.general_admin_percent),
        "profit": _need("profit_percent", r.profit_percent),
        "industrial_accident_insurance": _need("industrial_accident_insurance_percent", r.industrial_accident_insurance_percent),
        "employment_insurance": _need("employment_insurance_percent", r.employment_insurance_percent),
    }
