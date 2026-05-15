"""4.3 / 4.4 / 4.5 산정 결과 스키마.

36_contract_meta 의 `Sourced[T]` 규약을 그대로 재사용해 출처 메타를 강제한다.
모든 산식 결과 셀에는 `formula` (예: '(노무비+경비)×4.50%') 와 `inputs` (입력 셀의 (label, value) 리스트) 가 부착된다.
"""

from __future__ import annotations

from datetime import date
from typing import Literal

from contract_meta.models import Sourced
from pydantic import BaseModel, ConfigDict, Field


# ─────────────────────────────────────────────────────────────────
# 산식 audit
# ─────────────────────────────────────────────────────────────────

class ComputedValue(BaseModel):
    """산식 결과 셀. 입력 셀 + 산식 + 결과를 함께 보관."""
    value: int                                  # 정수 원 단위
    formula: str = Field(description="예: '(간접노무비+경비)×4.50%'")
    inputs: list[tuple[str, int]] = Field(default_factory=list, description="[(label, value)] — 산식의 직접 입력값")
    rate_percent: float | None = None


# ─────────────────────────────────────────────────────────────────
# 4.3.2 간접노무비
# ─────────────────────────────────────────────────────────────────

class Personnel(BaseModel):
    """간접노무비 산정 대상 인원 1명."""
    affiliation: Sourced[str]                   # ㈜A건설 / ㈜D
    name: Sourced[str]
    role: Sourced[str]                          # 소장 / 토목 / 안전 / 품질 / 미화 / 서무 / 취사
    period_start: Sourced[date]
    period_end: Sourced[date]


class SalaryEntry(BaseModel):
    """대상인원 1명의 급여 합산 결과 (4.3.2.2 표 한 행)."""
    name: str
    role: str
    salary_actual_krw: Sourced[int]             # ① 급여 (실비)
    work_days_actual: Sourced[int]              # ② 일수
    severance_actual_krw: ComputedValue         # ③ 퇴직급여충당금 = 급여 / 12
    subtotal_actual_krw: ComputedValue          # ④ 소계 = ① + ③
    days_estimate: Sourced[int]                 # ⑤ 추정 일수
    daily_rate_estimate: ComputedValue          # ⑥ 1일평균노무비 = ④ / ②
    subtotal_estimate_krw: ComputedValue        # ⑦ 추정 소계 = ⑤ × ⑥
    total_krw: ComputedValue                    # C 합계 = ④ + ⑦


# ─────────────────────────────────────────────────────────────────
# 4.3.3 경비
# ─────────────────────────────────────────────────────────────────

class ExpenseItem(BaseModel):
    """경비 비목 1줄 (전력수도광열비 / 여비교통통신비 / ...)."""
    label: str
    actual_krw: Sourced[int]
    estimate_krw: Sourced[int]
    total_krw: ComputedValue                    # = actual + estimate


class DirectExpense(BaseModel):
    """직접계상비목 묶음 (보고서 4.3.3.1)."""
    items: list[ExpenseItem]
    total: ComputedValue


class RateBasedExpense(BaseModel):
    """승률계상비목 (보고서 4.3.3.2). 산재·고용 보험료. 노무비 × 요율."""
    industrial_accident_insurance: ComputedValue    # 노무비 × industrial_accident_percent
    employment_insurance: ComputedValue              # 노무비 × employment_insurance_percent
    total: ComputedValue


# ─────────────────────────────────────────────────────────────────
# 4.3 / 4.4 한 회사(원도급 또는 하도급)의 집계표 한 묶음
# ─────────────────────────────────────────────────────────────────

class CompanyCost(BaseModel):
    """한 회사의 4.3 집계 한 묶음."""
    company_role: Literal["원도급사", "하도급사"]
    company_name: Sourced[str]
    period_actual: tuple[date, date]                 # 실비 산정 구간
    period_estimate: tuple[date, date] | None        # 추정 구간 (없으면 None)

    indirect_labor_total_krw: ComputedValue          # 1. 간접노무비 합계
    salaries: list[SalaryEntry]                      # 33행

    direct_expense: DirectExpense                    # 2-가
    rate_based_expense: RateBasedExpense             # 2-나
    expense_total: ComputedValue                     # 2. 경비 합계

    subtotal_krw: ComputedValue                      # 3. 소계 = 1 + 2
    general_admin_krw: ComputedValue                 # 4. 일반관리비 = 3 × general_admin_percent
    profit_krw: ComputedValue                        # 5. 이윤 = (3+4) × profit_percent
    gross_krw: ComputedValue                         # 6. 총원가 = 3+4+5
    vat_krw: Sourced[int]                            # 7. 부가가치세 (영세율이면 0)
    grand_total_krw: ComputedValue                   # 8. 총액 (천원단위절사)


# ─────────────────────────────────────────────────────────────────
# 4.5 결론 (원도급사 + 하도급사 합계)
# ─────────────────────────────────────────────────────────────────

class Aggregate(BaseModel):
    """4.5 결론 — 원도급사 + 하도급사 합계."""
    prime: CompanyCost
    subs: list[CompanyCost] = Field(default_factory=list)
    grand_total: ComputedValue


# ─────────────────────────────────────────────────────────────────
# 최상위
# ─────────────────────────────────────────────────────────────────

class CostResult(BaseModel):
    """37_cost_aggregation 최상위 산정 결과."""
    schema_version: Literal["0.1.0"] = "0.1.0"
    contract_meta_ref: str = Field(description="입력으로 사용한 contract_meta.json 의 상대경로")
    project_name: str
    period_actual: tuple[date, date]
    period_estimate: tuple[date, date] | None

    aggregate: Aggregate

    warnings: list[str] = Field(default_factory=list)

    model_config = ConfigDict(populate_by_name=True)
