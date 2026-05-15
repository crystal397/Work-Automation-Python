"""cost_input.yaml + contract_meta.json → CostResult 빌더.

cost_input.yaml 의 manual 부분(노무비 합계, 경비 비목별 합계, 부가세)을 받고,
contract_meta.rates 의 4종 요율로 자동 산식 계산해 집계표 완성.
"""

from __future__ import annotations

import json
from dataclasses import dataclass
from datetime import date
from pathlib import Path

import yaml

from contract_meta.audit import make_source
from contract_meta.models import ContractMeta, Sourced

from cost_aggregation import calc
from cost_aggregation.audit import make_xlsx_source
from cost_aggregation.models import (
    Aggregate,
    CompanyCost,
    ComputedValue,
    CostResult,
    DirectExpense,
    ExpenseItem,
    RateBasedExpense,
    SalaryEntry,
)


def build_cost_result(input_yaml_path: Path, contract_meta_path: Path) -> CostResult:
    data = yaml.safe_load(input_yaml_path.read_text(encoding="utf-8"))
    meta_dict = json.loads(contract_meta_path.read_text(encoding="utf-8"))
    meta = ContractMeta.model_validate(meta_dict)

    rates = calc.rates_from_contract_meta(meta)
    prime = _build_company(data["prime"], "원도급사", rates)
    subs = [_build_company(s, "하도급사", rates) for s in data.get("subs", [])]

    grand_total_value = prime.grand_total_krw.value + sum(s.grand_total_krw.value for s in subs)
    grand_total = ComputedValue(
        value=grand_total_value,
        formula="원도급사 + 하도급사 총액 합산",
        inputs=[("원도급사", prime.grand_total_krw.value)] + [(f"하도급사_{i+1}", s.grand_total_krw.value) for i, s in enumerate(subs)],
    )

    # 인원별 합계 ↔ 명시 labor_total 정합성 검증 (warning)
    warnings: list[str] = []
    for company in [prime, *subs]:
        if not company.salaries:
            continue
        sal_sum = sum(e.total_krw.value for e in company.salaries)
        declared = company.indirect_labor_total_krw.value
        diff = abs(sal_sum - declared)
        # 1% 또는 1만원 초과 차이만 경고
        if diff > max(10000, declared * 0.01):
            warnings.append(
                f"[{company.company_role} {company.company_name.value}] 인원별 합계 "
                f"{sal_sum:,}원 ≠ 간접노무비 합계 {declared:,}원 (차이 {diff:,}원)"
            )

    return CostResult(
        contract_meta_ref=str(contract_meta_path),
        project_name=meta.project.name.value,
        period_actual=tuple(_to_date(d) for d in data["period_actual"]),
        period_estimate=tuple(_to_date(d) for d in data["period_estimate"]) if data.get("period_estimate") else None,
        aggregate=Aggregate(prime=prime, subs=subs, grand_total=grand_total),
        warnings=warnings,
    )


def _build_salary_entry(s: dict) -> SalaryEntry:
    """4.3.2.2 인원별 1행 자동 산정 — ③④⑥⑦ ComputedValue 자동.

    cost_input.yaml 의 salaries 항목 예::

        - name: 우동석
          role: 소장
          salary_actual_krw:    { value: 55850, _source: { ... } }    # ①
          work_days_actual:     { value: 1,     _source: { ... } }    # ②
          days_estimate:        { value: 95,    _source: { ... } }    # ⑤
    """
    salary_actual = s["salary_actual_krw"]["value"]
    work_days = s["work_days_actual"]["value"]
    days_est = s["days_estimate"]["value"]

    salary_actual_sourced = Sourced[int](
        value=salary_actual, _source=_src(s["salary_actual_krw"]["_source"]),
    )
    work_days_sourced = Sourced[int](
        value=work_days, _source=_src(s["work_days_actual"]["_source"]),
    )
    days_est_sourced = Sourced[int](
        value=days_est, _source=_src(s["days_estimate"]["_source"]),
    )

    severance = calc.calc_severance(salary_actual)                          # ③
    subtotal_actual = calc.calc_subtotal_actual(salary_actual, severance.value)  # ④
    daily_rate = calc.calc_daily_rate(subtotal_actual.value, work_days)     # ⑥
    estimate_sub = calc.calc_estimate_subtotal(days_est, daily_rate.value)  # ⑦
    total = calc.calc_total_salary(subtotal_actual.value, estimate_sub.value)  # C

    return SalaryEntry(
        name=s["name"],
        role=s["role"],
        salary_actual_krw=salary_actual_sourced,
        work_days_actual=work_days_sourced,
        severance_actual_krw=severance,
        subtotal_actual_krw=subtotal_actual,
        days_estimate=days_est_sourced,
        daily_rate_estimate=daily_rate,
        subtotal_estimate_krw=estimate_sub,
        total_krw=total,
    )


def _build_company(d: dict, company_role: str, rates: dict[str, float]) -> CompanyCost:
    # ── salaries[] 처리 (4.3.2.2 인원별 ⑤⑥⑦ 자동 산정) ─────────────────
    salaries: list[SalaryEntry] = []
    salaries_sum = 0
    for s in d.get("salaries", []):
        entry = _build_salary_entry(s)
        salaries.append(entry)
        salaries_sum += entry.total_krw.value

    # indirect_labor_total_krw: 명시 입력 vs salaries 합산 — 명시가 우선
    labor_decl = d.get("indirect_labor_total_krw")
    if labor_decl is not None:
        labor_total = labor_decl["value"]
        labor_src = _src(labor_decl["_source"])
        labor_total_cv = ComputedValue(
            value=labor_total,
            formula="간접노무비 합계 (외부 자료에서 집계)",
            inputs=[],
        )
    elif salaries:
        labor_total = salaries_sum
        labor_total_cv = ComputedValue(
            value=labor_total,
            formula="간접노무비 합계 = Σ 인원별 합계(C)",
            inputs=[(e.name, e.total_krw.value) for e in salaries],
        )
    else:
        raise ValueError(f"{company_role}: indirect_labor_total_krw 또는 salaries[] 둘 중 하나는 필수")

    direct_items: list[ExpenseItem] = []
    for it in d["direct_expense"]:
        actual = it["actual_krw"]["value"]
        estimate = it["estimate_krw"]["value"]
        total = ComputedValue(
            value=actual + estimate,
            formula="실비 + 추정",
            inputs=[("실비", actual), ("추정", estimate)],
        )
        direct_items.append(ExpenseItem(
            label=it["label"],
            actual_krw=Sourced[int](value=actual, _source=_src(it["actual_krw"]["_source"])),
            estimate_krw=Sourced[int](value=estimate, _source=_src(it["estimate_krw"]["_source"])),
            total_krw=total,
        ))
    direct_total = sum(it.total_krw.value for it in direct_items)
    direct_total_cv = ComputedValue(
        value=direct_total,
        formula="직접계상비목 합계",
        inputs=[(it.label, it.total_krw.value) for it in direct_items],
    )
    direct_expense = DirectExpense(items=direct_items, total=direct_total_cv)

    industrial_pct = rates["industrial_accident_insurance"] if company_role == "원도급사" else 0.0
    ind, emp = calc.calc_rate_based_insurance(
        labor_total,
        industrial_pct,
        rates["employment_insurance"],
    )
    rate_total = ComputedValue(
        value=ind.value + emp.value,
        formula="산재보험료 + 고용보험료",
        inputs=[("산재보험료", ind.value), ("고용보험료", emp.value)],
    )
    rate_based_expense = RateBasedExpense(
        industrial_accident_insurance=ind,
        employment_insurance=emp,
        total=rate_total,
    )

    expense_total = calc.calc_expense_total(direct_total, ind.value, emp.value)
    subtotal = calc.calc_subtotal(labor_total, expense_total.value)
    general_admin = calc.calc_general_admin(subtotal.value, rates["general_admin"])
    profit = calc.calc_profit(subtotal.value, general_admin.value, rates["profit"])
    gross = calc.calc_gross(subtotal.value, general_admin.value, profit.value)

    vat = d.get("vat_krw", {"value": 0, "_source": {"file": "", "method": "computed"}})
    vat_value = vat["value"]
    vat_sourced = Sourced[int](value=vat_value, _source=_src(vat["_source"]))

    grand_total = calc.calc_grand_total(gross.value, vat_value)

    return CompanyCost(
        company_role=company_role,
        company_name=Sourced[str](value=d["company_name"]["value"], _source=_src(d["company_name"]["_source"])),
        period_actual=tuple(_to_date(x) for x in d["period_actual"]),
        period_estimate=tuple(_to_date(x) for x in d["period_estimate"]) if d.get("period_estimate") else None,
        indirect_labor_total_krw=labor_total_cv,
        salaries=salaries,
        direct_expense=direct_expense,
        rate_based_expense=rate_based_expense,
        expense_total=expense_total,
        subtotal_krw=subtotal,
        general_admin_krw=general_admin,
        profit_krw=profit,
        gross_krw=gross,
        vat_krw=vat_sourced,
        grand_total_krw=grand_total,
    )


def _to_date(d) -> date:
    if isinstance(d, date):
        return d
    return date.fromisoformat(str(d))


def _src(d: dict):
    from contract_meta.models import Source
    return Source.model_validate(d)
