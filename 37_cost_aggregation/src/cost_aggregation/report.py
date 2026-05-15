"""cost_report.md 생성기 — 4.3.1 / 4.4.1 / 4.5 표 마크다운 출력."""

from __future__ import annotations

from cost_aggregation.models import CompanyCost, CostResult


def _fmt(n: int) -> str:
    return f"{n:,}"


def build_cost_report(result: CostResult) -> str:
    lines: list[str] = []
    lines.append(f"# 공기연장 간접비 산정 결과 — {result.project_name}")
    lines.append("")
    lines.append(f"- 실비 구간: {result.period_actual[0]} ~ {result.period_actual[1]}")
    if result.period_estimate:
        lines.append(f"- 추정 구간: {result.period_estimate[0]} ~ {result.period_estimate[1]}")
    lines.append("")

    for c in [result.aggregate.prime, *result.aggregate.subs]:
        lines.append(f"## {c.company_role} — {c.company_name.value}")
        lines.append("")
        lines.append(_render_company_table(c))
        lines.append("")

    lines.append("## 4.5 결론 — 합계")
    lines.append("")
    lines.append("| 구분 | 금액(원) | 비고 |")
    lines.append("|---|---:|---|")
    lines.append(f"| 원도급사 | {_fmt(result.aggregate.prime.grand_total_krw.value)} ||")
    for i, s in enumerate(result.aggregate.subs, start=1):
        lines.append(f"| 하도급사{i} ({s.company_name.value}) | {_fmt(s.grand_total_krw.value)} ||")
    lines.append(f"| **합계** | **{_fmt(result.aggregate.grand_total.value)}** | 천원단위절사 |")
    lines.append("")

    return "\n".join(lines)


def _render_company_table(c: CompanyCost) -> str:
    rows: list[str] = []
    rows.append("| 구분 | 금액(원) | 산식 |")
    rows.append("|---|---:|---|")
    rows.append(f"| 1. 간접노무비 | {_fmt(c.indirect_labor_total_krw.value)} | 외부 집계 |")
    rows.append(f"| 2. 경비 | {_fmt(c.expense_total.value)} | {c.expense_total.formula} |")
    rows.append(f"| 　가. 직접계상비목 | {_fmt(c.direct_expense.total.value)} | 비목 합산 |")
    for it in c.direct_expense.items:
        rows.append(f"| 　　　{it.label} | {_fmt(it.total_krw.value)} | 실비 {_fmt(it.actual_krw.value)} + 추정 {_fmt(it.estimate_krw.value)} |")
    rows.append(f"| 　나. 승률계상비목 | {_fmt(c.rate_based_expense.total.value)} | 노무비 × 요율 |")
    rows.append(f"| 　　　산재보험료 | {_fmt(c.rate_based_expense.industrial_accident_insurance.value)} | {c.rate_based_expense.industrial_accident_insurance.formula} |")
    rows.append(f"| 　　　고용보험료 | {_fmt(c.rate_based_expense.employment_insurance.value)} | {c.rate_based_expense.employment_insurance.formula} |")
    rows.append(f"| 3. 소계 | {_fmt(c.subtotal_krw.value)} | {c.subtotal_krw.formula} |")
    rows.append(f"| 4. 일반관리비 | {_fmt(c.general_admin_krw.value)} | {c.general_admin_krw.formula} |")
    rows.append(f"| 5. 이윤 | {_fmt(c.profit_krw.value)} | {c.profit_krw.formula} |")
    rows.append(f"| 6. 총원가 | {_fmt(c.gross_krw.value)} | {c.gross_krw.formula} |")
    rows.append(f"| 7. 부가가치세 | {_fmt(c.vat_krw.value)} | 영세율 |")
    rows.append(f"| 8. 총액 | **{_fmt(c.grand_total_krw.value)}** | {c.grand_total_krw.formula} |")
    return "\n".join(rows)
