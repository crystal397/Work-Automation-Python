"""extraction_report.md 생성."""

from __future__ import annotations

from datetime import datetime

from contract_meta.models import ContractMeta
from contract_meta.validators.consistency import ValidationReport


def build_report(meta: ContractMeta, val: ValidationReport) -> str:
    lines: list[str] = []
    lines.append("# 계약 메타데이터 추출 보고서")
    lines.append("")
    lines.append(f"- 생성: {datetime.now().isoformat(timespec='seconds')}")
    lines.append(f"- 사업: {meta.project.name.value}")
    lines.append(f"- 계약상대자: {meta.contractor.name_legal.value}")
    lines.append(f"- 발주자: {meta.owner.name.value}")
    lines.append("")

    # ── 검증 결과 ──
    lines.append("## ✅ 자동 검증 결과")
    lines.append("")
    if not val.failed and not val.warnings:
        lines.append(f"모든 검증 통과 ({len(val.passed)}건).")
        lines.append("")
    else:
        lines.append(f"- 통과: {len(val.passed)}건")
        lines.append(f"- 실패: {len(val.failed)}건")
        lines.append(f"- 경고: {len(val.warnings)}건")
        lines.append("")

    if val.failed:
        lines.append("### ❌ 실패")
        lines.append("")
        for r in val.failed:
            lines.append(f"- **{r.name}** — {r.detail}")
        lines.append("")

    if val.warnings:
        lines.append("### ⚠️ 경고")
        lines.append("")
        for r in val.warnings:
            lines.append(f"- **{r.name}** — {r.detail}")
        lines.append("")

    lines.append("### 🔍 통과한 검증")
    lines.append("")
    for r in val.passed:
        lines.append(f"- {r.name} — {r.detail}")
    lines.append("")

    # ── 출처 요약 ──
    lines.append("## 📂 출처 (입력 파일)")
    lines.append("")
    for f in meta.extraction.input_files:
        lines.append(f"- `{f.role}` :: {f.path} ({f.method}, sha256={f.sha256[:12]}…)")
    lines.append("")

    # ── 핵심 필드 ──
    lines.append("## 🎯 핵심 필드")
    lines.append("")
    lines.append("| 영역 | 필드 | 값 | 출처 |")
    lines.append("|---|---|---|---|")
    rows: list[tuple[str, str, str, str]] = []
    rows.append(("project", "name", meta.project.name.value, _src(meta.project.name.src)))
    rows.append(("project", "client_type", meta.project.client_type.value.value, _src(meta.project.client_type.src)))
    rows.append(("project", "contract_form", meta.project.contract_form.value.value, _src(meta.project.contract_form.src)))
    rows.append(("project", "bidding_method", meta.project.bidding_method.value.value, _src(meta.project.bidding_method.src)))
    rows.append(("owner", "name", meta.owner.name.value, _src(meta.owner.name.src)))
    rows.append(("contractor", "name_legal", meta.contractor.name_legal.value, _src(meta.contractor.name_legal.src)))
    rows.append(("total_contract", "initial.amount", f"{meta.total_contract.initial.amount.krw.value:,}원",
                 _src(meta.total_contract.initial.amount.krw.src)))
    if meta.total_contract.revisions:
        last = meta.total_contract.revisions[-1]
        rows.append(("total_contract", f"revisions[-1].amount({last.seq}회)",
                     f"{last.amount.krw.value:,}원", _src(last.amount.krw.src)))
    rows.append(("first_year_contract", "initial.period",
                 f"{meta.first_year_contract.initial.period_start.value} ~ {meta.first_year_contract.initial.period_end.value} ({meta.first_year_contract.initial.duration_days.value}일)",
                 _src(meta.first_year_contract.initial.period_start.src)))
    if meta.first_year_contract.revisions:
        last = meta.first_year_contract.revisions[-1]
        rows.append(("first_year_contract", f"revisions[-1].period({last.seq}회)",
                     f"{last.period_start.value} ~ {last.period_end.value} ({last.duration_days.value}일, +{last.duration_diff_days.value})",
                     _src(last.period_start.src)))
    if meta.calculation_target is not None:
        ct = meta.calculation_target
        rows.append(("calculation_target", "period",
                     f"{ct.period_start.value} ~ {ct.period_end.value} ({ct.days.value}일)",
                     _src(ct.period_start.src)))
    if meta.rates is not None:
        if meta.rates.general_admin_percent is not None:
            rows.append(("rates", "general_admin_percent",
                         f"{meta.rates.general_admin_percent.value}%",
                         _src(meta.rates.general_admin_percent.src)))
        if meta.rates.profit_percent is not None:
            rows.append(("rates", "profit_percent",
                         f"{meta.rates.profit_percent.value}%",
                         _src(meta.rates.profit_percent.src)))

    for area, field_name, value, src in rows:
        lines.append(f"| {area} | {field_name} | {value} | {src} |")
    lines.append("")

    return "\n".join(lines)


def _src(s) -> str:
    parts: list[str] = []
    if s.file:
        parts.append(s.file)
    if s.page is not None:
        parts.append(f"p.{s.page}")
    if s.sheet:
        parts.append(f"sheet={s.sheet}")
    if s.cell:
        parts.append(f"cell={s.cell}")
    if s.field_label:
        parts.append(f"§{s.field_label}")
    parts.append(f"[{s.method}]")
    return " ".join(parts) if parts else "(no source)"
