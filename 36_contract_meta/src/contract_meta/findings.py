"""findings.md 생성기 — 검증 실패·경고만 분리한 검토 보고용 리포트.

extraction_report.md 가 전체 결과를 담는다면, findings.md 는 의뢰처에게 보낼 1차 검토서 초안.
"""

from __future__ import annotations

from datetime import datetime

from contract_meta.models import ContractMeta
from contract_meta.validators.consistency import ValidationReport


def build_findings(meta: ContractMeta, val: ValidationReport) -> str | None:
    """검증 실패·경고가 하나라도 있으면 findings.md 본문을 반환. 없으면 None."""
    if not val.failed and not val.warnings:
        return None

    lines: list[str] = []
    lines.append(f"# 검토 의견서 — {meta.project.name.value}")
    lines.append("")
    lines.append(f"- 작성일: {datetime.now().date().isoformat()}")
    lines.append(f"- 계약상대자: {meta.contractor.name_legal.value}")
    lines.append(f"- 발주자: {meta.owner.name.value}")
    lines.append(f"- 검토 자료: {len(meta.extraction.input_files)}건")
    lines.append("")
    lines.append(
        "본 검토 의견서는 제출된 계약 자료의 정합성 자동 검증 결과 중 "
        "확인이 필요한 항목만을 정리한 1차 검토서 초안입니다."
    )
    lines.append("")

    if val.failed:
        lines.append("## 1. 정합성 불일치 — 확인 필요")
        lines.append("")
        for i, r in enumerate(val.failed, start=1):
            lines.append(f"### 1.{i} {r.name}")
            lines.append("")
            lines.append(f"- 검토 결과: {r.detail}")
            lines.append("- 권고: 원본 자료(계약서·산출내역서)와 표기 일치 여부 재확인 요망.")
            lines.append("")

    if val.warnings:
        lines.append("## 2. 참고 사항")
        lines.append("")
        for i, r in enumerate(val.warnings, start=1):
            lines.append(f"### 2.{i} {r.name}")
            lines.append("")
            lines.append(f"- {r.detail}")
            lines.append("")

    lines.append("---")
    lines.append("")
    lines.append("문의: 사단법인 관계기관(건설관리) 분쟁지원팀")
    return "\n".join(lines)
