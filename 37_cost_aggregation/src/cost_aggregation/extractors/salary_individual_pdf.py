"""인원별 월별 급여명세서 PDF (현채직 양식) → 인원별 급여 자동 추출.

㈜A건설 현채직 양식
- 파일명: '<이름> <YYMM>.pdf' 또는 '<이름>_<YYMM>.pdf'
- 1쪽 PDF, 텍스트 레이어 있음
- 필드:
    사번 (S2400405), 성명, 직급, 소속, 근무일수, 지급일자,
    소득총액, 공제계, 수령액

산정기간 ∩ 인원별 활동 월의 소득총액·근무일수 합산.
"""

from __future__ import annotations

import re
from calendar import monthrange
from dataclasses import dataclass, field
from datetime import date
from pathlib import Path

import pdfplumber

from contract_meta.audit import make_source
from contract_meta.models import Sourced


FILENAME_PAT = re.compile(r"^(?P<name>[가-힣]+)[\s_]+(?P<yymm>\d{4})\.pdf$", re.IGNORECASE)


@dataclass
class MonthlyIndividualSlip:
    year: int
    month: int
    name: str
    employee_id: str | None
    role: str | None
    work_days: int
    gross_income_krw: int
    deduction_krw: int
    net_pay_krw: int
    source_file: str


@dataclass
class AggregatedIndividual:
    name: Sourced[str]
    employee_id: str | None
    role: Sourced[str] | None
    slips: list[MonthlyIndividualSlip] = field(default_factory=list)

    @property
    def total_gross_krw(self) -> int:
        return sum(s.gross_income_krw for s in self.slips)

    @property
    def total_work_days(self) -> int:
        return sum(s.work_days for s in self.slips)


def parse_filename(name: str) -> tuple[str, int, int] | None:
    m = FILENAME_PAT.match(name)
    if not m:
        return None
    person = m.group("name").strip()
    yymm = m.group("yymm")
    yy, mm = int(yymm[:2]), int(yymm[2:])
    if not (20 <= yy <= 30 and 1 <= mm <= 12):
        return None
    return (person, 2000 + yy, mm)


def extract_slip(pdf_path: str | Path) -> MonthlyIndividualSlip | None:
    pdf_path = Path(pdf_path)
    fp = parse_filename(pdf_path.name)
    if not fp:
        return None
    name, year, month = fp

    with pdfplumber.open(pdf_path) as pdf:
        text = pdf.pages[0].extract_text() or ""

    emp_id_m = re.search(r"\b(S\d{7})\b", text)
    employee_id = emp_id_m.group(1) if emp_id_m else None

    # 근무일수: "근무일수\n... \n31" 또는 한 줄에 인접
    work_days = _grab_int_after(text, r"근무일수", default=0, limit_lines=20)
    gross = _grab_money_after(text, r"소득총액", default=0)
    deduction = _grab_money_after(text, r"공제계", default=0)
    net = _grab_money_after(text, r"수령액", default=0)

    role = _grab_role(text)

    return MonthlyIndividualSlip(
        year=year, month=month, name=name,
        employee_id=employee_id, role=role,
        work_days=work_days,
        gross_income_krw=gross, deduction_krw=deduction, net_pay_krw=net,
        source_file=str(pdf_path),
    )


def _grab_int_after(text: str, label_pat: str, *, default: int, limit_lines: int) -> int:
    """label_pat 라벨 이후 limit_lines 내 첫 정수."""
    lines = text.splitlines()
    for i, line in enumerate(lines):
        if re.search(label_pat, line):
            for j in range(i + 1, min(i + 1 + limit_lines, len(lines))):
                m = re.search(r"\b(\d{1,3})\b", lines[j])
                if m:
                    return int(m.group(1))
            break
    return default


def _grab_money_after(text: str, label_pat: str, *, default: int) -> int:
    lines = text.splitlines()
    for i, line in enumerate(lines):
        if re.search(label_pat, line):
            for j in range(i, min(i + 5, len(lines))):
                m = re.search(r"([\d,]{5,})", lines[j])
                if m:
                    return int(m.group(1).replace(",", ""))
            break
    return default


def _grab_role(text: str) -> str | None:
    """'직급' 헤더 다음 라인 (공란이면 None)."""
    lines = text.splitlines()
    for i, line in enumerate(lines):
        if line.strip() == "직급":
            if i + 1 < len(lines):
                v = lines[i + 1].strip()
                if v and v not in ("성명", "소속"):
                    return v
    return None


def aggregate_individuals(
    dir_path: str | Path,
    period_start: date,
    period_end: date,
) -> tuple[list[AggregatedIndividual], list[str]]:
    """디렉터리 내 인원별 월별 PDF 일괄 처리 → 인원별 합계."""
    by_name: dict[str, AggregatedIndividual] = {}
    warnings: list[str] = []

    for pdf_path in sorted(Path(dir_path).glob("*.pdf")):
        slip = extract_slip(pdf_path)
        if slip is None:
            warnings.append(f"파싱 실패: {pdf_path.name}")
            continue
        m_first = date(slip.year, slip.month, 1)
        m_last = date(slip.year, slip.month, monthrange(slip.year, slip.month)[1])
        if m_last < period_start or m_first > period_end:
            continue

        if slip.name not in by_name:
            file_name = pdf_path.name
            by_name[slip.name] = AggregatedIndividual(
                name=Sourced[str](
                    value=slip.name,
                    _source=make_source(file=file_name, method="pdf_text", page=1,
                                         field_label="성명", raw_text=slip.name),
                ),
                employee_id=slip.employee_id,
                role=Sourced[str](
                    value=slip.role,
                    _source=make_source(file=file_name, method="pdf_text", page=1, field_label="직급"),
                ) if slip.role else None,
            )
        by_name[slip.name].slips.append(slip)

    return sorted(by_name.values(), key=lambda x: x.name.value), warnings
