"""급여지불조서 (일반직 / 현채직) 월별 PDF → 부서 합계 자동 추출.

㈜A건설 C공구(철도) 양식 (텍스트 레이어 있음)
- 1쪽 PDF, 부서 단위 합계만 표시 (인원별 명세 X)
- 핵심 필드:
    부서코드(예: 230233), 부서명(C공구(철도) 현장)
    인원, 기본급, 연장수당, 소득총액, 공제 총액, 지급 총액
- 합계 행 ('총인원 :XX') 의 12개 컬럼 숫자가 우리가 원하는 부서 합계.

파일명에서 YYMM 추출 (예: 2412.pdf → 2024-12).
"""

from __future__ import annotations

import re
from calendar import monthrange
from dataclasses import dataclass
from datetime import date
from pathlib import Path

import pdfplumber

from contract_meta.models import Sourced
from cost_aggregation.audit import make_xlsx_source
from contract_meta.audit import make_source


_HEAD_FIELDS = [
    "기본급", "연장수당", "기타", "학자금_고", "전월_기타", "소득세",
    "노트북공제", "현지가불_원", "경조비", "정산주민세", "소득_총액",
]


@dataclass
class MonthlyPayrollTotal:
    year: int
    month: int
    file: str
    department_code: str | None         # 부서 코드 (예: 230233)
    department_name: str | None         # 부서명
    headcount: int                      # 인원
    gross_income_krw: Sourced[int]      # 소득 총액 (한 페이지의 마지막 큰 숫자)
    base_pay_krw: Sourced[int]          # 기본급
    overtime_krw: Sourced[int]          # 연장수당


FILENAME_YYMM = re.compile(r"(?<![\d])(\d{2})(\d{2})(?![\d])")


def _parse_yymm(name: str) -> tuple[int, int] | None:
    m = FILENAME_YYMM.search(name)
    if not m:
        return None
    yy, mm = int(m.group(1)), int(m.group(2))
    if 20 <= yy <= 30 and 1 <= mm <= 12:
        return (2000 + yy, mm)
    return None


def extract_monthly(pdf_path: str | Path) -> MonthlyPayrollTotal | None:
    """단일 월별 지불조서 PDF → MonthlyPayrollTotal."""
    pdf_path = Path(pdf_path)
    ym = _parse_yymm(pdf_path.name)
    if not ym:
        return None
    year, month = ym

    with pdfplumber.open(pdf_path) as pdf:
        page = pdf.pages[0]
        text = page.extract_text() or ""

    # 부서 코드/명
    dept_match = re.search(r"부\s*서\s*명\s*:\s*(\d+)\s+([^\n\r]+)", text)
    dept_code = dept_match.group(1) if dept_match else None
    dept_name = dept_match.group(2).strip() if dept_match else None

    # 인원 (예: '인원 :22' 또는 '총인원 :22')
    head_match = re.search(r"총?인\s*원\s*:\s*(\d+)", text)
    headcount = int(head_match.group(1)) if head_match else 0

    # 첫 데이터 행: '1 86,867,406 40,770,994 0 750,000 0 19,447,750 0 0 0 0 172,674,090'
    # = No(1) + 11개 숫자 토큰. 첫=기본급, 둘=연장수당, 마지막=소득총액.
    base_pay = overtime = gross_income = 0
    for line in text.splitlines():
        tokens = line.strip().split()
        if len(tokens) < 12:
            continue
        if not re.fullmatch(r"\d{1,3}", tokens[0]):
            continue
        try:
            data = [int(t.replace(",", "")) for t in tokens[1:12]]
        except ValueError:
            continue
        base_pay = data[0]
        overtime = data[1]
        gross_income = data[-1]
        break

    file_name = pdf_path.name
    return MonthlyPayrollTotal(
        year=year,
        month=month,
        file=str(pdf_path),
        department_code=dept_code,
        department_name=dept_name,
        headcount=headcount,
        gross_income_krw=Sourced[int](
            value=gross_income,
            _source=make_source(file=file_name, method="pdf_text", page=1,
                                 field_label="소득총액 (합계 행)",
                                 raw_text=f"인원 {headcount}, 부서 {dept_code}/{dept_name}"),
        ),
        base_pay_krw=Sourced[int](
            value=base_pay,
            _source=make_source(file=file_name, method="pdf_text", page=1, field_label="기본급 (합계 행)"),
        ),
        overtime_krw=Sourced[int](
            value=overtime,
            _source=make_source(file=file_name, method="pdf_text", page=1, field_label="연장수당 (합계 행)"),
        ),
    )


def extract_period_total(
    payslip_dir: str | Path,
    period_start: date,
    period_end: date,
    *,
    pattern: str = "*.pdf",
) -> tuple[int, list[MonthlyPayrollTotal], list[str]]:
    """디렉터리 내 월별 지불조서 PDF 일괄 처리 → 산정기간 ∩ 월 단위 합산.

    반환: (소득총액 합계 KRW, 월별 상세, warnings)
    """
    out: list[MonthlyPayrollTotal] = []
    warnings: list[str] = []

    for path in sorted(Path(payslip_dir).glob(pattern)):
        rec = extract_monthly(path)
        if rec is None:
            warnings.append(f"YYMM 미인식: {path.name}")
            continue
        m_first = date(rec.year, rec.month, 1)
        m_last = date(rec.year, rec.month, monthrange(rec.year, rec.month)[1])
        if m_last < period_start or m_first > period_end:
            continue
        out.append(rec)

    total = sum(m.gross_income_krw.value for m in out)
    return total, out, warnings
