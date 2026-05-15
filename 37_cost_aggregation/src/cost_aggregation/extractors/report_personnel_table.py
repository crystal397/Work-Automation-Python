"""보고서 본문 4.3.2.2 / 4.4.2.2 표 자동 추출 (cross_check / 검증용).

보고서 자체에서 33명/21명 인원별 ① 급여 + ② 일수 + ③④⑤⑥⑦ 산식 결과를 자동 추출.
보고서 재생성·검증 시 0원 일치 보장. 새 프로젝트는 다른 자료에서 ①을 산정.

표 구조 (보고서 p.61~65 / p.71~72)
- 헤더: 'A. 실비 ... B. 추정 ... C. 합계'
- 컬럼:
    No 소속 이름 직무 ① 급여 ② 일수 ③ 퇴직 ④ 소계 [⑤ 일수 ⑥ 1일평균 ⑦ 소계] C 합계
- 일부 인원은 ⑤⑥⑦ 비어있음 (퇴직 인원 등)
- 일부 인원은 ① 값 자체가 비어있음 (강상진 류 — 명세서 미제출, 일수만 표시)
- 마지막 행: '계 ... 합계'
"""

from __future__ import annotations

import re
from dataclasses import dataclass
from pathlib import Path

import pdfplumber

from contract_meta.audit import make_source
from contract_meta.models import Sourced


@dataclass
class PersonnelTableRow:
    seq: int
    affiliation: str
    name: Sourced[str]
    role: str
    salary_actual_krw: Sourced[int] | None     # ① 급여 (None 이면 표에서 비어있음)
    days_actual: Sourced[int] | None           # ②
    severance_actual_krw: int | None           # ③ = ① / 12
    subtotal_actual_krw: int | None            # ④ = ① + ③
    days_estimate: int | None                  # ⑤
    daily_rate_estimate: int | None            # ⑥
    subtotal_estimate_krw: int | None          # ⑦
    total_krw: int | None                      # C 합계


# No 소속 이름 직무 + 9~11 숫자 (천단위 쉼표 또는 정수)
_TOKEN_NUM = r"[\d,]+"


def extract_table(pdf_path: str | Path, page_range: tuple[int, int]) -> tuple[list[PersonnelTableRow], int | None]:
    """보고서 본문 p.start~end 의 인원별 표 추출. (rows, total_① 합계) 반환."""
    pdf_path = Path(pdf_path)
    file_name = pdf_path.name
    rows: list[PersonnelTableRow] = []
    total_salary: int | None = None

    with pdfplumber.open(pdf_path) as pdf:
        for pno in range(page_range[0], page_range[1] + 1):
            if pno < 1 or pno > len(pdf.pages):
                continue
            text = pdf.pages[pno - 1].extract_text() or ""
            for line in text.splitlines():
                row, ts = _parse_line(line, file_name, pno)
                if row is not None:
                    rows.append(row)
                elif ts is not None:
                    total_salary = ts
    return rows, total_salary


_ROW_PAT = re.compile(
    r"^\s*(?P<no>\d{1,3})\s+(?P<aff>\S+)\s+(?P<name>[가-힣]{2,5})\s+(?P<role>\S+)\s+(?P<rest>.+)$"
)
_TOTAL_PAT = re.compile(r"^\s*계\s+(.+)$")


def _parse_line(line: str, file_name: str, page: int) -> tuple[PersonnelTableRow | None, int | None]:
    m = _ROW_PAT.match(line)
    if m:
        rest = m.group("rest")
        nums = [n.replace(",", "") for n in re.findall(_TOKEN_NUM, rest)]
        if not nums:
            return None, None
        try:
            nums_int = [int(n) for n in nums]
        except ValueError:
            return None, None

        seq = int(m.group("no"))
        aff = m.group("aff")
        name = m.group("name")
        role = m.group("role")

        salary = days_a = severance = subtotal_a = None
        days_e = daily = subtotal_e = total = None

        if len(nums_int) >= 5:
            salary = nums_int[0]
            days_a = nums_int[1]
            severance = nums_int[2]
            subtotal_a = nums_int[3]
            if len(nums_int) >= 8:
                days_e = nums_int[4]
                daily = nums_int[5]
                subtotal_e = nums_int[6]
                total = nums_int[7]
            else:
                total = nums_int[-1]
        elif len(nums_int) == 2:
            days_a = nums_int[0]
            days_e = nums_int[1]

        row = PersonnelTableRow(
            seq=seq, affiliation=aff,
            name=Sourced[str](
                value=name,
                _source=make_source(file=file_name, method="pdf_text", page=page,
                                     field_label=f"4.x.2.2 표 {seq}행", raw_text=line.strip()),
            ),
            role=role,
            salary_actual_krw=Sourced[int](
                value=salary,
                _source=make_source(file=file_name, method="pdf_text", page=page,
                                     field_label=f"① 급여 (행 {seq})", raw_text=str(salary)),
            ) if salary is not None else None,
            days_actual=Sourced[int](
                value=days_a,
                _source=make_source(file=file_name, method="pdf_text", page=page,
                                     field_label=f"② 일수 (행 {seq})", raw_text=str(days_a)),
            ) if days_a is not None else None,
            severance_actual_krw=severance,
            subtotal_actual_krw=subtotal_a,
            days_estimate=days_e,
            daily_rate_estimate=daily,
            subtotal_estimate_krw=subtotal_e,
            total_krw=total,
        )
        return row, None

    tm = _TOTAL_PAT.match(line)
    if tm:
        nums = re.findall(_TOKEN_NUM, tm.group(1))
        if nums:
            ints = [int(n.replace(",", "")) for n in nums]
            if ints:
                return None, ints[0]
    return None, None
