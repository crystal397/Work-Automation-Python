"""공기연장 간접비 보고서 PDF에서 산정 요율(rates) 4종을 자동 추출.

근거 위치
- 산정결과표(보고서 본문 p.4 또는 4.3 집계표) 비고란의 산식: "1×3.70%", "1×1.57%", "3×4.50%", "(3+4)×5.18%"
- 4.3.4 일반관리비/이윤 단락의 "초과하지 않는 N.NN%임을 확인"

두 위치를 모두 확인해서 일치하면 신뢰도 ↑, 불일치하면 warning 반환.
"""

from __future__ import annotations

import re
from dataclasses import dataclass
from pathlib import Path

from contract_meta.audit import make_source
from contract_meta.extractors.pdf_text import extract_pages
from contract_meta.models import Source


_PATTERNS = {
    "industrial_accident_insurance_percent": re.compile(r"산재보험료[^\n]+1[×x*]\s*(\d+(?:\.\d+)?)\s*%"),
    "employment_insurance_percent":          re.compile(r"고용보험료[^\n]+1[×x*]\s*(\d+(?:\.\d+)?)\s*%"),
    "general_admin_percent":                 re.compile(r"일반관리비[^\n]+3[×x*]\s*(\d+(?:\.\d+)?)\s*%"),
    "profit_percent":                        re.compile(r"이윤[^\n]+\(3\+4\)\s*[×x*]\s*(\d+(?:\.\d+)?)\s*%"),
}

# 4.3.4 본문 단락 검증 패턴
_TEXT_VERIFY = {
    "general_admin_percent": re.compile(r"일반관리비율이?\s*[\d.]+\s*%를?\s*초과하지\s*않는\s*(\d+(?:\.\d+)?)\s*%"),
    "profit_percent":        re.compile(r"이윤율이?\s*[\d.]+\s*%를?\s*초과하지\s*않는\s*(\d+(?:\.\d+)?)\s*%"),
    "industrial_accident_insurance_percent": re.compile(r"산재보험료는?\s*(\d+(?:\.\d+)?)\s*%"),
    "employment_insurance_percent":          re.compile(r"고용보험료는?\s*(\d+(?:\.\d+)?)\s*%"),
}


@dataclass
class ExtractedRate:
    field: str
    value: float
    source: Source
    verified_against_body: bool


def extract_rates(pdf_path: str | Path) -> tuple[list[ExtractedRate], list[str]]:
    """보고서 PDF에서 rates 4종을 추출한다. (rates, warnings) 튜플."""
    pages = extract_pages(pdf_path)
    file_name = Path(pdf_path).name

    rates: list[ExtractedRate] = []
    warnings: list[str] = []

    table_hits: dict[str, tuple[float, int, str]] = {}
    body_hits: dict[str, tuple[float, int, str]] = {}

    for p in pages:
        for field, pat in _PATTERNS.items():
            if field in table_hits:
                continue
            m = pat.search(p.text)
            if m:
                table_hits[field] = (float(m.group(1)), p.page_no, m.group(0))
        for field, pat in _TEXT_VERIFY.items():
            if field in body_hits:
                continue
            m = pat.search(p.text)
            if m:
                body_hits[field] = (float(m.group(1)), p.page_no, m.group(0))

    for field in _PATTERNS:
        if field not in table_hits:
            warnings.append(f"산정결과표에서 {field} 미발견")
            continue
        value, page, raw = table_hits[field]
        verified = False
        if field in body_hits:
            body_value = body_hits[field][0]
            if abs(body_value - value) < 1e-9:
                verified = True
            else:
                warnings.append(
                    f"{field}: 산정결과표 {value}% vs 본문 단락 {body_value}% 불일치"
                )
        src = make_source(
            file=file_name,
            method="pdf_text",
            page=page,
            field_label="4.3.1 산정결과표 비고",
            raw_text=raw,
        )
        rates.append(ExtractedRate(field=field, value=value, source=src, verified_against_body=verified))

    return rates, warnings
