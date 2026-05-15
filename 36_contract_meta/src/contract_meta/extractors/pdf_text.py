"""PDF 텍스트·표 추출 공통 모듈.

pdfplumber 를 1순위로 사용하고, 텍스트 레이어가 비어 있으면 PyMuPDF 로 폴백.
OCR 은 별도(나중에)에서 처리.
"""

from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path

import pdfplumber


@dataclass
class PdfPage:
    page_no: int          # 1-base
    text: str
    tables: list[list[list[str | None]]]


def extract_pages(pdf_path: str | Path) -> list[PdfPage]:
    pages: list[PdfPage] = []
    with pdfplumber.open(pdf_path) as pdf:
        for idx, page in enumerate(pdf.pages, start=1):
            text = page.extract_text() or ""
            tables = page.extract_tables() or []
            pages.append(PdfPage(page_no=idx, text=text, tables=tables))
    return pages
