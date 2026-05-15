"""보고서 5장 (첨부) — PDF 분류·합철.

입력 yaml 형식
---
sections:
  - title: "5.1. 공기연장 간접비 산정근거"
    files:
      - path: "samples/산정근거1.pdf"
        page_range: [1, 10]   # optional, 생략 시 전체
      - path: "samples/산정근거2.pdf"
  - title: "5.2. 계약문서"
    files: [...]
---

출력
- 5_appendix.pdf — 섹션 표지 페이지 + 합쳐진 PDF
- appendix_toc.md — 섹션별 시작 페이지 번호 TOC
"""

from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path

import fitz
import yaml


@dataclass
class _SectionResult:
    title: str
    start_page: int   # 1-base
    end_page: int


def build_appendix(
    spec_yaml_path: Path,
    out_pdf: Path,
    *,
    cover_page_dpi: int = 200,
) -> tuple[Path, list[_SectionResult]]:
    """yaml 명세 → 합쳐진 PDF + 섹션 TOC."""
    spec = yaml.safe_load(spec_yaml_path.read_text(encoding="utf-8"))
    out_pdf.parent.mkdir(parents=True, exist_ok=True)

    result_doc = fitz.open()
    sections_meta: list[_SectionResult] = []

    skipped: list[tuple[str, str]] = []  # (path, reason)

    for sec in spec.get("sections", []):
        title = sec.get("title", "")
        section_start = result_doc.page_count + 1
        _insert_cover_page(result_doc, title)
        for f in sec.get("files", []):
            path = Path(f["path"])
            if not path.exists():
                skipped.append((str(path), "파일 없음"))
                continue
            if path.suffix.lower() != ".pdf":
                # docx 등 비-PDF 는 합본 불가 — appendix.yaml 에는 등록되었지만 합본은 skip
                skipped.append((str(path), f"{path.suffix} 비-PDF 형식 (PDF 변환 후 재시도)"))
                continue
            try:
                src = fitz.open(str(path))
            except Exception as e:
                skipped.append((str(path), f"열기 실패: {e}"))
                continue
            try:
                page_range = f.get("page_range")
                if page_range:
                    start = max(1, int(page_range[0]))
                    end = min(src.page_count, int(page_range[1]))
                    result_doc.insert_pdf(src, from_page=start - 1, to_page=end - 1)
                else:
                    result_doc.insert_pdf(src)
            except Exception as e:
                skipped.append((str(path), f"insert_pdf 실패 (손상 가능): {e}"))
            finally:
                src.close()
        section_end = result_doc.page_count
        sections_meta.append(_SectionResult(title=title, start_page=section_start, end_page=section_end))

    if skipped:
        # 경고 파일 함께 출력 (사용자가 검토하도록)
        warn_path = out_pdf.with_name("appendix_skipped.txt")
        warn_path.write_text(
            "\n".join(f"{p}\t{r}" for p, r in skipped),
            encoding="utf-8",
        )

    result_doc.save(str(out_pdf))
    result_doc.close()
    return out_pdf, sections_meta


def _insert_cover_page(doc: fitz.Document, title: str) -> None:
    """A4 가로 한 페이지에 큰 제목만 가운데 배치."""
    page = doc.new_page(width=595, height=842)
    rect = fitz.Rect(50, 350, 545, 500)
    page.insert_textbox(
        rect,
        title,
        fontname="helv",
        fontsize=28,
        align=fitz.TEXT_ALIGN_CENTER,
    )


def write_toc(sections: list[_SectionResult], out_md: Path) -> None:
    lines = ["# 첨부 목차 (5장)", ""]
    lines.append("| 섹션 | 시작 페이지 | 끝 페이지 |")
    lines.append("|---|---:|---:|")
    for s in sections:
        lines.append(f"| {s.title} | {s.start_page} | {s.end_page} |")
    out_md.write_text("\n".join(lines), encoding="utf-8")
