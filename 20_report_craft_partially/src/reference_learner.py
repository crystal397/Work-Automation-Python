"""
reference 폴더의 보고서(docx/pdf)에서 '2장 귀책분석' 파트를 추출.
결과를 output/reference_patterns.md 로 저장 → claude.ai에 붙여넣어 패턴 학습에 사용.
"""

from __future__ import annotations

import re
from concurrent.futures import ThreadPoolExecutor, as_completed
from pathlib import Path
from typing import NamedTuple

from docx import Document
from tqdm import tqdm
from .text_extractor import extract_pdf


# ── 귀책분석 섹션 탐지 패턴 ────────────────────────────────────────────────────

# 띄어쓰기 무관하게 귀책분석 관련 헤딩 탐지
_SECTION_PATTERNS = [
    r"귀\s*책\s*분\s*석",      # 귀책분석 (띄어쓰기 무관)
    r"귀\s*책\s*사\s*유",      # 귀책사유
    r"공\s*기\s*지\s*연\s*귀\s*책",  # 공기지연 귀책
    r"지\s*연\s*귀\s*책",
    r"책\s*임\s*소\s*재",
]

_SECTION_RE = re.compile("|".join(_SECTION_PATTERNS))

# 장(章) 번호 추출
_CHAPTER_NUM_RE = re.compile(r"제?\s*(\d+)\s*장")


class ReferencePattern(NamedTuple):
    file_name: str
    raw_section: str      # 귀책분석 섹션 원문 (단락 텍스트)
    tables_text: str      # 해당 섹션 내 표 내용


def _is_section_start(text: str) -> bool:
    return bool(_SECTION_RE.search(text))


def _extract_chapter_num(text: str) -> int | None:
    """텍스트에서 장(章) 번호 추출. 예: '제3장' → 3"""
    m = _CHAPTER_NUM_RE.search(text)
    return int(m.group(1)) if m else None


def _is_chapter_after(start_num: int | None, text: str) -> bool:
    """start_num 보다 큰 장 번호가 나오면 True (섹션 종료 신호)."""
    m = _CHAPTER_NUM_RE.search(text)
    if m:
        num = int(m.group(1))
        if start_num is None:
            return num >= 3   # 시작 장 모를 경우 3장 이후부터 종료
        return num > start_num
    return False


def extract_accountability_section(docx_path: Path) -> ReferencePattern:
    """
    docx 파일에서 귀책분석 섹션을 추출.
    - 띄어쓰기 무관하게 '귀책분석/귀책사유' 패턴으로 모든 후보 위치를 찾는다.
    - 각 후보에서 다음 장(章) 시작 전까지 수집한 뒤, 내용이 가장 많은 후보를 선택.
      (목차 항목은 내용이 거의 없으므로 자동으로 탈락)
    - 장 번호를 인식해, 섹션이 3장이면 4장 이후에서만 종료한다.
    """
    try:
        doc = Document(str(docx_path))
    except Exception as e:
        return ReferencePattern(docx_path.name, f"[읽기 실패: {e}]", "")

    from docx.oxml.ns import qn

    body = doc.element.body
    blocks: list[tuple[str, str]] = []
    for child in body:
        tag = child.tag.split("}")[-1]
        if tag == "p":
            text = "".join(run.text for run in child.iter(qn("w:t")))
            blocks.append(("para", text))
        elif tag == "tbl":
            rows = []
            for tr in child.iter(qn("w:tr")):
                cells = []
                for tc in tr.iter(qn("w:tc")):
                    cell_text = "".join(t.text for t in tc.iter(qn("w:t")))
                    cells.append(cell_text.strip())
                rows.append(" | ".join(cells))
            blocks.append(("table", "\n".join(rows)))

    # ── 모든 귀책분석 후보 위치 수집 ──────────────────────────────────────────
    candidate_starts: list[int] = [
        i for i, (kind, content) in enumerate(blocks)
        if kind == "para" and _is_section_start(content)
    ]

    best_blocks: list[tuple[str, str]] = []
    best_content_len = 0

    for start_idx in candidate_starts:
        start_chapter = _extract_chapter_num(blocks[start_idx][1])
        section: list[tuple[str, str]] = []

        for i in range(start_idx, len(blocks)):
            kind, content = blocks[i]
            # 시작 블록 이후에 상위 장이 나오면 종료
            if i > start_idx and kind == "para" and content.strip():
                if _is_chapter_after(start_chapter, content):
                    break
            section.append((kind, content))

        # 헤더 제외 실질 내용 길이로 비교 (목차 항목은 자동 탈락)
        content_len = sum(len(c) for k, c in section[1:] if c.strip())
        if content_len > best_content_len:
            best_content_len = content_len
            best_blocks = section

    if not best_blocks:
        full_text = "\n".join(c for _, c in blocks if c.strip())
        return ReferencePattern(
            docx_path.name,
            "[섹션 자동 탐지 실패 — 전체 본문 포함]\n" + full_text[:3000],
            ""
        )

    paras = [c for k, c in best_blocks if k == "para" and c.strip()]
    tables = [c for k, c in best_blocks if k == "table" and c.strip()]

    return ReferencePattern(
        docx_path.name,
        "\n".join(paras),
        "\n\n[표]\n".join(tables) if tables else ""
    )



def extract_accountability_section_pdf(pdf_path: Path) -> ReferencePattern:
    """
    PDF 파일에서 귀책분석 섹션을 추출.
    text_extractor.extract_pdf 로 전체 텍스트를 뽑은 뒤, 귀책분석 관련 구간을 잘라낸다.
    """
    result = extract_pdf(pdf_path)
    if not result.text.strip():
        return ReferencePattern(pdf_path.name, f"[읽기 실패: {result.error or '텍스트 없음'}]", "")

    full_text = result.text
    match = _SECTION_RE.search(full_text)
    if not match:
        return ReferencePattern(
            pdf_path.name,
            "[귀책분석 섹션 자동 탐지 실패 — 전체 본문 일부 포함]\n" + full_text[:5000],
            ""
        )

    section_text = full_text[match.start():]

    start_num = None
    start_match = _CHAPTER_NUM_RE.search(full_text[:match.start()])
    if start_match:
        start_num = int(start_match.group(1))

    chapter_end = _CHAPTER_NUM_RE.search(section_text[50:])
    if chapter_end:
        chapter_num = int(chapter_end.group(1))
        if start_num is None or chapter_num > start_num:
            section_text = section_text[:50 + chapter_end.start()]

    return ReferencePattern(pdf_path.name, section_text[:8000], "")

def learn_all(reference_dir: Path, output_dir: Path) -> Path:
    """
    모든 reference docx에서 귀책분석 섹션을 추출하고
    output/reference_patterns.md 에 저장.
    """
    output_dir.mkdir(parents=True, exist_ok=True)
    out_path = output_dir / "reference_patterns.md"

    docx_files = sorted(reference_dir.glob("*.docx")) + sorted(reference_dir.glob("*.DOCX"))
    pdf_files = sorted(reference_dir.glob("*.pdf")) + sorted(reference_dir.glob("*.PDF"))
    # 중복 제거 (동일 파일이 두 glob에 걸리는 경우 — Windows는 대소문자 무관하나 Linux 대비)
    seen: set[str] = set()
    unique_files: list[Path] = []
    for f in docx_files + pdf_files:
        key = str(f).lower()
        if key not in seen:
            seen.add(key)
            unique_files.append(f)
    all_files = sorted(unique_files)

    if not all_files:
        print(f"[오류] reference 폴더에 docx/pdf 파일이 없습니다: {reference_dir}")
        return out_path

    print(f"reference 보고서 {len(all_files)}개 처리 중 (docx: {len(docx_files)}, pdf: {len(pdf_files)})...")

    patterns: list[ReferencePattern] = []
    failed: list[str] = []

    def _process(f: Path) -> ReferencePattern:
        if f.suffix.lower() == ".pdf":
            return extract_accountability_section_pdf(f)
        return extract_accountability_section(f)

    with tqdm(total=len(all_files), desc="reference 학습", unit="파일") as bar:
        with ThreadPoolExecutor() as ex:
            futures = {ex.submit(_process, f): f for f in all_files}
            for fut in as_completed(futures):
                f = futures[fut]
                bar.set_postfix(파일=f.name[:30])
                p = fut.result()
                if "[읽기 실패" in p.raw_section:
                    failed.append(f.name)
                    tqdm.write(f"  [실패] {f.name}: {p.raw_section}")
                else:
                    patterns.append(p)
                    has_table = bool(p.tables_text.strip())
                    tqdm.write(f"  [완료] {f.name} — 단락 {len(p.raw_section.splitlines())}줄, 표 {'있음' if has_table else '없음'}")
                bar.update(1)

    # 병렬 처리로 인한 순서 불일치 보정 — 파일명 기준 재정렬
    patterns.sort(key=lambda p: p.file_name)

    # ── 마크다운 저장 ──────────────────────────────────────────────────────────
    lines = [
        "# REFERENCE 보고서 — 귀책분석 섹션 추출 결과",
        "",
        "이 파일을 claude.ai에 붙여넣어 패턴 학습을 요청하세요.",
        "",
        f"- 처리 완료: {len(patterns)}개",
        f"- 읽기 실패: {len(failed)}개" + (f" ({', '.join(failed)})" if failed else ""),
        "",
        "---",
        "",
    ]

    for p in patterns:
        lines += [
            f"## {p.file_name}",
            "",
            "### 단락 텍스트",
            "",
            p.raw_section if p.raw_section.strip() else "_없음_",
            "",
        ]
        if p.tables_text.strip():
            lines += [
                "### 표 내용",
                "",
                "```",
                p.tables_text,
                "```",
                "",
            ]
        lines += ["---", ""]

    if failed:
        lines += ["## [읽기 실패 목록]", ""]
        for f in failed:
            lines.append(f"- {f}")
        lines.append("")

    out_path.write_text("\n".join(lines), encoding="utf-8")
    print(f"\n저장 완료: {out_path}")
    return out_path
