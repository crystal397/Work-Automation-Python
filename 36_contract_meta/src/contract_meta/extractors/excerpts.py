"""ContractMeta 트리의 모든 (file, page) 출처를 순회해 PDF 1쪽 캡쳐를 자동 생성.

매칭 우선순위
1. _source.file == input_files[].name
2. _source.file == basename(input_files[].path)
"""

from __future__ import annotations

from pathlib import Path
from typing import Iterable

import fitz

from contract_meta.models import ContractMeta, Source


def emit_excerpts(meta: ContractMeta, out_dir: Path, *, dpi: int = 150) -> tuple[list[Path], list[str]]:
    """모든 (file, page) 페어를 1회 캡쳐. (captured_paths, warnings) 튜플 반환."""
    name_to_path: dict[str, str] = {}
    for f in meta.extraction.input_files:
        if f.name:
            name_to_path[f.name] = f.path
        name_to_path.setdefault(Path(f.path).name, f.path)

    pairs: set[tuple[str, int]] = set()
    for src in _walk_sources(meta):
        if src.file and src.page is not None and src.method not in ("computed", "cross_check"):
            pairs.add((src.file, src.page))

    out_dir.mkdir(parents=True, exist_ok=True)
    captured: list[Path] = []
    warnings: list[str] = []

    for file_name, page_no in sorted(pairs):
        if file_name not in name_to_path:
            warnings.append(f"매칭되는 input_files 없음: '{file_name}' (page {page_no})")
            continue
        pdf_path = name_to_path[file_name]
        if not Path(pdf_path).exists():
            warnings.append(f"파일 없음: {pdf_path}")
            continue
        if not pdf_path.lower().endswith(".pdf"):
            continue
        try:
            doc = fitz.open(pdf_path)
            if page_no < 1 or page_no > len(doc):
                warnings.append(f"페이지 범위 초과: {pdf_path} p.{page_no} (total {len(doc)})")
                continue
            pix = doc[page_no - 1].get_pixmap(dpi=dpi)
            out_path = out_dir / f"{Path(pdf_path).stem}_p{page_no:03d}.png"
            pix.save(str(out_path))
            captured.append(out_path)
        except Exception as e:
            warnings.append(f"캡쳐 실패 {pdf_path} p.{page_no}: {e}")

    return captured, warnings


def _walk_sources(obj) -> Iterable[Source]:
    if isinstance(obj, Source):
        yield obj
        return
    if hasattr(obj, "model_fields"):
        for name in obj.model_fields:
            yield from _walk_sources(getattr(obj, name))
    elif isinstance(obj, (list, tuple)):
        for v in obj:
            yield from _walk_sources(v)
    elif isinstance(obj, dict):
        for v in obj.values():
            yield from _walk_sources(v)
