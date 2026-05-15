"""출처(Source) 메타 빌더."""

from __future__ import annotations

import hashlib
from pathlib import Path
from typing import Literal

from contract_meta.models import Source, Sourced

Method = Literal["pdf_text", "ocr", "llm", "manual", "computed", "cross_check"]


def make_source(
    file: str | Path,
    method: Method,
    *,
    page: int | None = None,
    sheet: str | None = None,
    cell: str | None = None,
    field_label: str | None = None,
    raw_text: str | None = None,
) -> Source:
    return Source(
        file=str(file),
        page=page,
        sheet=sheet,
        cell=cell,
        field_label=field_label,
        raw_text=raw_text,
        method=method,
    )


def sourced[T](value: T, source: Source) -> Sourced[T]:
    """값 + 출처를 한 줄로 묶는 헬퍼."""
    return Sourced[type(value) if value is not None else object](value=value, _source=source)


def sha256_of(path: str | Path) -> str:
    h = hashlib.sha256()
    with open(path, "rb") as f:
        for chunk in iter(lambda: f.read(1 << 20), b""):
            h.update(chunk)
    return h.hexdigest()
