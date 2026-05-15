"""38 의 Source 빌더 — contract_meta.audit 재사용."""

from __future__ import annotations

from pathlib import Path

from contract_meta.audit import make_source
from contract_meta.models import Source


def make_xlsx_source(
    path: str | Path,
    sheet: str,
    cell: str,
    *,
    raw: str | None = None,
    label: str | None = None,
) -> Source:
    return make_source(
        file=Path(path).name,
        method="xlsx",
        sheet=sheet,
        cell=cell,
        field_label=label,
        raw_text=raw,
    )
