"""본문 템플릿 렌더러 — contract_meta JSON → 마크다운 본문.

Jinja2 환경. 템플릿은 contract_meta/body/<chapter>/<file>.md.j2 형태.
"""

from __future__ import annotations

import json
from datetime import date, datetime
from importlib import resources
from pathlib import Path

from jinja2 import Environment, FileSystemLoader, select_autoescape

from contract_meta.models import ContractMeta


# ── Jinja2 커스텀 필터 ────────────────────────────────────────────────

_DATE_FORMATS = ("%Y-%m-%d", "%Y.%m.%d", "%Y/%m/%d", "%Y%m%d")


def _parse_date(value) -> date | None:
    if value is None or value == "":
        return None
    if isinstance(value, datetime):
        return value.date()
    if isinstance(value, date):
        return value
    if isinstance(value, str):
        s = value.strip()
        for fmt in _DATE_FORMATS:
            try:
                return datetime.strptime(s, fmt).date()
            except ValueError:
                continue
    return None


def kdate(value, style: str = "dot") -> str:
    """한국 보고서 날짜 포맷 필터.

    스타일:
    - ``dot``     : ``2023. 12. 22.``  (KICM 표준)
    - ``long``    : ``2023년 12월 22일``
    - ``compact`` : ``2023.12.22``      (zero-pad)

    파싱 실패 시 원본 문자열 그대로 반환 (안전 폴백).

    사용 예 (Jinja2)::

        {{ contract_date | kdate }}            → 2023. 12. 22.
        {{ contract_date | kdate('long') }}    → 2023년 12월 22일
        {{ start | kdate }} ~ {{ end | kdate }}
    """
    d = _parse_date(value)
    if d is None:
        return "" if value is None else str(value)
    if style == "long":
        return f"{d.year}년 {d.month}월 {d.day}일"
    if style == "compact":
        return f"{d.year}.{d.month:02d}.{d.day:02d}"
    # 기본: dot (KICM 표준)
    return f"{d.year}. {d.month}. {d.day}."


def kperiod(start, end, sep: str = " ~ ", style: str = "dot") -> str:
    """기간 포맷 — ``2023. 12. 22. ~ 2024. 08. 31.``"""
    s = kdate(start, style)
    e = kdate(end, style)
    if not s and not e:
        return ""
    if not s:
        return f"{sep.lstrip()}{e}"
    if not e:
        return f"{s}{sep.rstrip()}"
    return f"{s}{sep}{e}"


def _register_filters(env: Environment) -> None:
    env.filters["kdate"] = kdate
    env.filters["kperiod"] = kperiod


def _to_context(meta: ContractMeta) -> dict:
    """ContractMeta → Jinja2 컨텍스트 dict (값만, _source 제거)."""
    d = meta.model_dump(by_alias=False)

    def strip(o):
        if isinstance(o, dict):
            if set(o.keys()) >= {"value", "src"}:
                return strip(o["value"])
            if set(o.keys()) >= {"value", "_source"}:
                return strip(o["value"])
            return {k: strip(v) for k, v in o.items() if not k.startswith("_") and k != "src"}
        if isinstance(o, list):
            return [strip(x) for x in o]
        return o

    return strip(d)


def render_chapter(meta: ContractMeta, chapter: str) -> str:
    """단일 챕터 렌더링.

    chapter 예: 'cover', 'chapter_2', 'chapter_3_2'
    """
    ctx = _to_context(meta)
    now = datetime.now()
    ctx["today_kr"] = f"{now.year}년 {now.month}월"

    template_root = resources.files("contract_meta.body")
    chapter_dir = template_root.joinpath(chapter)
    if not chapter_dir.is_dir():
        raise FileNotFoundError(f"챕터 폴더 없음: {chapter}")

    env = Environment(
        loader=FileSystemLoader(str(chapter_dir)),
        autoescape=select_autoescape(disabled_extensions=("j2",)),
        keep_trailing_newline=True,
    )
    _register_filters(env)

    chunks: list[str] = []
    for tpl_path in sorted(chapter_dir.iterdir()):
        if not tpl_path.name.endswith(".md.j2"):
            continue
        tpl = env.get_template(tpl_path.name)
        chunks.append(tpl.render(**ctx))

    return "\n\n".join(chunks)


def render_all(meta: ContractMeta, chapters: list[str]) -> str:
    """여러 챕터 연결."""
    parts = [render_chapter(meta, c) for c in chapters]
    return "\n\n---\n\n".join(parts)
