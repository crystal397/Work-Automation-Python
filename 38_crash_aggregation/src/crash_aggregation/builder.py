"""월별 xlsx 시리즈 → CrashResult 빌더."""

from __future__ import annotations

import re
from calendar import monthrange
from datetime import date
from pathlib import Path

from crash_aggregation.extractors.crew_xlsx import extract_workers
from crash_aggregation.models import CategoryTotal, CrashResult, MonthlyCrash


FILENAME_YYMM_PAT = re.compile(r"(?<![\d.])(\d{2})\.(\d{2})(?![\d.])")


def parse_yymm_from_filename(path: Path) -> tuple[int, int] | None:
    """파일명에서 YY.MM 추출. 예: '보안해제_01. 24.06_노무(직접시공).xlsx' → (2024, 6).

    `YY.MM` 양옆이 숫자·점이 아닌 경계여야 매칭. `_01. 24.06_` 에서 `24.06` 만 잡힘.
    """
    for yy_s, mm_s in FILENAME_YYMM_PAT.findall(path.name):
        try:
            yy, mm = int(yy_s), int(mm_s)
        except ValueError:
            continue
        if 1 <= mm <= 12 and 20 <= yy <= 30:
            return (2000 + yy, mm)
    return None


def build_crash_result(
    source_files: list[Path],
    *,
    project_name: str,
    contract_meta_ref: str | None = None,
) -> CrashResult:
    months: list[MonthlyCrash] = []
    all_dates: list[date] = []

    for f in source_files:
        ym = parse_yymm_from_filename(f)
        if ym is None:
            continue
        year, month = ym
        workers, _ = extract_workers(f, year=year, month=month)

        cat_buckets: dict[str, dict] = {}
        for w in workers:
            for d in w.days:
                key = d.category.value
                slot = cat_buckets.setdefault(key, {"manday": 0.0, "krw": 0, "by_gongjong": {}})
                slot["manday"] += d.manday.value
                slot["krw"] += d.amount_krw.value
                g = w.gongjong.value
                slot["by_gongjong"][g] = slot["by_gongjong"].get(g, 0) + d.amount_krw.value

        cat_totals = [
            CategoryTotal(category=k, total_manday=v["manday"], total_krw=v["krw"], by_gongjong=v["by_gongjong"])
            for k, v in sorted(cat_buckets.items())
        ]
        months.append(MonthlyCrash(
            year=year, month=month, source_file=str(f),
            workers=workers, category_totals=cat_totals,
        ))
        all_dates.append(date(year, month, 1))
        all_dates.append(date(year, month, monthrange(year, month)[1]))

    months.sort(key=lambda m: (m.year, m.month))

    if not all_dates:
        raise ValueError("처리된 월이 0건 — 파일명에서 YY.MM 추출 실패")

    return CrashResult(
        project_name=project_name,
        contract_meta_ref=contract_meta_ref,
        months=months,
        period_start=min(all_dates),
        period_end=max(all_dates),
    )
