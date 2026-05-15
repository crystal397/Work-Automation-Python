"""crash_report.md 생성."""

from __future__ import annotations

from collections import defaultdict

from crash_aggregation.models import CrashResult


def _fmt(n: int) -> str:
    return f"{n:,}"


def build_report(r: CrashResult) -> str:
    lines: list[str] = []
    lines.append(f"# 돌관공사비 노무 집계 — {r.project_name}")
    lines.append("")
    lines.append(f"- 기간: {r.period_start} ~ {r.period_end} ({len(r.months)}개월)")
    lines.append(f"- 총 인원(중복 제외): {r.total_workers}명")
    lines.append(f"- 총 공수: {r.total_manday:,.1f}일")
    lines.append(f"- 총 노무비: **{_fmt(r.total_krw)}원**")
    lines.append("")

    lines.append("## 월별 집계")
    lines.append("")
    lines.append("| 월 | 인원 | 공수(일) | 노무비(원) |")
    lines.append("|---|---:|---:|---:|")
    for m in r.months:
        lines.append(f"| {m.label} | {len({w.name.value for w in m.workers})} | {sum(w.total_manday for w in m.workers):,.1f} | {_fmt(sum(w.total_krw for w in m.workers))} |")
    lines.append("")

    lines.append("## 공종별 합계")
    lines.append("")
    by_gongjong: dict[str, dict] = defaultdict(lambda: {"manday": 0.0, "krw": 0, "workers": set()})
    for m in r.months:
        for w in m.workers:
            g = w.gongjong.value
            by_gongjong[g]["manday"] += w.total_manday
            by_gongjong[g]["krw"] += w.total_krw
            by_gongjong[g]["workers"].add(w.name.value)
    lines.append("| 공종 | 인원(고유) | 공수(일) | 노무비(원) |")
    lines.append("|---|---:|---:|---:|")
    for g, v in sorted(by_gongjong.items(), key=lambda x: -x[1]["krw"]):
        lines.append(f"| {g} | {len(v['workers'])} | {v['manday']:,.1f} | {_fmt(v['krw'])} |")
    lines.append("")

    lines.append("## 카테고리(투입 분류)별 합계")
    lines.append("")
    cat_buckets: dict[str, dict] = defaultdict(lambda: {"manday": 0.0, "krw": 0})
    for m in r.months:
        for c in m.category_totals:
            cat_buckets[c.category]["manday"] += c.total_manday
            cat_buckets[c.category]["krw"] += c.total_krw
    lines.append("| 카테고리 | 공수(일) | 노무비(원) |")
    lines.append("|---|---:|---:|")
    for c, v in sorted(cat_buckets.items(), key=lambda x: -x[1]["krw"]):
        lines.append(f"| {c} | {v['manday']:,.1f} | {_fmt(v['krw'])} |")
    lines.append("")

    return "\n".join(lines)
