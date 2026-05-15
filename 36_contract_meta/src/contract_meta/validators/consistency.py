"""필드 일관성 검증.

검증 항목
- 금액: 한글표기 ↔ 숫자(KRW) 일치
- 기간: period_start + duration_days - 1 == period_end
- 변경계약: revisions[i].duration_days == revisions[i-1].duration_days + revisions[i].duration_diff_days
- 산정 대상 기간: calculation_target.days == (period_end - period_start) + 1 - sum(excluded.days)
"""

from __future__ import annotations

import re
from dataclasses import dataclass, field
from datetime import date

from contract_meta.models import ContractMeta, ContractTrack, Money

_DIGITS = {"일": 1, "이": 2, "삼": 3, "사": 4, "오": 5, "육": 6, "칠": 7, "팔": 8, "구": 9}
_UNITS_SMALL = {"십": 10, "백": 100, "천": 1000}
_UNITS_BIG = {"만": 10**4, "억": 10**8, "조": 10**12}


def parse_korean_amount(text: str) -> int:
    """'금일천팔백삼십구억구천육백만원' → 183_996_000_000"""
    s = text.strip()
    if s.startswith("금"):
        s = s[1:]
    if s.endswith("원"):
        s = s[:-1]
    s = s.strip()
    if not s:
        return 0

    parts = re.split(r"([조억만])", s)
    total = 0
    i = 0
    while i < len(parts):
        chunk = parts[i]
        if i + 1 < len(parts) and parts[i + 1] in _UNITS_BIG:
            total += _parse_small(chunk) * _UNITS_BIG[parts[i + 1]]
            i += 2
        else:
            total += _parse_small(chunk)
            i += 1
    return total


def _parse_small(s: str) -> int:
    if not s:
        return 0
    n = 0
    cur = 0
    for ch in s:
        if ch in _DIGITS:
            cur = _DIGITS[ch]
        elif ch in _UNITS_SMALL:
            n += (cur if cur else 1) * _UNITS_SMALL[ch]
            cur = 0
    n += cur
    return n


# ─────────────────────────────────────────────────────────────────

@dataclass
class CheckResult:
    name: str
    ok: bool
    detail: str


@dataclass
class ValidationReport:
    passed: list[CheckResult] = field(default_factory=list)
    failed: list[CheckResult] = field(default_factory=list)
    warnings: list[CheckResult] = field(default_factory=list)

    def add(self, r: CheckResult, *, severity: str = "fail") -> None:
        if r.ok:
            self.passed.append(r)
        elif severity == "warn":
            self.warnings.append(r)
        else:
            self.failed.append(r)


def _check_money(label: str, m: Money | None) -> CheckResult | None:
    if m is None or m.korean is None:
        return None
    parsed = parse_korean_amount(m.korean.value)
    ok = parsed == m.krw.value
    return CheckResult(
        name=f"금액 일치 — {label}",
        ok=ok,
        detail=(
            f"한글 '{m.korean.value}' → 산출 {parsed:,} vs 기재 {m.krw.value:,}"
            + ("" if ok else " ❌")
        ),
    )


def _check_period(label: str, start: date, end: date, days: int) -> CheckResult:
    actual = (end - start).days + 1
    return CheckResult(
        name=f"기간 산술 — {label}",
        ok=actual == days,
        detail=f"{start} ~ {end} = {actual}일 vs 기재 {days}일",
    )


def _check_track(label: str, track: ContractTrack) -> list[CheckResult]:
    results: list[CheckResult] = []
    prev_days = track.initial.duration_days.value
    results.append(_check_period(
        f"{label} 최초",
        track.initial.period_start.value,
        track.initial.period_end.value,
        prev_days,
    ))
    for r in track.revisions:
        results.append(_check_period(
            f"{label} {r.seq}회",
            r.period_start.value,
            r.period_end.value,
            r.duration_days.value,
        ))
        expected = prev_days + r.duration_diff_days.value
        results.append(CheckResult(
            name=f"누계 — {label} {r.seq}회",
            ok=expected == r.duration_days.value,
            detail=f"직전 {prev_days}일 + 증감 {r.duration_diff_days.value}일 = {expected}일 vs 기재 {r.duration_days.value}일",
        ))
        prev_days = r.duration_days.value
        money_check = _check_money(f"{label} {r.seq}회", r.amount)
        if money_check is not None:
            results.append(money_check)
    return results


def validate(meta: ContractMeta) -> ValidationReport:
    rep = ValidationReport()

    # 최초 계약 금액
    for label, track in [("총공사", meta.total_contract), ("1차수", meta.first_year_contract)]:
        money_check = _check_money(f"{label} 최초", track.initial.amount)
        if money_check is not None:
            rep.add(money_check)
        for r in _check_track(label, track):
            rep.add(r)

    # 산정 대상 기간
    if meta.calculation_target is not None:
        ct = meta.calculation_target
        gross = (ct.period_end.value - ct.period_start.value).days + 1
        excluded = sum(e.days for e in ct.excluded_periods)
        expected = gross - excluded
        rep.add(CheckResult(
            name="산정 대상 일수",
            ok=expected == ct.days.value,
            detail=f"{ct.period_start.value} ~ {ct.period_end.value} = {gross}일, 제외 {excluded}일 → {expected}일 vs 기재 {ct.days.value}일",
        ))

    return rep
