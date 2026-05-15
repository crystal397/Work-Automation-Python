"""돌관공사비 노무 데이터 모델.

24 v12 의 구조 (시트 = `NN.공종명`, 블록 = '작업자명' 헤더 단위) 를 그대로 흡수.
36/37 의 audit 규약(`Sourced[T]`, `_source` 의무) 적용.
"""

from __future__ import annotations

from datetime import date
from typing import Literal

from contract_meta.models import Sourced
from pydantic import BaseModel, ConfigDict, Field


# ─────────────────────────────────────────────────────────────────
# 카테고리 enum (24 의 CATEGORIES 와 일치)
# ─────────────────────────────────────────────────────────────────

class Category(str):
    """투입 카테고리. 24 v12 의 classify_category 결과를 그대로 따른다.
    예: '본선', '#01', '#02', '기타', 'BP', '추가', '추가복합', '추가가설공', ...
    """


# ─────────────────────────────────────────────────────────────────
# 노동 데이터
# ─────────────────────────────────────────────────────────────────

class WorkerDay(BaseModel):
    """작업자 1명의 1일 투입."""
    day: int = Field(ge=1, le=31)
    manday: Sourced[float]                         # 공수 (D열)
    amount_krw: Sourced[int]                       # 노무비 (E열)
    category: Sourced[str]                         # 본선 / #01 / 추가 / ...  (F열·H열 분류 결과)
    work_content: Sourced[str] | None = None       # F열 작업내용
    note: Sourced[str] | None = None               # 비고


class WorkerMonth(BaseModel):
    """한 달 동안의 작업자 1명 데이터 (한 공종 기준)."""
    name: Sourced[str]
    gongjong: Sourced[str]                          # 가시설 / 형틀목공 / ...
    role: Sourced[str] | None = None                # 반장 / 조장 / 기능 / 보통
    unit_price_per_day: Sourced[int] | None = None  # B열 '일 기준' 아래 단가
    unit_price_per_hour: Sourced[int] | None = None # B열 '시간 기준' 아래 단가
    days: list[WorkerDay] = Field(default_factory=list)
    block_row_range: tuple[int, int] = Field(description="시트 내 블록 행 범위 (1-base, inclusive)")

    @property
    def total_manday(self) -> float:
        return sum(d.manday.value for d in self.days)

    @property
    def total_krw(self) -> int:
        return sum(d.amount_krw.value for d in self.days)


# ─────────────────────────────────────────────────────────────────
# 월별 / 카테고리 집계
# ─────────────────────────────────────────────────────────────────

class CategoryTotal(BaseModel):
    """카테고리별 월 집계 (예: '본선': 1,234만원)."""
    category: str
    total_manday: float
    total_krw: int
    by_gongjong: dict[str, int] = Field(default_factory=dict)


class MonthlyCrash(BaseModel):
    """월별 돌관공사비 한 묶음 — 한 xlsx 파일 한 달 데이터."""
    year: int
    month: int
    source_file: str
    workers: list[WorkerMonth] = Field(default_factory=list)
    category_totals: list[CategoryTotal] = Field(default_factory=list)

    @property
    def label(self) -> str:
        return f"{self.year}.{self.month:02d}"


# ─────────────────────────────────────────────────────────────────
# 전체 집계 (여러 월)
# ─────────────────────────────────────────────────────────────────

class CrashResult(BaseModel):
    """돌관공사비 산정 결과 (여러 월 합산)."""
    schema_version: Literal["0.1.0"] = "0.1.0"
    project_name: str
    contract_meta_ref: str | None = None
    months: list[MonthlyCrash] = Field(default_factory=list)
    period_start: date
    period_end: date

    @property
    def total_workers(self) -> int:
        names = set()
        for m in self.months:
            for w in m.workers:
                names.add(w.name.value)
        return len(names)

    @property
    def total_manday(self) -> float:
        return sum(w.total_manday for m in self.months for w in m.workers)

    @property
    def total_krw(self) -> int:
        return sum(w.total_krw for m in self.months for w in m.workers)

    model_config = ConfigDict(populate_by_name=True)
