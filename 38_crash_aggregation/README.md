# 38_crash_aggregation

**돌관공사비** 노무비 집계·산출 자동화. 12_manhour_aggregation 과 24_crash_construction 의 기능을 36/37 의 audit 규약(`Sourced[T]`, 출처 메타 의무) 위에서 재정리.

## 범위 (공기연장 간접비 와 별도 트랙)

| 영역 | 구분 |
|---|---|
| 공기연장 간접비 (현장사무소 인원 월급) | 37_cost_aggregation 담당 |
| **돌관공사비 (일용노무비, 공종별 일별 공수)** | **38_crash_aggregation 담당** |

## 12 / 24 통합 plan

### 12 의 기능
- 회사별(기장원자로·㈜대형건설사1·삼화피앤씨·스마트에스지) 일용노무비 자료(xlsx·PDF 혼합) → 공수 집계
- `aggregator.py`: 업체별 공수 합산
- `filler.py`: 산출내역서 노임 시트 기입
- `formula_writer.py`: 합계 수식 자동 작성
- `pdf_to_excel.py`: PDF 노임명세서 → 페이지별 시트 Excel
- `readers/`: PDF·Excel 파일 형식별 읽기 모듈

### 24 의 기능
- ㈜C산업 월별 노무비 xlsx (공종별 작업자) → 돌관공사비 산출근거 시트 자동 작성
- 공종별 헤더·머지 셀·인쇄 영역·페이지 나눔(48행)·휴일 표시(holidays) 처리
- v12 최신 — 1일차 데이터 누락 버그 수정 등

### 통합 후 38 구조 (제안)
```
38_crash_aggregation/
├── pyproject.toml + README.md
├── src/crash_aggregation/
│   ├── models.py                  # WorkerMonth, Crew, MonthlyCrash, AggregatedCrash
│   ├── audit.py                   # contract_meta.audit 재사용
│   ├── extractors/
│   │   ├── crew_xlsx.py           # 24 패턴: 월별 공종별 xlsx → WorkerMonth[]
│   │   ├── company_pdf.py         # 12 패턴: 회사별 PDF 노임명세서 → WorkerMonth[]
│   │   ├── company_xlsx.py        # 12 패턴: 회사별 xlsx → WorkerMonth[]
│   │   └── unifier.py             # 다양한 어댑터 결과를 단일 스키마로 통합
│   ├── builders/
│   │   ├── monthly_report.py      # 24 의 출력일보 (xlsx 시트 채우기 + 인쇄 영역)
│   │   └── industry_summary.py    # 12 의 회사별 공수 집계
│   ├── validators/
│   │   ├── holidays_kr.py         # 한국 공휴일 자동 표시
│   │   └── consistency.py         # 공수↔금액 일치 검증
│   └── cli.py
└── out/<project>/
    ├── workers.json               # 통합 인원 데이터 (출처 포함)
    ├── crash_report.xlsx          # 24 양식 출력일보
    └── findings.md
```

### 우선순위 (구현 단계)
1. **24 v12 → 38 클린 구현** (현재 활성 케이스 — ㈜C산업)
2. **12 회사별 어댑터 통합** (4개 회사 형식 흡수)
3. **공기연장 간접비 보고서와 연동** — 돌관공사비도 보고서에 첨부되는 케이스 처리
4. **36/37 의 `_source` 규약 적용** — 모든 셀에 audit trail

## 현재 상태

**골격만 마련. 12·24 의 기존 동작은 그대로 유지** — 새 구현이 완성되기 전까지 기존 스크립트를 계속 사용한다.

본격 구현은 별도 세션에서 진행 예정. 36/37 의 패턴을 따라:
- pydantic 모델 단일 진실원
- 모든 leaf 값에 `_source` 의무
- audit 가능한 산식 (입력 셀 → 출력 셀)
- CLI: `init / build / extract / schema`
