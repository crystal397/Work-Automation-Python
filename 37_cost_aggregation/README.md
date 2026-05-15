# 37_cost_aggregation

공기연장 간접비 보고서 **4.3 / 4.4 / 4.5** 자동 생성 파이프라인.
36_contract_meta 의 메타·요율(rates)을 입력으로 받아 산정 결과 JSON + 보고서용 표를 만든다.

## 산출 범위

| 보고서 영역 | 자동화 | 입력 |
|---|---|---|
| 4.3.2.1 대상인원 표 | 자동 | 인원투입현황 xlsx |
| 4.3.2.2 급여내역 표 | 자동(텍스트 PDF) / manual | 급여명세서·지불조서 |
| 4.3.3.1 직접계상비목 | manual + 자동 합산 | 비목별 영수증·전표 |
| 4.3.3.2 승률계상비목 | 자동(산식) | 노무비 × `rates.{산재,고용}_보험_percent` |
| 4.3.4.1 일반관리비 | 자동(산식) | (노무비+경비) × `rates.general_admin_percent` |
| 4.3.4.2 이윤 | 자동(산식) | (노무비+경비+일관) × `rates.profit_percent` |
| 4.3.1 / 4.4.1 집계표 | 자동 | 위 결과 합산 |
| 4.5 결론 표 | 자동 | 원도급사+하도급사 합산 |

## 설계 원칙

1. **모든 leaf 값에 `_source` 의무** — 36 의 `Sourced[T]` 와 동일 규약
2. **산식의 audit trail** — 결과 셀마다 입력 셀·산식·중간값 보관 (예: `47,512,897 = (간접노무비 + 경비) × 4.50%`)
3. **단일 진실원 36** — 요율·기간·계약유형은 `contract_meta.json` 에서 직접 읽음 (재입력 금지)
4. **인원 단위 traceability** — 33명 각자의 급여 → 합계까지 직선 추적

## 디렉터리

```
37_cost_aggregation/
├── pyproject.toml
├── samples/                                 # C공구(철도) 입력 PDF·xlsx (gitignore)
├── src/cost_aggregation/
│   ├── models.py                            # 산정 결과 스키마 (pydantic v2)
│   ├── audit.py                             # Source 빌더 (contract_meta.audit 재사용)
│   ├── cli.py
│   ├── extractors/
│   │   ├── personnel.py                     # 인원투입현황 xlsx → Personnel[]
│   │   ├── salary.py                        # 급여명세서/지불조서 PDF → Salary[]
│   │   └── expense_loader.py                # 경비 manual yaml 로더
│   └── validators/
│       └── arithmetic.py                    # 산식 일관성 검증
└── out/<project>/
    ├── cost_input.yaml                      # 사람이 채우는 부분 (경비 등)
    ├── cost_result.json                     # 산정 결과
    ├── cost_report.md                       # 4.3 / 4.4 / 4.5 표 마크다운
    └── findings.md                          # 검증 실패·경고
```

## CLI

```bash
cost-agg personnel <xlsx>                              # 인원 → 4.3.2.1 표
cost-agg salary <pdf|xlsx>                             # 급여 → 4.3.2.2 표
cost-agg build <input.yaml> --meta <contract_meta.json> # 전체 산정 + 집계표
cost-agg schema                                         # JSON Schema export
```
