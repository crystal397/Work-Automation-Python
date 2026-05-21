# 7케이스 baseline 회귀 테스트

공기연장 간접비 산정 빌더(`37_cost_aggregation`)의 회귀를 검출하는 pytest 스위트.

## 검증 가치

CLAUDE.md 핵심 원칙 — **7케이스 baseline 산정값 절대 깨지지 않도록 검증**.
빌더 코드 한 줄 수정으로 산정값이 변하면 즉시 RED.

## 케이스 (7개)

pytest fixture 에 등록된 7케이스 baseline. 실제 케이스명·산정값은 보안상 코드 내부에서만 관리 (`tests/regression/test_baseline_cost.py` 의 `CASES` dict).

| ID | 유형 |
|---|---|
| `case1` ~ `case3` | 단독 산정 (원도급사만) |
| `case4` | 도시철도 분리 산정 (하도급 산재율 분기) |
| `case5` | 하도급 4종 요율 override |
| `case6` | 다중 하도급사 별도 집계 |
| `case7` | 도로공사 장기 케이스 |

> pytest ID 는 ASCII — Windows PowerShell 한글 인자 통과 문제 회피.
> 실제 케이스명·금액은 외부 노출 차단을 위해 본 문서에서는 ID 만 표시.

## 검증 항목 (테스트 2 × 7케이스 = 14 항목)

| 테스트 | 검증 |
|---|---|
| `test_grand_total[...]` | `aggregate.grand_total.value` 가 baseline 과 정확 일치 (단일 정수) |
| `test_cost_result_deep_equal[...]` | `cost_result.json` 전체 dict 가 baseline 과 deep-equal (정규화 후) |

deep_equal 은 grand_total 만이 아니라 prime/subs subtotal + direct_expense items + indirect_labor + rates 전체를 회귀 검증.

## 정규화 규칙

cost_result.json 에는 timestamp 필드가 없어 거의 모든 필드를 직접 비교한다. 단 1개 예외:

- **`contract_meta_ref`** — fresh build 는 절대경로 (`C:/.../36_contract_meta/.../`), baseline 은 상대경로 (`..\\36_contract_meta\\...`) 라서 백슬래시→슬래시 통일 + `36_contract_meta/` 부터 keep. 경로 메타이지 산정값과 무관.

## 동작 원리 — 디스크 비파괴 RETEST

CLAUDE.md "RETEST 표준 패턴" 의 **in-process 버전**:

```
1. baseline cost_result.json 디스크 그대로 두기
2. build_cost_result(cost_input.yaml, contract_meta.json) 직접 호출
    → fresh CostResult 객체 (메모리)
3. fresh.model_dump_json() → dict 로 직렬화 (디스크 안 씀)
4. 디스크 baseline 과 deep-equal
```

기존 RETEST 가 `Copy-Item *.bak` → `cost-agg build` (덮어쓰기) → diff → `Remove-Item *.bak` 4 단계였던 것을 단일 함수 호출로 압축. backup/restore 불필요.

## 실행

```powershell
$env:PYTHONIOENCODING="utf-8"
python -m pytest pipeline_guide/tests/regression/test_baseline_cost.py -v
```

### 단일 케이스만

```powershell
python -m pytest "pipeline_guide/tests/regression/test_baseline_cost.py::test_grand_total[case1]" -v
python -m pytest "pipeline_guide/tests/regression/test_baseline_cost.py::test_cost_result_deep_equal[case5]" -v
```

### grand_total 만 / deep_equal 만

```powershell
python -m pytest pipeline_guide/tests/regression/test_baseline_cost.py -k "grand_total" -v
python -m pytest pipeline_guide/tests/regression/test_baseline_cost.py -k "deep_equal" -v
```

## 실측 (2026-05-21)

```
14 passed in 0.52s
```

- grand_total ×7 PASS
- cost_result_deep_equal ×7 PASS

> 12 passed / 2 skipped 상태에서 B-029 해결로 14 passed 도달.

## 알려진 제약

### ~~B-029~~ — 옛 schema 잔존 케이스 (해결됨 2026-05-20)
의뢰인 산정 보고서 (작성중 최신본) 의 4.3.1 원도급 + 4.4.1 하도급 집계표 인용으로 신 `prime/subs[]` schema 재작성 완료. builder 산출과 의뢰인 표기 사이에 천원절사 단위 차이 발생 가능.

### pytest ID ASCII 제약
Windows PowerShell 에서 한글 parametrize ID 가 `\uXXXX` escape 로 변환돼 명령행 인자 통과가 불가. ASCII ID 매핑 (case1~case7) 으로 회피. 실제 케이스명·금액은 `case_name`·`baseline` 인자로 fixture 내부에서만 사용.

## 어떻게 작동하는가 (회귀 발생 시나리오)

빌더 코드 한 줄 수정 예시:
```python
# 37_cost_aggregation/src/cost_aggregation/calc.py
# 가령 일할 산정 분모를 30 → 31 로 (실수) 변경
days_in_month = 31  # was 30
```

→ pytest 실행:
```
FAILED test_grand_total[case4]
  [<프로젝트>] grand_total 회귀 — fresh=AAA,AAA,000 baseline=BBB,BBB,000 (diff=+N,NNN,000)
FAILED test_cost_result_deep_equal[case4]
  [<프로젝트>] cost_result.json 회귀 — 첫 차이: $.aggregate.prime.indirect_labor_total_krw.value — fresh=… vs baseline=…
```

`_find_first_diff()` 가 JSONPath 풍 경로 + 값 차이를 즉시 노출 → 어느 산식이 깨졌는지 진단 쉬움.
