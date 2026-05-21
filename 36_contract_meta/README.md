# 36_contract_meta

공기연장 보고서용 **계약 메타데이터 추출 파이프라인**.

입력으로 받은 계약서·변경계약서·산출내역서를 파싱해, 보고서의 **1·2·3.2·4·결론** 모든 장이 참조할 단일 JSON(`contract_meta.json`) 한 개를 생성한다. 모든 leaf 필드에는 `_source`(파일/페이지/원문) 메타가 의무로 붙는다.

## 설계 원칙

1. **정확도 최우선** — 한글금액↔숫자 일치, 변경계약 누계, 의뢰처 요약본과의 cross-check를 시스템 차원에서 강제.
2. **출처 추적 의무** — 모든 leaf 필드에 `_source` 객체 필수. pydantic validator로 누락 시 에러.
3. **사람 검증 가능** — JSON 외에 `extraction_report.md`(자동/검증필요/교차검증 결과)와 원문 발췌 이미지 동봉.
4. **단일 진실원** — `models.py`(pydantic v2)가 곧 스키마. JSON Schema는 모델에서 자동 생성.

## 빠른 사용법

```bash
pip install -e .
contract-meta extract <input_dir> --project "프로젝트명"
# → out/<프로젝트명>/contract_meta.json
# → out/<프로젝트명>/extraction_report.md
# → out/<프로젝트명>/source_excerpts/*.png
```

## 디렉터리

```
36_contract_meta/
├── pyproject.toml
├── schemas/contract_meta.schema.json   # models.py 에서 자동 생성
├── samples/                            # 사건별 PDF (gitignore)
├── src/contract_meta/
│   ├── models.py                       # pydantic v2 모델 (= 단일 진실원)
│   ├── audit.py                        # _source 메타 빌더
│   ├── cli.py
│   ├── extractors/
│   │   ├── classifier.py               # PDF 양식 판별
│   │   ├── change_contract.py          # 변경계약서 파서
│   │   ├── initial_contract.py         # 최초계약서 파서
│   │   ├── estimate_sheet.py           # 산출내역서(xlsx) 파서
│   │   └── pdf_text.py
│   └── validators/
│       ├── consistency.py              # 한글금액↔숫자, 산술 검증
│       └── cross_doc.py                # 변경계약 누계·기간 연속성
└── out/<project>/
    ├── contract_meta.json
    ├── extraction_report.md
    └── source_excerpts/
```

## 보고서 영역 ↔ 출력 필드 매핑

| 보고서 영역 | 참조 필드 |
|---|---|
| 표지 / 제출문 | `project.name`, `contractor.name_legal` |
| 요약문 | `calculation_target.period`, `calculation_target.days` |
| 1.1.1 공사 개요 | `project`, `owner`, `contractor`, `supervisor`, `scope`, `total_contract`, `first_year_contract` |
| 1.2.3 계약당사자 | `owner`, `contractor` |
| 1.3.1 계약 현황 | `total_contract`, `first_year_contract` |
| 1.3.2 변경계약 표 + 다이어그램 | `first_year_contract.revisions` |
| 2장 분기(국가/지방/민간 · 장기계속 · 내역입찰) | `project.client_type`, `project.contract_form`, `project.bidding_method` |
| 3.2 청구 근거 | 위와 동일 |
| 4.1.1.2 공기연장 일수 산정 | `first_year_contract.revisions[].duration_diff_days`, `calculation_target` |
| 4.3.4 일반관리비/이윤 | `rates.general_admin_percent`, `rates.profit_percent` |
| 4.3 / 4.4 보험요율 | `rates.industrial_accident_insurance_percent`, `rates.employment_insurance_percent` |
| 4.4 하도급사 | `subcontractors[]` |
