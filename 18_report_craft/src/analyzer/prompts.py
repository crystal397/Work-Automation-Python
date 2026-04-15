"""
분석용 프롬프트 및 JSON 스키마 가이드
"""

ANALYSIS_SCHEMA_GUIDE = """
> **중요**: 이 문서와 함께 프로젝트 루트의 `분석지시서.md`를 반드시 읽고 지침을 준수하세요.
> 특히 **하도급사 목록**, **간접노무 급여**, **요율 확인** 항목은 분석지시서의 기준을 우선 적용합니다.

아래 JSON 형식으로 `analysis_result.json` 파일을 작성하세요.

**핵심 규칙**
- 문서에서 직접 확인된 값만 기재. 불확실하면 `null` (임의 추정·채워넣기 금지).
- 모든 항목의 `_source` 필드에 출처 필수 기재. 예: `"계약서.pdf | p.1"`
- 날짜는 `YYYY-MM-DD` 형식.
- 하도급사는 **하도급계약서에 명시된 업체만** 포함 (공문·구두 언급 업체 제외).
- FAIL 파일 관련 항목은 값 뒤에 `[출처 확인 필요]` 표기.

**금액 데이터 처리 방침 (중요)**
- `monthly_salary` (월급여): **null 기재** — 외부 확인 후 수동 입력 예정
- `amount_actual` / `amount_estimated` (경비 금액): **null 기재** — 외부 확인 후 수동 입력 예정
- 대신 **항목명(item), 인원명(name), 소속(org), 직무(role), 기간(period)은 최대한 추출**
- 요율(industrial_accident, employment, general_admin, profit)은 문서에서 확인되면 기재
- 금액이 명확히 기재된 공식 문서(계약서 금액, 변경계약 금액 등)는 예외로 기재

**유형 분류 기준**
- `"A"`: "지방자치단체를 당사자로 하는 계약에 관한 법률" 인용
- `"B"`: "국가를 당사자로 하는 계약에 관한 법률" 인용
- `"C"`: 민간·혼합 계약 (공사계약일반조건 단독 적용)

```json
{
  "report_type": "A",
  "report_type_reason": "지방자치단체를 당사자로 하는 계약에 관한 법률 제22조 인용",

  "contract": {
    "name": "공사명",
    "name_source": "계약서.pdf | p.1",
    "client": "발주처명",
    "client_source": "출처",
    "contractor": "계약상대자명 (JV 시 각 구성원명 및 지분율 포함)",
    "contractor_source": "출처",
    "jv_share_ratio": null,
    "_jv_share_ratio_comment": "공동수급체(JV)인 경우 의뢰사 지분율 (예: 60% → 0.6). 단독 수급이면 null 또는 생략.",
    "initial_date": "YYYY-MM-DD",
    "initial_date_source": "출처",
    "initial_amount": 0,
    "initial_amount_source": "출처",
    "initial_start": "YYYY-MM-DD",
    "initial_end": "YYYY-MM-DD",
    "contract_type": "장기계속계약 또는 총액계약 또는 기타",
    "contract_type_source": "출처",
    "supervisor": "건설사업관리단명 (없으면 null)",
    "supervisor_source": "출처",
    "scope_description": "공사 주요내용 1~2줄 요약 (없으면 null)"
  },

  "total_contract": {
    "_comment": "장기계속계약의 총공사 계약 현황. 총액계약이면 null 또는 생략.",
    "initial_date": "YYYY-MM-DD",
    "initial_amount": 0,
    "initial_start": "YYYY-MM-DD",
    "initial_end": "YYYY-MM-DD",
    "final_amount": 0,
    "total_days": "총공사 기간 (일수 또는 기간 표기)",
    "changes": [
      {
        "seq": 1,
        "date": "YYYY-MM-DD",
        "amount": 0,
        "new_end_date": "YYYY-MM-DD",
        "reason": "변경 사유"
      }
    ],
    "source": "출처"
  },

  "vat_zero_rated": false,
  "_vat_zero_rated_comment": "부가가치세 영세율 적용 여부. 일반적으로 false(10%). 영세율 근거가 있으면 true로 변경.",

  "changes": [
    {
      "seq": 1,
      "date": "YYYY-MM-DD",
      "date_source": "출처",
      "reason": "변경 사유 (예: 교통처리계획 협의 지연)",
      "reason_source": "출처",
      "extension_days": 0,
      "new_end_date": "YYYY-MM-DD",
      "amount": 0
    }
  ],

  "extension": {
    "total_days": 0,
    "start_date": "YYYY-MM-DD",
    "end_date": "YYYY-MM-DD",
    "source": "출처"
  },

  "laws": [
    {
      "name": "지방자치단체를 당사자로 하는 계약에 관한 법률",
      "article": "제22조",
      "purpose": "계약금액 조정 근거",
      "source": "출처"
    }
  ],

  "indirect_labor": [
    {
      "name": "이름",
      "org": "소속",
      "role": "직무",
      "period_start": "YYYY-MM-DD",
      "period_end": "YYYY-MM-DD",
      "monthly_salary": 0,
      "retirement_rate": 0.0833,
      "source": "출처"
    }
  ],

  "expenses_direct": [
    {
      "item": "항목명 (예: 지급임차료)",
      "amount_actual": 0,
      "amount_estimated": 0,
      "source": "출처"
    }
  ],

  "rates": {
    "industrial_accident": 0.0,
    "industrial_accident_source": "출처",
    "employment": 0.0,
    "employment_source": "출처",
    "general_admin": 0.0,
    "general_admin_source": "출처",
    "profit": 0.0,
    "profit_source": "출처"
  },

  "correspondence": [
    {
      "date": "YYYY-MM-DD",
      "direction": "발신처→수신처",
      "title": "공문 제목",
      "doc_number": "문서번호 (있을 시)",
      "source": "출처"
    }
  ],

  "subcontractors": [
    {
      "name": "하도급사명",
      "contract_date": "YYYY-MM-DD",
      "period_start": "YYYY-MM-DD",
      "period_end": "YYYY-MM-DD",
      "source": "출처"
    }
  ],

  "unresolved": [
    {
      "item": "확인이 필요한 항목명",
      "reason": "확인 불가 이유",
      "source": "관련 출처 (있을 시)"
    }
  ],

  "source_totals": {
    "_comment": "소스 문서에 기재된 합계 금액. 계산값과 교차검증에 사용.",
    "indirect_labor_total": null,
    "indirect_labor_total_source": "급여대장.xlsx | 합계 행",
    "expenses_total": null,
    "expenses_total_source": "출처",
    "grand_total": null,
    "grand_total_source": "출처"
  }
}
```
"""
