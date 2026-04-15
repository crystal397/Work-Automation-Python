# 공기연장 간접비 보고서 자동 생성 시스템 — 설계 문서

> 수신자료를 넣으면 → 유형 판단 → 수치 계산 → Word 보고서 생성

---

## 전체 흐름

```
input/
 └── 수신자료 (PDF, Excel, HWP, HTML, XML, TIF, Scan)
        │
        ▼
[1단계] file_extractor.py
        포맷별 텍스트 추출
        + 출처 태깅 (파일명 | 페이지/시트)
        │
        ▼
[1.5단계] quality_checker.py          ◀ NEW
        추출 품질 검증
        → OK: 다음 단계 진행
        → WARN/FAIL: extraction_report.md 생성 후 중단
        │
        ▼
[2단계] analyzer.py  ← Claude API
        정보 구조화 + 유형 판단 (A/B/C)
        + 각 추출 항목에 출처 기록
        │
        ▼
[3단계] calculator.py
        간접비 수치 계산
        + 수치마다 출처 연결 유지
        │
        ▼
[4단계] generator.py
        마크다운 초안 → Word(.docx) 생성
        + 각 수치 옆에 출처 각주/주석
        │
        ▼
output/
 ├── 보고서_초안.md
 ├── 보고서_초안.docx
 └── extraction_report.md    ← 품질 검증 결과 (항상 생성)
```

---

## 폴더 구조

```
18_report_craft/
├── input/                      # 수신자료 투입 폴더
├── output/                     # 생성된 보고서 출력
├── reference/                  # 참고 보고서 13개
├── 법령/                        # 법령 파일 14개
├── src/
│   ├── extractor/
│   │   └── file_extractor.py   # [1단계] 포맷별 텍스트 추출
│   ├── analyzer/
│   │   ├── analyzer.py         # [2단계] Claude API 정보 추출
│   │   └── prompts.py          # 추출용 프롬프트 모음
│   ├── calculator/
│   │   └── calculator.py       # [3단계] 간접비 계산 엔진
│   └── generator/
│       ├── md_generator.py     # [4단계-1] 마크다운 생성
│       └── docx_generator.py   # [4단계-2] Word 문서 생성
├── main.py                     # 진입점 (전체 파이프라인 실행)
├── config.py                   # API 키, 경로 등 설정
└── ARCHITECTURE.md             # 이 문서
```

---

## 1단계 — 포맷별 텍스트 추출 (`file_extractor.py`)  +  1.5단계 품질 검증 (`quality_checker.py`)

### 역할
`input/` 폴더의 모든 파일을 읽어 텍스트로 변환.  
이후 단계는 포맷에 상관없이 **순수 텍스트**만 다룬다.

### 포맷별 처리 방식

| 포맷 | 라이브러리 | 처리 방식 |
|------|-----------|----------|
| PDF (텍스트) | pdfplumber | 페이지별 텍스트 직접 추출 |
| PDF (스캔) | pdfplumber + Tesseract | 이미지 변환 후 OCR |
| Excel (.xlsx/.xls) | pandas | 시트별 행 추출, `\|` 구분 |
| HTML | BeautifulSoup | script/style 제거 후 텍스트 |
| XML | lxml | 모든 노드 텍스트 순회 |
| HWP / HWPX | win32com (한컴 COM) | HWP 앱 자동화 → GetText() |
| TIF / 이미지 | Tesseract OCR | 멀티프레임 처리, kor+eng |

### 출력 구조 (페이지/시트 단위로 출처 포함)

```python
# ExtractResult 객체 리스트
[
  ExtractResult(
    file      = "계약서.pdf",
    format    = "pdf",
    chunks    = [
      Chunk(source="계약서.pdf | p.1",  text="광주 도시철도 2호선...", method="text"),
      Chunk(source="계약서.pdf | p.2",  text="계약금액 일천육백...",   method="text"),
      Chunk(source="계약서.pdf | p.5",  text="",                       method="ocr",  quality="EMPTY"),
    ],
    quality   = "WARN",          # OK / WARN / FAIL
    issues    = ["p.5 OCR 결과 없음 — 수동 확인 필요"]
  ),
  ExtractResult(
    file      = "급여명세서.xlsx",
    format    = "excel",
    chunks    = [
      Chunk(source="급여명세서.xlsx | Sheet1 | r2", text="홍길동 | 소장 | 8500000", method="direct"),
    ],
    quality   = "OK",
    issues    = []
  ),
  ...
]
```

### 출처(source) 형식

| 포맷 | 출처 형식 |
|------|----------|
| PDF | `파일명.pdf \| p.{페이지번호}` |
| PDF(OCR) | `파일명.pdf \| p.{페이지번호} [OCR]` |
| Excel | `파일명.xlsx \| {시트명} \| r{행번호}` |
| HTML | `파일명.html \| section.{n}` |
| XML | `파일명.xml \| node.{태그명}` |
| HWP | `파일명.hwp \| para.{n}` |
| TIF | `파일명.tif \| frame.{n} [OCR]` |

---

## 1.5단계 — 품질 검증 (`quality_checker.py`)

### 검증 항목

| 검증 항목 | 판정 기준 | 결과 |
|----------|----------|------|
| 빈 추출 | 텍스트 길이 0 | FAIL |
| 극소량 텍스트 | 100자 미만 (이미지·스캔 제외) | WARN |
| 한국어 비율 낮음 | 한글 문자 비율 < 10% (한글 문서 예상 시) | WARN |
| 깨진 문자 | `?`, `□`, `▯` 등 대체 문자 비율 > 5% | WARN |
| OCR 신뢰도 낮음 | Tesseract confidence < 60 | WARN |
| Excel 빈 시트 | 유효 행 0개 | WARN |
| HWP COM 오류 | 예외 발생 | FAIL |

### 품질 등급

```
OK   → 모든 항목 정상
WARN → 일부 이슈 있으나 계속 진행 가능 (사람 확인 권장)
FAIL → 추출 실패, 반드시 수동 확인 후 재투입 필요
```

### 출력: `output/extraction_report.md` (항상 생성)

```markdown
# 추출 품질 보고서
생성일시: 2026-04-03 10:30

## 요약
- 전체 파일: 12개
- OK: 9개 | WARN: 2개 | FAIL: 1개

## ✅ 정상 파일
| 파일명 | 포맷 | 추출 분량 |
| 계약서.pdf | PDF | 4,231자 (12페이지) |
...

## ⚠️ 확인 권장 (WARN)
| 파일명 | 이슈 |
| 영수증_스캔.tif | frame.3 OCR 신뢰도 42% — 수동 확인 필요 |
| 내역서.pdf | p.8 텍스트 추출 0자, OCR 전환 후 230자 |

## ❌ 재투입 필요 (FAIL)
| 파일명 | 원인 | 조치 방법 |
| 공문_구버전.hwp | HWP COM 열기 실패 | PDF로 변환 후 재투입 |

## 출처 인덱스 (전체)
| 항목 | 출처 |
| 계약체결일: 2023-12-28 | 계약서.pdf \| p.1 |
| 일반관리비율: 4.5% | 산출내역서.xlsx \| Sheet1 \| r15 |
...
```

---

## 2단계 — 정보 추출 + 유형 판단 (`analyzer.py`)

### 역할
추출된 텍스트 전체를 Claude API에 전달하여  
보고서 작성에 필요한 **구조화된 데이터**를 JSON으로 추출.

### Claude API 호출 구조

```
[시스템 프롬프트]
  - 공기연장 간접비 보고서 작성 전문가 역할
  - 지방계약법(A) / 국가계약법(B) / 민간(C) 판단 기준 주입
  - 추출 대상 항목 명시

[유저 프롬프트]
  - 추출된 문서 텍스트 전체

[출력: JSON]
  - 계약 정보
  - 변경 이력
  - 유형 (A/B/C)
  - 인원 및 급여 데이터
  - 경비 항목
  - 요율 정보
  - 공문 이력
```

### 추출 항목 상세

```json
{
  "report_type": "A",          // A(지방계약법) / B(국가계약법) / C(민간)
  "contract": {
    "name": "공사명",
    "client": "발주처",
    "contractor": "계약상대자",
    "initial_date": "2023-12-28",
    "initial_amount": 166168310000,
    "initial_end_date": "2028-10-22"
  },
  "changes": [
    {
      "seq": 1,
      "date": "2025-01-21",
      "reason": "경찰청 단계별 교통처리계획 협의 지연",
      "extension_days": 159,
      "new_end_date": "2025-06-30"
    }
  ],
  "extension": {
    "total_days": 159,
    "start_date": "2025-01-23",
    "end_date": "2025-06-30"
  },
  "indirect_labor": [
    {
      "name": "홍길동",
      "org": "시공사",
      "role": "소장",
      "period_start": "2025-01-23",
      "period_end": "2025-06-30",
      "monthly_salary": 8500000,
      "retirement_rate": 0.0833
    }
  ],
  "expenses_direct": [
    { "item": "지급임차료", "amount": 36278715 },
    { "item": "복리후생비", "amount": 20500765 }
  ],
  "rates": {
    "industrial_accident": 0.037,
    "employment": 0.0157,
    "general_admin": 0.045,
    "profit": 0.0571
  },
  "correspondence": [
    {
      "date": "2025-01-15",
      "direction": "계약상대자→감리단",
      "title": "제1회 공사기간 연장 요청 실정보고"
    }
  ]
}
```

---

## 3단계 — 간접비 계산 엔진 (`calculator.py`)

### 역할
2단계 JSON을 입력받아 항목별 금액을 계산.  
**실비 구간**과 **추정 구간**을 분리하여 산정.

### 계산 흐름

```
[입력: analyzer JSON]
        │
        ▼
1. 간접노무비
   = Σ(일수 × 일급) + 퇴직급여충당금
   ※ 준공일 도래 여부로 실비/추정 분기
        │
        ▼
2. 경비
   ├ 직접계상: 항목별 실비 합산
   └ 승률계상:
      산재보험료 = 간접노무비 × 산재요율
      고용보험료 = 간접노무비 × 고용요율
        │
        ▼
3. 소계 = 간접노무비 + 경비
        │
        ▼
4. 일반관리비 = 소계 × 일반관리비율
   (한도: 지방계약법 6%, 국가계약법 300억↑ 5%)
        │
        ▼
5. 이윤 = (소계 + 일반관리비) × 이윤율
   (한도: 15%)
        │
        ▼
6. 총원가 = 소계 + 일반관리비 + 이윤
        │
        ▼
[출력: 집계표 dict]
```

### 출력 예시

```python
{
  "indirect_labor":   { "actual": 242368042, "estimated": 336931139, "total": 579299181 },
  "expenses": {
    "direct":         { "actual": 148115795, "estimated": 198213778, "total": 346329573 },
    "rate_based": {
      "industrial":   { "actual":   8967617, "estimated":  12466452, "total":  21434069 },
      "employment":   { "actual":   3805178, "estimated":   5289818, "total":   9094996 },
    }
  },
  "subtotal":         { "actual": 403256632, "estimated": 552901187, "total": 956157819 },
  "general_admin":    { "actual":  18146548, "estimated":  24880553, "total":  43027101 },
  "profit":           { "actual":  24062121, "estimated":  32991337, "total":  57053458 },
  "grand_total":      { "actual": 445465301, "estimated": 610773077, "total": 1056238378 },
  "final_rounded":    1056238000
}
```

---

## 4단계 — 보고서 생성 (`md_generator.py` + `docx_generator.py`)

### 4-1. 마크다운 생성 (`md_generator.py`)

유형별 템플릿(A/B/C)에 2·3단계 데이터를 채워 마크다운 초안 생성.

```
report_template_guide.md 의 템플릿 구조
        +
analyzer JSON (계약 정보, 공문 이력 등)
        +
calculator 결과 (집계표 수치)
        │
        ▼
output/보고서_초안.md
```

### 4-2. Word 생성 (`docx_generator.py`)

마크다운을 python-docx로 변환.

| 마크다운 요소 | Word 스타일 |
|-------------|------------|
| `# 제목` | Heading 1 |
| `## 제목` | Heading 2 |
| `### 제목` | Heading 3 |
| 일반 문단 | Normal |
| `\|표\|` | 표(Table) — 자동 열 너비 |
| `**굵게**` | Bold |
| `` `코드` `` | 고정폭 |

---

## 진입점 (`main.py`)

```
python main.py [--input 폴더경로] [--output 폴더경로]
```

실행 순서:
1. `input/` 폴더의 모든 파일 텍스트 추출
2. Claude API로 정보 구조화
3. 간접비 계산
4. 마크다운 + Word 생성
5. `output/` 저장

---

## 설정 (`config.py`)

모든 민감값은 `.env` 또는 환경변수로 관리합니다. `.env.example` 참고.

```python
# .env 또는 환경변수로 설정
REPORT_INPUT_DIR   = ...   # 수신자료 폴더 (기본: input/)
REPORT_OUTPUT_DIR  = ...   # 결과물 저장 (기본: output/)
REPORT_AUTHOR      = ...   # 보고서 작성 회사명

# config.py 내 고정값
TESSERACT_PATH  = r"C:\Program Files\Tesseract-OCR\tesseract.exe"
TESSERACT_LANG  = "kor+eng"
LAW_DIR         = "법령"
TEMPLATE_GUIDE  = "report_template_guide.md"
```

---

## 폴더 구조 (최종)

```
18_report_craft/
├── input/                          # 수신자료 투입 폴더
├── output/                         # 생성된 보고서 출력
│   ├── 보고서_초안.md
│   ├── 보고서_초안.docx
│   └── extraction_report.md        # 품질 검증 결과 (항상 생성)
├── reference/                      # 참고 보고서 13개
├── 법령/                            # 법령 파일 14개
├── src/
│   ├── extractor/
│   │   ├── file_extractor.py       # [1단계] 포맷별 추출 + 출처 태깅
│   │   └── quality_checker.py      # [1.5단계] 품질 검증 + 보고서 생성
│   ├── analyzer/
│   │   ├── prompts.py              # Claude API 프롬프트
│   │   └── analyzer.py             # [2단계] 정보 구조화 + 유형 판단
│   ├── calculator/
│   │   └── calculator.py           # [3단계] 간접비 계산
│   └── generator/
│       ├── md_generator.py         # [4단계-1] 마크다운 생성
│       └── docx_generator.py       # [4단계-2] Word 생성
├── main.py                         # 진입점
├── config.py                       # 설정값
└── ARCHITECTURE.md
```

## 개발 순서

| 순서 | 파일 | 상태 |
|------|------|------|
| 1 | `src/extractor/file_extractor.py` | ✅ 완료 |
| 2 | `src/extractor/file_classifier.py` | ✅ 완료 |
| 3 | `src/extractor/quality_checker.py` | ✅ 완료 |
| 4 | `config.py` | ✅ 완료 |
| 5 | `src/analyzer/prompts.py` | ✅ 완료 |
| 6 | `src/analyzer/analyzer.py` | ✅ 완료 |
| 7 | `src/analyzer/data_checker.py` | ✅ 완료 |
| 8 | `src/calculator/calculator.py` | ✅ 완료 |
| 9 | `src/generator/laws_db.py` | ✅ 완료 |
| 10 | `src/generator/build_templates.py` | ✅ 완료 |
| 11 | `src/generator/md_generator.py` | ✅ 완료 |
| 12 | `src/generator/docx_generator.py` | ✅ 완료 |
| 13 | `main.py` | ✅ 완료 |

---

## 유의사항

- **법령 자동 인용**: `법령/` 폴더의 파일명으로 해당 조문을 매칭하여 보고서에 삽입
- **할루시네이션 방지**: 법령 조항 번호는 직접 문서에서 추출한 것만 사용, 확인 불가 시 `[법령 확인 필요]` 표기
- **추정 산정 명시**: 준공일 미도래 구간은 실비(A)와 추정(B)을 반드시 분리
- **하도급 포함 감지**: 수신자료에 하도급 급여명세서가 있으면 자동으로 하도급 섹션 추가
