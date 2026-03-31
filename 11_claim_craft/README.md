# ClaimCraft — 건설공사 분쟁 보고서 작성 도구

건설공사 분쟁 관련 기술검토 보고서(원인·책임 분석 / 손실금액 적정성 검토) 작성을 위한
Claude.ai 프롬프트 자동 생성 파이프라인입니다.

---

## 디렉토리 구조

```
11_claim_craft/
├── projects/                        # 프로젝트별 폴더 (한 폴더 = 한 사건)
│   └── 창원용원/
│       ├── config.py                # 프로젝트 설정 (이것만 수정하면 됨)
│       ├── 수신자료/                 # 원본 자료 (PDF, HWP, XLSX 등)
│       ├── processed/               # 추출된 텍스트 (extractor.py 출력)
│       ├── 참고예시/                 # 완성된 보고서 예시 (Claude 학습용)
│       │   ├── 참고_part1_개요_예시.txt
│       │   ├── 참고_part2_예시.txt
│       │   ├── 참고_part3_예시.txt
│       │   └── 참고_part4_예시.txt
│       └── prompts_cause/           # 최종 출력 — Claude.ai에 붙여넣을 프롬프트
│
├── templates/                       # 공통 프롬프트 템플릿 (프로젝트 무관)
│   ├── prompt_cause.txt             # 원인·책임 분석 보고서용 템플릿
│   └── prompt.txt                   # 손실금액 적정성 검토 보고서용 템플릿
│
├── 법령_processed/                  # 공통 법령 텍스트 (프로젝트 간 공유)
│
├── extractor.py                     # 수신자료 → processed/ 텍스트 추출
├── generate_prompts_cause.py        # 원인·책임 분석 보고서 프롬프트 생성
└── generate_prompts.py              # 손실금액 적정성 검토 보고서 프롬프트 생성
```

---

## 전체 작업 흐름

```
수신자료/ (PDF·HWP·XLSX)
    ↓  extractor.py
processed/ (텍스트 파일)
    ↓  generate_prompts_cause.py
prompts_cause/ (Claude.ai에 붙여넣을 프롬프트 파일 5개)
    ↓  Claude.ai에서 파트별로 실행
보고서 초안 (Word 등으로 정리)
```

---

## 사용법

### 1단계 — 텍스트 추출

```bash
python extractor.py 창원용원
```

- `projects/창원용원/수신자료/` 내 파일을 읽어 `projects/창원용원/processed/`에 저장
- 프로젝트명 생략 시 기본값 `창원용원` 사용

### 2단계 — 프롬프트 생성

**원인·책임 분석 보고서용:**

```bash
python generate_prompts_cause.py 창원용원
```

`projects/창원용원/prompts_cause/` 에 파트별 프롬프트 5개 생성:

| 파일명 | 보고서 목차 |
|--------|------------|
| `cause_part1_개요.txt` | 1장 과업 개요 및 사업 개요 |
| `cause_part2_사실관계_발주자귀책.txt` | 2.1 사실관계 + 2.2 발주자 귀책 |
| `cause_part3a_시공사귀책_전반.txt` | 2.3.1~2.3.5 시공사 귀책 전반 |
| `cause_part3b_시공사귀책_후반.txt` | 2.3.6~2.3.10 시공사 귀책 후반 (감정사항별) |
| `cause_part4_손실금액_결론.txt` | 3장 손실금액 + 4장 결론 |

**손실금액 적정성 검토 보고서용:**

```bash
python generate_prompts.py 창원용원
```

`projects/창원용원/prompts/` 에 파트별 프롬프트 생성.

### 3단계 — Claude.ai에서 실행

생성된 `.txt` 파일을 파트 순서대로 Claude.ai 대화창에 붙여넣기하여 보고서 초안을 받습니다.

---

## 새 프로젝트 추가 방법

### 1. 프로젝트 폴더 생성

```bash
mkdir "projects/새프로젝트명"
mkdir "projects/새프로젝트명/수신자료"
mkdir "projects/새프로젝트명/참고예시"
```

### 2. config.py 복사 후 수정

```bash
copy "projects/창원용원/config.py" "projects/새프로젝트명/config.py"
```

`config.py`를 열어 아래 항목을 수정합니다:

```python
# 프로젝트 메타데이터
PROJECT_NAME    = "새 공사명"
PLAINTIFF       = "원고명"
DEFENDANT       = "피고명"
CASE_NUMBER     = "사건번호"
CONTRACT_DATE   = "계약일"
CONTRACT_AMOUNT = "계약금액"

# 참조 보고서 (발주자측 손실금액 산정 보고서 파일명, processed/ 기준)
REFERENCE_DOC = "03_감정보고서_XXX.txt"

# 감정보고서 페이지 범위 — 실제 보고서 페이지 구성에 맞게 수정
GDOC_CAUSE_PAGES     = set(range(10, 91))    # 원인분석 파트 페이지
GDOC_31_32_PAGES     = set(range(91, 133))   # 사업주체 책임비율 파트 페이지
GDOC_REST_PAGES      = set(range(446, 492))  # 집행예정·원상복구 파트 페이지
GDOC_SUM_PAGES       = set(range(491, 514))  # 종합의견 파트 페이지
GDOC_GITUPIP_1_PAGES = set(range(133, 290))  # 기투입비용 전반부 페이지
GDOC_GITUPIP_2_PAGES = set(range(290, 446))  # 기투입비용 후반부 페이지

# 추출 대상 파일 목록 (출력파일명, 수신자료/ 내 상대경로, 추출방식, 비고)
TARGETS = [
    ("01_진행사항정리.txt", "폴더명/파일명.pdf", "pdf", "비고"),
    ...
]

# 페이지 필터 (감정보고서 등 대용량 PDF에서 필요한 범위만 추출)
FILE_PAGE_FILTERS = {
    "03_감정보고서_XXX": set(range(10, 91)),   # 제거할 페이지 범위
}
FILE_KEEP_FILTERS = {
    "03_감정보고서_원인분석부분": set(range(10, 91)),  # 보존할 페이지 범위
}
```

### 3. 수신자료 배치

`projects/새프로젝트명/수신자료/` 안에 TARGETS에 정의한 경로 구조대로 원본 파일을 넣습니다.

### 4. 참고예시 작성 (선택)

완성된 유사 보고서가 있다면 장별로 텍스트를 추출하여 `참고예시/` 에 저장합니다.
Claude가 글의 길이·문체를 이 예시에 맞춰 작성하므로, 있으면 결과물 품질이 높아집니다.

```
참고예시/참고_part1_개요_예시.txt      ← 1장 텍스트
참고예시/참고_part2_예시.txt           ← 2장 텍스트
참고예시/참고_part3_예시.txt           ← 3장 텍스트
참고예시/참고_part4_예시.txt           ← 4장 텍스트
```

### 5. 실행

```bash
python extractor.py 새프로젝트명
python generate_prompts_cause.py 새프로젝트명
```

---

## config.py 주요 설정 항목 설명

| 항목 | 설명 |
|------|------|
| `SOURCE_DIR` | 수신자료 폴더명 (기본값: `"수신자료"`) |
| `PROJECT_NAME` | 보고서 제목에 사용되는 공사명 |
| `PLAINTIFF` / `DEFENDANT` | 원고·피고명 |
| `CASE_NUMBER` | 사건번호 |
| `TARGETS` | 추출 대상 파일 목록. `(출력명, 원본 경로, 방식, 비고)` 튜플의 리스트 |
| `FILE_PAGE_FILTERS` | 지정 페이지를 **제거**하는 필터. 대용량 PDF에서 불필요한 부분 삭제 시 사용 |
| `FILE_KEEP_FILTERS` | 지정 페이지만 **보존**하는 필터. 같은 PDF에서 다른 범위를 별도 파일로 추출 시 사용 |
| `GDOC_*_PAGES` | 감정보고서 내 각 파트의 페이지 범위. 프롬프트 생성 시 해당 범위만 Claude에 전달 |

### 추출 방식 (`mode`) 종류

| 값 | 대상 형식 |
|----|----------|
| `"pdf"` | PDF (텍스트 레이어 있는 것) |
| `"xlsx"` | Excel |
| `"hwp"` | 한글(HWP) — 한글 프로그램 설치 필요 |

> 스캔 PDF(텍스트 레이어 없음)는 추출이 안 됩니다. 해당 파일은 수동으로 텍스트를 `processed/` 에 저장하거나 TARGETS에서 제외하세요.

---

## 의존성

```bash
pip install pymupdf openpyxl
# HWP 추출 시 추가로 필요:
pip install pywin32   # Windows 전용, 한글 프로그램 설치 필요
```
