# ClaimCraft — 건설공사 분쟁 보고서 작성 도구

건설공사 분쟁 관련 기술검토 보고서(원인·책임 분석 / 손실금액 적정성 검토) 작성을 위한
Claude.ai 프롬프트 자동 생성 파이프라인입니다.

---

## 디렉토리 구조

```
11_claim_craft/
│
├── projects/                        # 프로젝트별 폴더 (한 폴더 = 한 사건)
│   └── 창원용원/                    # 프로젝트 예시
│       ├── config.py                # 메타데이터, TARGETS, 페이지 범위
│       ├── sections_cause.py        # 원인·책임 보고서 섹션 정의
│       ├── sections.py              # 손실금액 보고서 섹션 정의
│       ├── 수신자료/                 # 원본 자료 (PDF, HWP, XLSX 등)
│       ├── processed/               # 추출된 텍스트 (extractor.py 출력)
│       ├── 참고예시/                 # 완성 보고서 예시 (Claude 문체 학습용)
│       │   ├── 참고_part1_개요_예시.txt
│       │   ├── 참고_part2_예시.txt
│       │   ├── 참고_part3_예시.txt
│       │   └── 참고_part4_예시.txt
│       ├── prompts_cause/           # 출력 — 원인·책임 보고서 프롬프트
│       └── prompts/                 # 출력 — 손실금액 보고서 프롬프트
│
├── templates/                       # 공통 템플릿 (모든 프로젝트에서 공유)
│   ├── prompt_cause.txt             # 원인·책임 보고서 역할·규칙·목차 템플릿
│   └── prompt.txt                   # 손실금액 보고서 역할·규칙·목차 템플릿
│
├── 법령_processed/                  # 공통 법령 텍스트 (모든 프로젝트에서 공유)
│
├── extractor.py                     # 수신자료 → processed/ 텍스트 추출
├── generate_prompts_cause.py        # 원인·책임 보고서 프롬프트 생성
├── generate_prompts.py              # 손실금액 보고서 프롬프트 생성
└── setup_project.py                 # 새 프로젝트 초기화
```

---

## 각 파일의 역할

### 스크립트 (루트)

| 파일 | 역할 | 사용법 |
|------|------|--------|
| `setup_project.py` | 새 프로젝트 폴더 생성, 수신자료 스캔, config.py·sections 템플릿 자동 생성 | `python setup_project.py 프로젝트명` |
| `extractor.py` | 수신자료(PDF·HWP·XLSX) → 텍스트 추출 → processed/ 저장 | `python extractor.py 프로젝트명` |
| `generate_prompts_cause.py` | templates/prompt_cause.txt + sections_cause.py + processed/ 파일 조합 → prompts_cause/ 저장 | `python generate_prompts_cause.py 프로젝트명` |
| `generate_prompts.py` | templates/prompt.txt + sections.py + processed/ 파일 조합 → prompts/ 저장 | `python generate_prompts.py 프로젝트명` |

> 프로젝트명 생략 시 기본값 `창원용원` 사용.
> 세 스크립트 모두 내부에 프로젝트 고유 내용이 없는 **완전 generic** 코드입니다.

---

### 프로젝트 폴더 내 파일

#### `config.py` — 프로젝트 설정

| 항목 | 설명 |
|------|------|
| `SOURCE_DIR` | 수신자료 폴더명 (기본값: `"수신자료"`) |
| `PROJECT_NAME`, `PLAINTIFF`, `DEFENDANT`, `CASE_NUMBER` | 보고서 헤더 등에 사용하는 메타데이터 |
| `GDOC_*_PAGES` | 감정보고서 내 각 파트의 페이지 범위. 프롬프트 생성 시 해당 범위만 Claude에 전달 |
| `TARGETS` | 추출 대상 파일 목록. `(출력파일명, 수신자료/ 내 상대경로, 추출방식, 비고)` 튜플의 리스트. `setup_project.py --rescan`으로 자동 갱신 가능 |
| `FILE_PAGE_FILTERS` | 지정 페이지를 **제거**하는 필터 (대용량 PDF에서 불필요한 부분 삭제) |
| `FILE_KEEP_FILTERS` | 지정 페이지만 **보존**하는 필터 (같은 PDF를 다른 범위로 별도 추출) |

#### `sections_cause.py` / `sections.py` — 섹션 정의

각 보고서를 몇 개의 파트로 나누어, 파트별로 아래 3가지를 정의합니다.

```python
SECTIONS = [
    {
        "filename":       "cause_part1_개요.txt",   # 출력 파일명
        "title":          "Part 1 — 1장 개요",       # 섹션 제목
        "output_request": "...",                     # Claude에게 전달하는 출력 지침
        "files": [                                   # 이 파트에 포함할 자료 목록
            ("## 라벨", "processed:파일명.txt"),
            ("## 라벨", ("filtered", "소스.txt", "GDOC_31_32_PAGES")),
        ],
    },
    ...
]
```

**파일 참조 형식:**

| 형식 | 의미 |
|------|------|
| `"processed:파일명.txt"` | `processed/` 폴더의 추출 텍스트 |
| `"예시:파일명.txt"` | `참고예시/` 폴더의 예시 파일 |
| `"법령:파일명.txt"` | 루트 `법령_processed/` 의 공통 법령 텍스트 |
| `("filtered", "소스.txt", "CONFIG_ATTR")` | processed/ 파일을 config의 페이지 범위로 필터 후 반환 |

#### `참고예시/` — Claude 문체 학습용 예시

완성된 보고서의 장(chapter)별 텍스트를 저장합니다.
각 섹션의 `files` 에 `"예시:참고_partN_예시.txt"` 형식으로 포함하면,
Claude가 이 예시의 **글의 길이, 서술 방식, 법령 인용 방식**을 참고하여 작성합니다.

---

### 공통 리소스 (루트)

| 폴더/파일 | 내용 | 특징 |
|-----------|------|------|
| `templates/prompt_cause.txt` | 원인·책임 보고서 역할·규칙·목차 전체 | 모든 프로젝트 공유. 수정 시 전 프로젝트에 반영 |
| `templates/prompt.txt` | 손실금액 보고서 역할·규칙·목차 전체 | 동일 |
| `법령_processed/` | 민법, 국가계약법, 건설산업기본법 등 텍스트 | 모든 프로젝트 공유 |

---

## 전체 작업 흐름

```
수신자료/ (PDF·HWP·XLSX 원본)
    ↓  python extractor.py 프로젝트명
processed/ (장별 텍스트 파일)
    ↓  python generate_prompts_cause.py 프로젝트명
prompts_cause/ (파트별 프롬프트 파일 5개)
    ↓  Claude.ai 대화창에 파트 순서대로 붙여넣기
보고서 초안 (Word 등으로 정리)
```

---

## 새 프로젝트 시작 순서

### 1단계 — 프로젝트 초기화

```bash
python setup_project.py 새프로젝트명
# 다른 프로젝트를 기준으로 복사하려면:
python setup_project.py 새프로젝트명 --source 다른프로젝트명
```

자동으로 생성/복사되는 것:

| 항목 | 내용 |
|------|------|
| 폴더 구조 | `수신자료/`, `processed/`, `참고예시/`, `prompts_cause/`, `prompts/` |
| `config.py` | 메타데이터 TODO 항목 포함 빈 설정 파일 (TARGETS는 --rescan으로 채움) |
| `sections_cause.py` | 창원용원 기준으로 복사 + 상단 수정 안내 주석 추가 |
| `sections.py` | 동일 |
| `참고예시/*.txt` | 창원용원 예시 파일 4개 복사 (이 사건 예시 생기면 교체) |

> 기본 복사 기준은 `창원용원`입니다. `--source` 옵션으로 다른 프로젝트를 지정할 수 있습니다.

### 2단계 — 수신자료 배치 후 TARGETS 자동 스캔

원본 파일(PDF·HWP·XLSX)을 `projects/새프로젝트명/수신자료/` 에 넣은 후:

```bash
python setup_project.py 새프로젝트명 --rescan
```

`수신자료/` 를 재귀 스캔하여 `config.py` 의 `TARGETS` 목록을 자동 갱신합니다.
출력 파일명은 `01_파일명.txt`, `02_파일명.txt` 형식으로 자동 부여됩니다.
필요 시 직접 수정하여 파일명·순서를 정리하세요.

### 3단계 — 텍스트 추출

```bash
python extractor.py 새프로젝트명
```

`수신자료/` 의 PDF·HWP·XLSX를 읽어 `processed/` 에 텍스트 파일로 저장합니다.

### 4단계 — config.py + sections_cause.py 작성 (Claude Code 사용)

extractor.py 실행 후, Claude Code 대화창에서 아래처럼 요청합니다:

```
새프로젝트명 프로젝트의 config.py와 sections_cause.py를 작성해줘.
processed/ 파일들을 읽고 사건 정보(공사명, 원고, 피고, 사건번호 등)와
감정보고서 페이지 범위를 파악해서 config.py를 채우고,
이 사건의 쟁점에 맞게 sections_cause.py도 창원용원 형식으로 만들어줘.
```

Claude Code가 `processed/` 파일들을 직접 읽고 사건 내용을 파악하여
`config.py` 메타데이터·페이지 범위와 `sections_cause.py` 를 작성·저장합니다.

```python
SECTIONS = [
    {
        "filename": "cause_part1_개요.txt",
        "title":    "Part 1 — 1장 개요",
        "output_request": """
            위 공사 정보를 바탕으로 1장 개요를 작성하십시오.
            ...
        """,
        "files": [
            ("## [참고 예시]", "예시:참고_part1_개요_예시.txt"),
            ("## 사업 진행사항", "processed:01_진행사항정리.txt"),
            ("## 국가계약법",   "법령:국가를 당사자로 하는 계약에 관한 법률(법률)...txt"),
            ("## 감정보고서 책임비율 파트", ("filtered", "03_감정보고서.txt", "GDOC_31_32_PAGES")),
        ],
    },
    ...
]
```

### 5단계 — 참고예시 교체 (선택, 권장)

1단계에서 창원용원 예시 4개가 이미 복사되어 있습니다.
이 사건과 유사한 완성 보고서가 있다면 장별로 텍스트를 저장하여 덮어씁니다.
없으면 창원용원 예시가 기본값으로 사용됩니다.

```
projects/새프로젝트명/참고예시/
    참고_part1_개요_예시.txt    ← 1장 텍스트
    참고_part2_예시.txt         ← 2장 텍스트
    참고_part3_예시.txt         ← 3장 텍스트
    참고_part4_예시.txt         ← 4장 텍스트
```

### 6단계 — 프롬프트 생성

```bash
# 원인·책임 분석 보고서 프롬프트 생성
python generate_prompts_cause.py 새프로젝트명

# 손실금액 적정성 검토 보고서 프롬프트 생성 (필요 시)
python generate_prompts.py 새프로젝트명
```

### 7단계 — Claude.ai에서 실행

`projects/새프로젝트명/prompts_cause/` 의 파일을 파트 순서대로 Claude.ai에 붙여넣습니다.

| 파일 | 보고서 목차 |
|------|------------|
| `cause_part1_개요.txt` | 1장 과업 개요 및 사업 개요 |
| `cause_part2_사실관계_발주자귀책.txt` | 2.1 사실관계 + 2.2 발주자 귀책 |
| `cause_part3a_시공사귀책_전반.txt` | 2.3.1~2.3.5 시공사 귀책 전반 |
| `cause_part3b_시공사귀책_후반.txt` | 2.3.6~2.3.10 시공사 귀책 후반 |
| `cause_part4_손실금액_결론.txt` | 3장 손실금액 + 4장 결론 |

---

## 추출 방식(mode) 종류

| 값 | 대상 형식 | 비고 |
|----|----------|------|
| `"pdf"` | PDF | 텍스트 레이어가 있는 PDF만 가능 |
| `"xlsx"` | Excel (xlsx, xls) | |
| `"hwp"` | 한글(HWP, HWPX) | 한글 프로그램 + pywin32 설치 필요 |

> **스캔 PDF** (텍스트 레이어 없음)는 자동 추출이 안 됩니다.
> 해당 파일은 TARGETS에서 제외하거나 텍스트를 수동으로 `processed/` 에 저장하세요.

---

## 의존성

```bash
pip install pymupdf openpyxl

# HWP 추출 시 추가 (Windows 전용, 한글 프로그램 설치 필요)
pip install pywin32
```
