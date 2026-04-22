# 24. 돌관공사비 노무비 출력일보 자동 작성

소스 Excel 파일(업체별 공수 내역)을 읽어 **돌관공사비 산출근거 — 노무비 출력일보** 시트를 자동으로 생성하는 도구.

---

## 파일 구성

```
24_crash_construction/
├── mandays_report_automation_v11.py   ← 최신: PDF 인쇄 설정 완전 복원
├── mandays_report_automation_v9.py
├── mandays_report_automation_v8.py
├── mandays_report_automation_v7.py
├── mandays_report_automation_v6.py
├── .env.example                       ← 경로 설정 템플릿 (.env는 git 제외)
├── source/                            ← 소스 Excel 파일 (파일명에 연/월 포함)
└── output/                            ← 출력 파일 (git 제외)
```

---

## 사용 방법

```bash
pip install openpyxl holidays python-dotenv
```

`.env.example`을 복사해 `.env`로 저장하고 실제 경로 입력:

```env
CRASH_TEMPLATE_PATH=C:\...\sample_돌관공사비 산출근거.xlsx
CRASH_SOURCE_DIR=C:\...\source
CRASH_OUTPUT_PATH=C:\...\output\돌관공사비_산출근거_자동생성_v11.xlsx
```

실행:

```bash
python mandays_report_automation_v11.py
```

실행 후 Excel에서 **Ctrl+Alt+F9** (전체 재계산) 실행.

---

## 소스 파일 구조

- 탭 이름: `01.가시설`, `02.형틀목공` 등 `숫자.공종명` 형식
- B열 `작업자명` 헤더 아래 행: 작업자 이름
- C열: 일자(1~31), D열: 공수, F·H열: 카테고리 키워드 (`본선`/`복합`/`삼성` → 없으면 `기타`)
- B열 `일 기준` 아래 값: 노무비 단가 (AB열)
- 블록 하단 E열 마지막 숫자: 카테고리별 노무비 총액 (AC열, v9~)

---

## 출력 구조

- **카테고리 4종**: 본선 / 복합 / 삼성 / 기타
- **페이지 분할**: 한 페이지 최대 20명 + 합계행 (PDF 한 페이지 기준)
- **합계행**: 마지막 페이지에만 삽입, E열에 `N명` 자동 기입
- **날짜 헤더 색상**:
  - 토요일 → 파란 글씨
  - 일요일 · 공휴일 → 빨간 글씨 + `휴` 표시
- **연도별 분할**: `노무비 출력일보` 시트만 포함한 `_YYYY.xlsx` 별도 저장
- **PDF 인쇄 설정 (v11~)**: A4 가로, 배율 79%, 여백 자동 설정, 48행 단위 페이지 나눔

---

## 버전 변경 이력

| 버전 | 주요 변경 |
|------|----------|
| v3 | 케이씨산업 템플릿 적용, 시트 재생성 방식, read_only 소스 읽기 |
| v4 | 가변 길이 섹션 (헤더 6 + 근로자 N×2 + 합계 2), 토·일 요일 색상 |
| v5 | External link 제거, 처리 속도 최적화 |
| v6 | definedNames 고아 제거 (Excel 복구 경고 방지), ArrayFormula 재생성 |
| v7 | 페이지 분할 (최대 21명), AB열 단가·AC열 총액 소스에서 자동 추출 |
| v8 | `holidays` 라이브러리로 공휴일 자동 처리, 연도별 파일 분할 저장, 합계행 E열 `N명` 기입 |
| v9 | AC열 총액 계산 방식 변경 — 카테고리별 투입일 기준으로 분리 합산 |
| v10 | 템플릿 열 너비 유지, 마지막 페이지 20명 + 합계 구조로 변경 (PDF 한 페이지 맞춤) |
| v11 | PDF 인쇄 설정 완전 복원 (A4 가로·배율 79%·여백), 인쇄 영역 자동 설정, 48행 단위 페이지 나눔 |

---

## 환경변수 (`.env`)

| 키 | 설명 |
|----|------|
| `CRASH_TEMPLATE_PATH` | 템플릿 Excel 파일 전체 경로 |
| `CRASH_SOURCE_DIR` | 소스 파일 폴더 경로 |
| `CRASH_OUTPUT_PATH` | 출력 파일 경로 |

`.env` 미설정 시 스크립트와 같은 폴더의 기본값 사용.

---

## 의존성

```
openpyxl
holidays
python-dotenv
```
