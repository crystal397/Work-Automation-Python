# Python Work — 건설/부동산 업무 자동화 도구 모음

파이썬을 활용한 실무 밀착형 업무 자동화 및 데이터 분석 프로젝트입니다.
건설·부동산 분야의 반복 업무를 자동화하는 독립 모듈로 구성되어 있습니다.

---

## 폴더 구조

```
python work/
├── .env                                    # API 키 (git 제외)
├── .gitignore
├── README.md
│
├── 01_web_automation/
│   ├── README.md
│   └── web_automation.py
│
├── 02_gis_visualization/
│   ├── README.md
│   ├── spatial_analysis_tool.py
│   ├── vector_layers.py
│   └── zone_analysis_distribution.py
│
├── 03_hwp_reporter/
│   ├── README.md
│   ├── excel_reader.py
│   ├── hwp_writer.py
│   ├── main.py
│   ├── one_file/
│   │   └── excel_to_hwp.py
│   ├── requirements.txt
│   └── utils.py
│
├── 04_pdf_reporter/
│   ├── README.md
│   ├── build.ps1
│   ├── excel_to_pdf.py
│   ├── pdf_merger.py
│   ├── pdf_merger_essential_numbers.py
│   ├── pdf_merger_v2.py
│   ├── pdf_merger_v3.py
│   ├── pdf_merger_v4.py              ← 최신: 폴더에만 간지, 파일은 간지 없음
│   ├── pdf_merger_with_bookmark.py
│   └── pdf_size_splitter.py
│
├── 05_file_status_checker/
│   ├── README.md
│   ├── check_files.py
│   └── check_files_build.ps1
│
├── 06_classification_of_cost_item_groups/
│   ├── README.md
│   ├── build_exe.ps1
│   ├── item_group_auto_classification_v8.0.py
│   ├── item_group_auto_classification_v8.1.py
│   ├── item_group_auto_classification_v8.2.py
│   └── item_group_auto_classification_v8.3.py   ← 최신
│
├── 07_lh-rental-price-matching/
│   ├── README.md
│   └── lh_realestate_api.py
│
├── 08_molit_trade_collector/
│   ├── README.md
│   ├── bulk_collector.bat
│   ├── bulk_collector.py
│   ├── check_zero_records.py
│   └── export_report.py
│
├── 09_pdf_invoice_extractor_and_excel_filler/
│   ├── README.md
│   ├── build_exe.bat
│   └── pdf_invoice_extractor_and_excel_filler.py
│
├── 10_weather_collector/
│   ├── README.md
│   ├── analyzer.py
│   ├── build.bat
│   ├── build.py
│   ├── collector.py
│   ├── config.py
│   ├── flags.py                ← 작업불가일 플래그 정의 단일 소스
│   ├── gui.py                  ← GUI 주 실행 파일 (customtkinter)
│   ├── kma_client.py
│   ├── main.py
│   ├── scheduler.py
│   ├── sites.example.json      ← 현장 설정 템플릿 (sites.json은 git 제외)
│   ├── station_mapper.py
│   └── storage.py
│
├── 11_report_craft/
│   ├── README.md
│   ├── extractor.py
│   ├── generate_prompts_cause.py
│   └── setup_project.py
│
├── 12_manhour_aggregation/
│   ├── README.md
│   ├── .env.example
│   ├── aggregator.py
│   ├── filler.py
│   ├── formula_writer.py
│   ├── main.py
│   ├── manhour_aggregation.py
│   ├── pdf_to_excel.py
│   └── readers/
│       ├── __init__.py
│       ├── common.py
│       └── pdf_reader.py
│
├── 13_한글2024_macro/
│   ├── README.md
│   └── hwp_renumber.py
│
├── 18_report_craft/
│   ├── ARCHITECTURE.md
│   ├── .env.example
│   ├── config.py
│   ├── main.py
│   ├── report_template_guide.md
│   ├── requirements.txt
│   ├── 분석지시서.md
│   ├── 사용안내.txt
│   └── src/
│       ├── analyzer/
│       │   ├── analyzer.py
│       │   ├── data_checker.py
│       │   └── prompts.py
│       ├── calculator/
│       │   └── calculator.py
│       ├── extractor/
│       │   ├── file_classifier.py
│       │   ├── file_extractor.py
│       │   └── quality_checker.py
│       └── generator/
│           ├── build_templates.py
│           ├── docx_generator.py
│           ├── laws_db.py
│           └── md_generator.py
│
├── 20_report_craft_partially/
│   ├── config.py
│   ├── main.py
│   ├── gui.py                       ← GUI 진입점
│   ├── make_template.py
│   ├── make_template_B.py
│   ├── remember.md
│   ├── requirements.txt
│   ├── 귀책분석_패턴집.md
│   ├── 귀책분석_자동화.spec         ← PyInstaller 빌드 설정
│   ├── build.bat
│   ├── 사용안내.txt
│   └── src/
│       ├── correspondence_scanner.py
│       ├── prompt_builder.py
│       ├── reference_learner.py
│       ├── report_generator.py
│       └── text_extractor.py
│
├── 21_crawling_pages/
│   ├── README.md
│   ├── cak_crawler.py
│   ├── csv_to_xlsx.py
│   ├── debug_api.py
│   ├── kosca_crawler.py
│   └── land_price_lookup.py
│
├── 22_laws_import/
│   ├── README.md
│   ├── main.py              ← 진입점
│   ├── gui.py               ← CustomTkinter GUI
│   ├── engine.py            ← 5단계 시행일 판단 엔진
│   ├── api_client.py        ← 법제처 API 클라이언트
│   ├── cache.py             ← SQLite 캐시 (WAL 모드)
│   ├── report.py            ← Word 리포트 생성
│   ├── scraper.py           ← 행정규칙 연혁 웹 스크래퍼
│   ├── config.py            ← API 설정·대상 법령·키워드 23개
│   ├── requirements.txt
│   └── requirements.md      ← 요구사항 정의서
│
├── 23_excel_margins/
│   ├── adjust_excel_margins.py   ← 여백 조정 + GUI 진입점
│   ├── build.bat                 ← PyInstaller EXE 빌드
│   └── dist/
│       └── 엑셀_여백_조정.exe    ← 배포용 단일 EXE
│
├── 24_crash_construction/
│   ├── README.md
│   ├── .env.example                          ← 경로 설정 템플릿
│   ├── mandays_report_automation_v6.py
│   ├── mandays_report_automation_v7.py
│   ├── mandays_report_automation_v8.py
│   ├── mandays_report_automation_v9.py
│   ├── mandays_report_automation_v11.py  ← 최신: PDF 인쇄 설정 완전 복원
│   └── source/                           ← 소스 파일 (연/월 파일명)
│
└── 27_pic_excel/
    ├── README.md
    └── insert_photos.py                  ← 사진대지 엑셀 자동 삽입
```

---

## 설치 방법

```bash
python -m venv .venv
.venv\Scripts\activate  # Windows

pip install pandas openpyxl python-dotenv requests
```

모듈별 추가 패키지는 각 폴더의 README를 참고하세요.

### 환경변수 설정 (`.env`)

```env
# 루트 .env (공통)
MOLIT_SERVICE_KEY=국토부_API_키
JUSO_API_KEY=행안부_도로명주소_API_키
KAKAO_API_KEY=카카오_API_키
KMA_API_KEY=기상청_API_키
VWORLD_API_KEY=브이월드_API_키

# 12_manhour_aggregation/.env
PDF_ROOT=노무비자료_폴더_경로

# 18_report_craft/.env
REPORT_AUTHOR=보고서_작성_회사명
REPORT_INPUT_DIR=수신자료_폴더_경로   # 선택, 기본값: input/
REPORT_OUTPUT_DIR=결과물_폴더_경로    # 선택, 기본값: output/
```

각 폴더의 `.env.example`을 참고해 설정하세요.

---

## 모듈별 요약

### 00. Claude.ai 업무 결과물 모음

코드 없이 Claude.ai를 활용하여 수행한 업무 산출물을 보관하는 폴더.
보고서 초안, 분석 결과, 문서 작성 등 AI 협업 결과물을 프로젝트별로 정리.

| 폴더 | 업무 내용 |
|------|----------|
| `14_용역수행계획서` | 프로젝트 기자재 역무구분 기준서 및 용역수행계획서 작성 |
| `17_전시&체험` | 전시·체험 종합기획서 작성 (Claude.ai + Gemini 활용) |
| `19_공정팀_제안서` | 공정관리 용역 기술제안서 작성 |

---

### 01. 웹 자동화 (Selenium RPA)

사내 시스템 자동 로그인 및 문서 업로드 자동화.

- Excel → PDF 변환 후 카테고리별 자동 업로드 (9개 메뉴 분류)
- 9MB 초과 PDF 자동 분할
- `progress.json`으로 중단 후 재개
- 로그인 정보는 `.env`로 관리

```bash
python web_automation.py
```

**의존성**: selenium, pypdf, pywin32, python-dotenv

---

### 02. GIS 지오코딩 및 지도 시각화

20,000건 이상의 주소를 좌표로 변환하고 인터랙티브 지도 생성.

- 카카오 API 병렬 지오코딩 (15 스레드, 최대 10배 속도)
- 메모리 캐싱으로 중복 API 호출 방지
- 히트맵 / 마커클러스터 / 반경 분석 (250m·500m·1km)
- Tkinter GUI: 반경 구역 분류 + Excel 내보내기

```bash
python spatial_analysis_tool.py       # 지오코딩 + 지도 생성
python zone_analysis_distribution.py  # 구역 분석 GUI
```

**의존성**: folium, pandas, requests
**API**: 카카오 로컬 API

---

### 03. 한글(HWP) 보고서 자동 작성

Excel 데이터를 읽어 HWP 템플릿 6개 표를 자동으로 채웁니다.

- 다중 시트 파싱: 전체총괄표, 간접노무비 집계표, 퇴직금
- 날짜·금액·백분율 포맷 자동 변환
- HWP·PDF 동시 저장, 날짜 기반 파일명 자동 생성
- **Windows 전용** (HWP COM API)

```bash
python main.py
```

**의존성**: openpyxl, pywin32

---

### 04. PDF 병합 및 갑지 자동 생성

폴더 구조를 읽어 계층 번호 갑지를 자동 생성하고 PDF를 하나로 병합.

- 폴더 깊이에 따른 자동 번호 부여 (1., 1.1., 1.1.1.)
- Excel → PDF 일괄 변환 (Excel COM)
- 한글 폰트 폴백 (한컴바탕 → 맑은 고딕 → 나눔명조)
- **v4**: 폴더에만 간지 삽입, 파일(PDF/Excel)은 간지 없이 바로 추가
- **v3**: 폴더·파일 모두 간지 삽입, 초기 갑지 커스텀 + 시작 번호 지정
- PyInstaller로 EXE 빌드 가능 (`build.ps1`)

```bash
python pdf_merger_v4.py   # GUI 실행 (폴더에만 간지)
python pdf_merger_v3.py   # GUI 실행 (폴더·파일 모두 간지)
```

**의존성**: reportlab, pypdf, pywin32

---

### 05. 파일 무결성 검사

폴더 내 파일을 재귀 탐색하여 상태를 검사하고 Excel 보고서 생성.

- Excel / Word / PDF / PPT / ZIP / 7Z / RAR 포맷별 검사
- 오류 분류: 빈 파일, 손상, 암호화, AIP/DRM, IRM/DRM, 내용 깨짐
- `.doc`/`.ppt`: OLE2 구조 + 핵심 스트림 검증 + IRM/DRM 스트림 감지
- `.hwp`: `.pfile` RMS/AIP 감지, OLE 구조 + FileHeader + BodyText 검증, HWP ParaText 레코드 파싱으로 텍스트 깨짐 검사
- `.hwpx`: section XML 파싱 → 텍스트 추출 후 깨짐 비율 검사
- `.pdf`: 한국형 PDF DRM(Markany·SyncEZ 등) 1페이지 안내 패턴 감지
- `.docx`/`.pptx`: 텍스트 추출 후 깨짐 비율(15% 이상) 경고
- 압축 파일 내부까지 재귀 검사, 인코딩 자동 감지 (UTF-8/CP949)
- Windows 260자 경로 제한 우회 (UNC 경로)

```bash
python check_files.py
```

**의존성**: openpyxl, python-docx, python-pptx, PyMuPDF, olefile, py7zr, rarfile

---

### 06. 산근 비목군 자동 분류

건설 산출내역서(산근)의 비목을 계약예규 제68조 기준으로 자동 분류.

- 분류 그룹: A(노무) / B(기계) / C(광산물) / D(제조) / E(공공요금) / F(농림) / G1~G5(표준시장) / Z(기타)
- 표준시장단가 9,607개 코드 매칭
- E그룹이 Z/G 충돌 시 우선 적용
- 기본비목(D/A/B)은 미표기, 재분류 항목만 A열 표기
- N열에 검토사유 자동 기록
- v8.1: GZZZZZZ 분리행 정규화 (.M/.L/.E 자동 추가 + 노란색 표시)
- v8.2/v8.3: 분류 정확도 추가 개선

```bash
python item_group_auto_classification_v8.3.py
```

**의존성**: openpyxl

---

### 07. LH 임대주택 실거래가 매칭

195,238건의 LH 임대주택을 국토부 실거래가 API와 4단계 신뢰도로 매칭.

- 도로명 → 지번 변환: 행안부 API + 카카오 폴백
- 매칭 등급: G1(정밀) / G2(근접) / G3(참조) / G4(저신뢰)
- 주택 유형별 API 분기: 아파트/연립다세대/오피스텔/단독다가구
- SQLite + JSON 캐시로 API 비용 절감

```bash
python lh_realestate_api.py
```

**의존성**: pandas, requests, openpyxl
**API**: 국토부 실거래가, 행안부 도로명주소, 카카오 로컬

---

### 08. 국토부 실거래가 대용량 수집

서울 25구 + 경기 43 시군구 × 매매 4종 × 10년치 데이터를 SQLite에 수집.

- 수집 유형: 아파트·연립다세대·단독다가구·오피스텔 **매매**
- 일일 한도 관리: 유형별 10,000회 × 4종 = 40,000회
- `progress.json`으로 중단 후 재개
- 지수 백오프 재시도 (최대 3회)
- `export_report.py`: 연도별 Excel 보고서 (요약 + 2016~2026 시트)
- `check_zero_records.py`: 0건 항목 패턴 분석 및 API 재조회

```bash
python bulk_collector.py           # 수집 (이어하기 지원)
python bulk_collector.py --status  # 진행 현황 확인
python export_report.py            # Excel 보고서 생성
python check_zero_records.py --analyze --verify
```

**의존성**: requests, openpyxl
**API**: 국토부 실거래가

---

### 09. 세금계산서 PDF → 엑셀 자동 입력

PDF 세금계산서에서 데이터를 추출하여 Excel 템플릿에 자동 입력.

- 텍스트 추출 (pdfplumber) → OCR 폴백 (Tesseract)
- 승인번호 5가지 패턴 매칭 (국세청/파일명/일반/스캔)
- 사업자번호·발행일·공급가액 자동 추출
- Tkinter GUI (드래그앤드롭 지원)
- Tesseract 및 한국어 언어팩(`kor`) 별도 설치 필요

```bash
python pdf_invoice_extractor_and_excel_filler.py
```

**의존성**: pdfplumber, pypdf, openpyxl, pytesseract, pdf2image

---

### 10. 기상청 ASOS 기상 데이터 수집

건설현장 좌표 기반으로 가장 가까운 ASOS 기상관측소를 자동 탐색하여 기상 데이터 수집 및 공종별 작업불가일 산정. 비개발자도 `.exe`로 바로 사용 가능한 GUI 포함.

- 전국 727개 ASOS 관측소 중 Haversine 거리 최근접 탐색 + ASOS 유효성 자동 검증
- 수집 항목: 기온(최고/최저), 강수량, 풍속(평균/최대), 순간최대풍속, 신적설, 습도, 일조시간, 지면온도, 증발량
- 작업불가일 플래그 12종 (강수·강풍·크레인풍속·적설·폭염·한파·지면동결·일조부족·증발과다·강수유무·강설유무·안개)
- **공종별 판정 기준값 사용자 직접 설정** — 토공사·RC·타워크레인 등 6개 프리셋 또는 직접 입력, 기준값(예: 강수 5mm↑)도 공종마다 개별 조정 가능
- `sites.json`에 현장별 위도·경도·수집기간 및 공종별 작업 기간(`works`) 설정 (`sites.example.json` 참고)
- 공종별 작업불가일 산정 → 현장별 엑셀 출력 (`{site_id}_작업불가일.xlsx`), 요약·공종별 상세 시트 + 서식 자동 적용
- APScheduler로 매일 오전 6시 자동 수집
- SQLite 저장

```bash
python gui.py          # GUI 실행 (5단계 wizard, sites.json 일괄 수집 포함) ← 권장
python main.py         # CLI 대화형 실행
python collector.py    # 전체 기간 수집 (sites.json의 start ~ end 기준)
python analyzer.py     # 공종별 작업불가일 산정 및 엑셀 출력
python scheduler.py    # 매일 6시 자동 수집 데몬 실행
```

**의존성**: requests, sqlalchemy, apscheduler, pandas, openpyxl, customtkinter
**API**: 기상청 ASOS

---

### 11. 건설공사 분쟁 기술검토 보고서 작성 도구

건설공사 분쟁 관련 **원인·책임 분석** 및 **손실금액 적정성 검토** 보고서 작성을 위한 Claude.ai 프롬프트 자동 생성 파이프라인.

- 프로젝트(사건)별 폴더 구성 — 한 폴더 = 한 사건
- 수신자료(PDF·HWP·XLSX)에서 텍스트 추출 후 섹션별 프롬프트 자동 생성
- 보고서 유형별 템플릿 분리: 원인·책임 보고서 / 손실금액 보고서
- 참고 예시 문서 기반 Claude 문체 학습 지원
- **Windows 전용** (HWP COM API)

```bash
python extractor.py               # Step 1: 수신자료 텍스트 추출
python generate_prompts.py        # Step 2: 손실금액 보고서 프롬프트 생성
python generate_prompts_cause.py  # Step 2: 원인·책임 보고서 프롬프트 생성
```

**의존성**: python-docx, python-pptx, openpyxl, PyMuPDF, pywin32

---

### 12. 돌관공사비 노임 시트 공수 취합

업체별 일용노무비 자료를 취합하여 산출내역서 노임 시트에 자동 기입하는 도구.

- PDF 노임 지급 명세서 → Excel 자동 변환 (`pdf_to_excel.py`)
- 업체별 공수 집계 → 노임 시트 자동 기입 (`aggregator.py`, `filler.py`)
- 루트 폴더는 `.env`의 `PDF_ROOT` 또는 CLI 인자로 지정

```bash
python pdf_to_excel.py [폴더경로]   # PDF → Excel 변환
python main.py                       # 공수 취합 및 시트 기입
```

**의존성**: pdfplumber, openpyxl

---

### 13. 한글 2024 매크로 도구

한글 구버전 매크로 관련 유틸리티 및 문서 자동화 도구 모음.

#### hwp_renumber — 표/그림 번호 재정렬

HWPX 파일에서 뒤죽박죽된 `[표 N]` / `<그림 N>` 번호를 문서 순서대로 1부터 자동 재정렬.

- HWPX(ZIP+XML)을 직접 수정 — HWP COM Find/Replace 우회
  → 표 셀 안·글상자 안 등 위치 무관하게 동작
- HWP 자동번호 필드(`<hp:autoNum>`) 자동 감지 → 번호 충돌 방지
- Tkinter GUI: 파일 선택 다이얼로그 + 진행 로그 창
- 한글이 열려 있으면 파일 경로 자동 감지 및 저장
- 원본은 보존하고 `_renumbered.hwpx`로 별도 저장
- PyInstaller EXE 빌드 가능 (`hwp_renumber.spec`)

```bash
python hwp_renumber.py
# 또는
pyinstaller hwp_renumber.spec   # → dist\hwp_renumber.exe
```

**의존성**: pywin32, tkinter (내장)

---

### 18. 공기연장 간접비 보고서 자동 생성

수신자료(PDF·Excel·HWP·HTML·XML·TIF 등)를 넣으면 공기연장 간접비 산정 보고서(.docx)를 자동으로 생성.

- **멀티포맷 추출**: 포맷별 폴백 체인 (예: PDF → pdfplumber → pymupdf → OCR 순)
- **출처 태깅**: 모든 수치에 `파일명 | 페이지/시트` 출처 자동 기록
- **품질 검사**: OK / WARN / FAIL 등급 — FAIL도 중단 없이 끝까지 진행
- **tqdm 진행바**: 처리 중 남은 시간 표시, 중단 시 캐시로 이어서 재개
- **보고서 유형 자동 판별**: A(지방계약법) / B(국가계약법) / C(민간)
- **간접비 계산**: 실비/추정 분리, 산재·고용보험료, 일반관리비(한도 검증), 이윤(15% 한도)
- **데이터 충족도 검사**: 생성 전 필수/권장 항목 미흡 시 조치 방법 안내
- **Claude Code 연동**: API 미사용, Claude Code가 직접 파일 분석 → JSON 생성
- **Windows 전용** (HWP COM API, Tesseract OCR)

```bash
# Step 1: 수신자료 텍스트 추출
python main.py extract

# Step 2: Claude Code에서 분석 요청
# → "output/extracted_for_analysis.md 파일을 읽고 분석해서 output/analysis_result.json 으로 저장해줘"

# Step 3: 보고서 생성
python main.py generate
```

**출력**: `output/보고서_초안.md` + `output/보고서_초안.docx`

**의존성**: pdfplumber, PyMuPDF, openpyxl, pandas, pytesseract, python-docx, pywin32, tqdm

---

### 20. 귀책분석 파트 자동 생성

공문·변경계약서 등 수신자료를 스캔하여 건설공사 분쟁 보고서의 **귀책분석 파트** 초안을 자동 생성.

- **3-Pass 공문 필터링**: 폴더명 분류 → 공문 여부 확인 → 귀책 관련성 검사 (키워드 매칭)
- **borderline 자동 재분류**: 키워드 미매칭 공문을 제목·파일명 기반으로 재검사 — 관련 공문은 자동 승격, 무관 문서(식대·고지서 등)는 자동 제외
- **scan_no 추적 체인**: 스캔 번호가 프롬프트 → data.json → 최종 docx까지 유지되어 원본 파일 역추적 가능
- **validate 검증**: hard_error(생성 중단) / soft_warning(경고 후 진행) 이중 구조 — 소결 누락·일수 불일치·필수 필드 누락 자동 감지
- **14개 reference 패턴 학습**: 실제 완성 보고서에서 귀책분석 섹션 추출 → Claude 문체·논리 구조 학습
- **멀티포맷 추출**: docx / pdf / hwp / 이미지(OCR)
- **docx 출력**: `02_귀책분석_[프로젝트명]_[날짜].docx`
- **Claude Code 연동**: API 미사용 — Claude Code가 직접 파일을 읽고 data.json 생성
- **EXE 빌드**: PyInstaller (`귀책분석_자동화.spec`)

```bash
# 권장: scan + prepare 한 번에
python main.py scanprepare "C:\...\수신자료" --project 프로젝트명

# validate + generate 한 번에
python main.py finish 프로젝트명

# 품질 비교 (output vs reference)
python main.py compare-all

# 전체 프로젝트 일괄 재스캔
python main.py rescan-all
```

**의존성**: python-docx, pdfminer.six, PyMuPDF, pytesseract, tqdm

---

### 24. 돌관공사비 노무비 출력일보 자동 작성

소스 Excel 파일(업체별 공수 내역)을 읽어 돌관공사비 산출근거의 **노무비 출력일보** 시트를 자동 생성.

- 카테고리 4종 분류: 본선 / 복합 / 삼성 / 기타 (소스 F·H열 키워드 기반)
- 페이지 분할: 한 페이지 최대 20명 + 합계행 (PDF 한 페이지 기준, v10~)
- 합계행 E열에 실제 근로자 수 `N명` 기입, AC열 총액은 카테고리별 투입일 기준 분리 합산 (v9~)
- 날짜 헤더 색상: 토=파란 글씨 / 일·공휴일=빨간 글씨 + `휴` 표시
- `holidays` 라이브러리로 대체·선거·임시공휴일까지 자동 처리
- 통합 파일 1개 + 연도별 분할 파일 (`_YYYY.xlsx`) 동시 저장
- PDF 인쇄 설정 자동 적용 (A4 가로·배율 79%·여백·48행 단위 페이지 나눔, v11~)
- 경로 설정은 `.env`로 분리 관리 (민감 정보 코드 제외)

```bash
pip install openpyxl holidays python-dotenv
python mandays_report_automation_v11.py
```

**의존성**: openpyxl, holidays, python-dotenv

---

### 23. 엑셀 인쇄 여백 일괄 조정

폴더 내 모든 엑셀 파일의 시트 여백을 일괄 검사·수정하는 도구.

- 좌+우 여백 합이 지정값(기본 2cm)인 시트를 자동 탐지
- 좌 1cm / 우 1cm로 재설정 후 저장
- `.xlsx` / `.xlsm` 대상, 임시 파일(`~$`) 자동 제외
- **Tkinter GUI**: 폴더 선택 → 실행 → 결과 로그 표시
- CLI 모드 겸용: `exe "경로"` 인자 전달 시 터미널 실행
- PyInstaller 단일 EXE 빌드 (`build.bat`)

```bash
python adjust_excel_margins.py          # GUI 실행
python adjust_excel_margins.py "경로"   # CLI 실행
```

**의존성**: openpyxl, tkinter (내장)

---

### 22. 입찰공고일 기준 법령 자동 매칭

공기연장 보고서 작성 시, 입찰공고일 시점에 실제로 시행 중이던 법령 버전을 자동 매칭하고 관련 조문 원문을 추출.

- **5단계 판단 로직**: 연혁조회 → 1차 후보 → 부칙 경과규정 탐지(전체/조문 단위 구분) → 2차 후보 → 사용자 검토
- **14개 법령 대상**: 국가·지방계약법(법률·시행령·시행규칙), 조달사업법, 하도급법, 공사계약일반조건 외 회계예규·고시
- **할루시네이션 방지**: 법제처 API 원문만 인용, AI 재작성 없음, 공포번호·시행일 5요소 추적
- **공기연장 키워드 23개**: 공기연장·간접비·지체상금·설계변경·귀책사유 등으로 관련 조문 자동 필터
- 행정규칙 연혁 3단계 fallback 스크래핑 (requests/BS4 → search_law)
- SQLite 캐시(7일 TTL, WAL 모드), Word 리포트 출력, 감사 로그 기록
- API 키는 `.env`로 관리 (`LAW_API_OC`)

```bash
python main.py   # GUI 실행
```

**의존성**: requests, xmltodict, python-docx, python-dotenv, customtkinter, tkcalendar, beautifulsoup4
**API**: 법제처 국가법령정보 공동활용 (open.law.go.kr — 사전 승인 필요)

---

### 27. 사진대지 엑셀 자동 삽입

사진 폴더의 이미지를 엑셀 사진대지 템플릿의 병합 셀 슬롯에 자동으로 삽입하는 CLI 도구.

- 엑셀 시트에서 큰 병합 셀 영역을 사진 슬롯으로 자동 탐지
- 파일명 오름차순으로 정렬 후 슬롯에 순서대로 삽입
- 비율 유지 + 슬롯 중앙 정렬, 경계 이탈 없음
- 스마트폰 EXIF 회전 정보 자동 반영
- 출력 파일은 원본과 별도 저장 (`원본명_완성.xlsx`)

```bash
pip install openpyxl pillow
python insert_photos.py 사진대지.xlsx ./photos/
python insert_photos.py 사진대지.xlsx ./photos/ 결과물.xlsx
```

**의존성**: openpyxl, pillow

---

### 21. 건설·부동산 관련 웹 크롤링 도구

건설협회·감정평가 관련 사이트 크롤링 및 공시지가 조회 도구 모음.

- `cak_crawler.py`: 대한건설협회 건설업체 정보 수집
- `kosca_crawler.py`: 한국건설감정원 데이터 수집
- `land_price_lookup.py`: V-World API 기반 공시지가 일괄 조회
- `csv_to_xlsx.py`: 수집 결과 CSV → Excel 변환

```bash
python cak_crawler.py       # 건설협회 업체 정보 수집
python land_price_lookup.py # 공시지가 조회
```

**API**: V-World 공간정보 오픈플랫폼

---

## 주요 기능 요약

| 모듈 | 처리 규모 | 핵심 기술 |
|------|----------|----------|
| 02 GIS | 주소 20,000건+ 병렬 지오코딩 | ThreadPoolExecutor, Kakao API |
| 07 LH 매칭 | 195,238건 임대주택 매칭 | 4단계 신뢰도, SQLite 캐시 |
| 08 실거래가 | 68지역 × 4유형 × 10년 | 일일 한도 관리, 재개 기능 |
| 10 기상 | 727개 관측소 자동 탐색 + 공종별 작업불가일 산정 | Haversine, APScheduler |
| 18 보고서 | PDF·HWP·Excel 등 멀티포맷 → Word 보고서 | 폴백 체인, Claude Code 연동 |
| 20 귀책분석 | 공문 3-Pass 필터링 + borderline 자동 재분류 → 귀책분석 docx 자동 생성 | scan_no 추적 체인, validate 이중 검증, 14개 reference 패턴 |
| 24 돌관공사비 | 소스 Excel → 노무비 출력일보 자동 생성 (통합 + 연도별 분할) | openpyxl 셀 스냅샷·붙여넣기, 공휴일 색상, 페이지 분할 |
| 27 사진대지 | 사진 폴더 → 엑셀 병합 셀 슬롯 자동 삽입 | openpyxl TwoCellAnchor, Pillow EXIF 보정 |

### 공통 설계 패턴

- **API 키 보안**: `.env` + python-dotenv
- **중단 후 재개**: JSON/SQLite 기반 진행 상태 저장
- **오류 복원**: 지수 백오프 재시도
- **캐싱**: 메모리 딕셔너리 + SQLite + JSON으로 중복 API 호출 방지
- **GUI**: Tkinter 공통 파일 선택 다이얼로그
- **EXE 빌드**: PyInstaller (`build.ps1` / `build_exe.bat`)

---

## 기술 스택

| 분류 | 라이브러리 |
|------|-----------|
| 웹 자동화 | Selenium, pywin32 |
| 데이터 처리 | Pandas, openpyxl |
| 지도/GIS | Folium, Requests, ThreadPoolExecutor |
| 문서 처리 | python-docx, python-pptx, PyMuPDF, reportlab, pypdf, pdfplumber |
| OCR | pytesseract, pdf2image |
| 데이터베이스 | SQLite3, SQLAlchemy |
| 스케줄링 | APScheduler |
| GUI | Tkinter, customtkinter |
| 배포 | PyInstaller |
| 진행 표시 | tqdm |
| 외부 API | 카카오 로컬, 국토부 실거래가, 행안부 도로명주소, 기상청 ASOS |
| 공휴일 | holidays (한국 공휴일·대체공휴일·선거일 자동 처리) |

---

## 참고 사항

- **Windows 전용 모듈**: `01`, `03`, `04`, `11`, `18`은 pywin32 COM API 사용 (Windows만 동작)
- **장경로 지원**: `05`는 Windows 260자 경로 제한을 UNC 경로(`\\?\`)로 우회
- **API 일일 한도**: `08`은 국토부 API 일일 10,000건 한도를 유형별로 자동 관리
- **Tesseract**: `09`, `18`, `20` OCR 기능 사용 시 Tesseract 및 한국어 언어팩(`kor`) 별도 설치 필요
- **Claude Code 연동**: `18`, `20`은 Anthropic API 미사용 — Claude Code CLI가 직접 파일을 읽고 분석·JSON 생성
- **현장 정보 보안**: `10`은 현장명·위경도를 `sites.json`으로 분리 관리 (git 비공개)
- **V-World API**: `21`의 공시지가 조회 기능 사용 시 V-World API 키 필요 (vworld.kr 발급)
