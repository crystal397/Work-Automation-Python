# Python Work — 건설/부동산 업무 자동화 도구 모음

파이썬을 활용한 실무 밀착형 업무 자동화 및 데이터 분석 프로젝트입니다.
건설·부동산 분야의 반복 업무를 자동화하는 10개의 독립 모듈로 구성되어 있습니다.

---

## 폴더 구조

```
python work/
├── .env                                          # API 키 및 인증 정보
├── README.md
│
├── 01_web_automation/                            # 웹 자동화 (Selenium RPA)
├── 02_gis_visualization/                         # GIS 지오코딩 및 지도 시각화
├── 03_hwp_reporter/                              # 한글(HWP) 보고서 자동 작성
├── 04_pdf_reporter/                              # PDF 병합 및 갑지 자동 생성
├── 05_file_status_checker/                       # 파일 무결성 검사 및 보고서 생성
├── 06_classification_of_cost_item_groups/        # 산근 비목군 자동 분류
├── 07_lh-rental-price-matching/                  # LH 임대주택 실거래가 매칭
├── 08_molit_trade_collector/                     # 국토부 실거래가 대용량 수집
├── 09_pdf_invoice_extractor_and_excel_filler/    # 세금계산서 PDF 추출 → 엑셀 입력
└── 10_weather_collector/                         # 기상청 ASOS 기상 데이터 수집
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
MOLIT_SERVICE_KEY=국토부_API_키
JUSO_API_KEY=행안부_도로명주소_API_키
KAKAO_API_KEY=카카오_API_키
KMA_API_KEY=기상청_API_키
```

---

## 모듈별 요약

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
- PyInstaller로 EXE 빌드 가능 (`build.ps1`)

```bash
python pdf_merger.py [폴더경로] [출력파일.pdf]
```

**의존성**: reportlab, pypdf, pywin32

---

### 05. 파일 무결성 검사

폴더 내 파일을 재귀 탐색하여 상태를 검사하고 Excel 보고서 + 재송부 요청 문서 생성.

- Excel / Word / PDF / PPT / ZIP / 7Z / RAR 포맷별 검사
- 오류 분류: 빈 파일, 손상, 암호화, AIP/DRM
- 압축 파일 내부까지 재귀 검사, 인코딩 자동 감지 (UTF-8/CP949)
- Windows 260자 경로 제한 우회 (UNC 경로)

```bash
python check_files.py
```

**의존성**: openpyxl, python-docx, python-pptx, PyMuPDF, py7zr, rarfile

---

### 06. 산근 비목군 자동 분류

건설 산출내역서(산근)의 비목을 계약예규 제68조 기준으로 자동 분류.

- 분류 그룹: A(노무) / B(기계) / C(광산물) / D(제조) / E(공공요금) / F(농림) / G1~G5(표준시장) / Z(기타)
- 표준시장단가 9,607개 코드 매칭
- E그룹이 Z/G 충돌 시 우선 적용
- v7.0: 담당자 피드백 18개 규칙 반영

```bash
python item_group_auto_classification_v7.0.py
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

건설현장 좌표 기반으로 가장 가까운 ASOS 기상관측소를 자동 탐색하여 기상 데이터 수집.

- 전국 727개 ASOS 관측소 중 Haversine 거리 최근접 탐색
- 수집 항목: 기온(최고/최저), 강수량, 풍속(평균/최대), 신적설, 습도
- 공사중지 플래그 자동 계산 (강수 10mm+ / 풍속 14m/s+ / 적설 1cm+ / 폭염 35°C+ / 혹한 -10°C-)
- `config.py`의 `SITES`에 현장별 위도·경도·수집기간(`start`/`end`) 설정
- APScheduler로 매일 오전 6시 자동 수집
- SQLite / PostgreSQL 선택 지원

```bash
python collector.py    # 전체 기간 수집 (config.py의 SITES.start ~ end 기준)
python scheduler.py    # 매일 6시 자동 수집 데몬 실행
```

**의존성**: requests, sqlalchemy, apscheduler, pandas
**API**: 기상청 ASOS

---

## 주요 기능 요약

| 모듈 | 처리 규모 | 핵심 기술 |
|------|----------|----------|
| 02 GIS | 주소 20,000건+ 병렬 지오코딩 | ThreadPoolExecutor, Kakao API |
| 07 LH 매칭 | 195,238건 임대주택 매칭 | 4단계 신뢰도, SQLite 캐시 |
| 08 실거래가 | 68지역 × 4유형 × 10년 | 일일 한도 관리, 재개 기능 |
| 10 기상 | 727개 관측소 자동 탐색 | Haversine, APScheduler |

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
| 데이터베이스 | SQLite3, SQLAlchemy, PostgreSQL (선택) |
| 스케줄링 | APScheduler |
| GUI | Tkinter |
| 배포 | PyInstaller |
| 외부 API | 카카오 로컬, 국토부 실거래가, 행안부 도로명주소, 기상청 ASOS |

---

## 참고 사항

- **Windows 전용 모듈**: `01`, `03`, `04`는 pywin32 COM API 사용 (Windows만 동작)
- **장경로 지원**: `05`는 Windows 260자 경로 제한을 UNC 경로(`\\?\`)로 우회
- **API 일일 한도**: `08`은 국토부 API 일일 10,000건 한도를 유형별로 자동 관리
- **Tesseract**: `09` OCR 기능 사용 시 Tesseract 및 한국어 언어팩(`kor`) 별도 설치 필요
