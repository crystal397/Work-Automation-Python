# Work-Automation-Python
Python-based Business Process Automation (RPA) &amp; GIS Data Visualization Toolbox for Real-world Workflows
파이썬을 활용한 실무 밀착형 업무 자동화 및 데이터 분석 프로젝트입니다.

### 📂 주요 기능 (Key Features)
#### **1. 업무 프로세스 자동화 (RPA)**
* Selenium 기반 웹 자동화
  * 사내 업무 시스템 로그인 및 카테고리별 영수증 자동 업로드.
* 스마트 파일 처리
  * 9MB 초과 PDF 자동 분할 업로드 및 Excel → PDF 일괄 변환 파이프라인.<br>

#### **2. GIS 및 공간 데이터 시각화**
* 권역별 분포 분석
  * 주소 데이터를 좌표로 변환(Geocoding)하고 특정 반경 내 현황 시각화.
* 대화형 지도 생성
  * Folium 기반의 점묘도, 버블맵 등 3가지 케이스별 시각화 도구 구현.<br>

#### **3. 엑셀 데이터 한글 문서 자동화**
* 엑셀 데이터 수집 및 변환
  * 다중 시트 자동 파싱: 전체총괄표, 간접노무비 집계표, 퇴직금 등 3개 시트에서 데이터를 자동으로 읽어 구조화.
  * 값 포맷 자동 변환: 날짜, 천단위 금액, 백분율 등 셀 데이터 타입별로 한글 문서에 맞는 형식으로 자동 변환.
* 한글(HWP) 문서 자동화
  * 누름틀 기반 템플릿 입력: 미리 정의된 한글 누름틀(필드)에 엑셀 데이터를 자동 매핑하여 6개 표를 일괄 완성.
  * 표 커서 제어: 격자(행×열) 및 세로 단일열 표 입력을 공통 함수로 처리하여 셀 밀림 없이 정확하게 입력.
* 결과물 저장 및 내보내기
  * HWP·PDF 동시 저장: 작성 완료된 보고서를 편집용(.hwp)과 배포용(.pdf) 형식으로 동시에 자동 저장.
  * 날짜 기반 파일명 자동 생성: 실행일 기준으로 공기연장보고서_YYYYMMDD 형식의 파일명을 자동 부여.<br>

#### *🛠 Tech Stack*
* Language: Python 3.11.9
* Libraries: Selenium, Pandas, Folium, PyPDF, Requests, Tkinter, Openpyxl, pywin32
* API: Kakao Mobility, 공공데이터포털(LH)
