# Work-Automation-Python
Python-based Business Process Automation (RPA) &amp; GIS Data Visualization Toolbox for Real-world Workflows
파이썬을 활용한 실무 밀착형 업무 자동화 및 데이터 분석 프로젝트입니다.

📂 주요 기능 (Key Features)
1. 업무 프로세스 자동화 (RPA)
* Selenium 기반 웹 자동화: 사내 업무 시스템 로그인 및 카테고리별 영수증 자동 업로드.
* 스마트 파일 처리: 9MB 초과 PDF 자동 분할 업로드 및 Excel → PDF 일괄 변환 파이프라인.
*관련 파일:* 

2. GIS 및 공간 데이터 시각화
* 권역별 분포 분석: 주소 데이터를 좌표로 변환(Geocoding)하고 특정 반경 내 현황 시각화.
* 대화형 지도 생성: Folium 기반의 점묘도, 버블맵 등 3가지 케이스별 시각화 도구 구현.
*관련 파일:* 

3. 데이터 ETL 및 파일 관리
* 공공데이터 API 연동: LH 실거래가 및 건축물대장 API 자동 수집 도구.
* 규칙 기반 파일 정렬: 난잡한 파일명을 '날짜_내용_금액' 규칙으로 추출 및 엑셀화.
*관련 파일:* 

🛠 Tech Stack
* Language: Python 3.11.9
* Libraries: Selenium, Pandas, Folium, PyPDF, Requests, Tkinter, Openpyxl
* API: Kakao Mobility, 공공데이터포털(LH)
