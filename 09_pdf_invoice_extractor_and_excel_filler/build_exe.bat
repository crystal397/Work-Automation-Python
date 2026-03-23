@echo off
chcp 65001 >nul
echo ==========================================
echo  세금계산서 추출기 - exe 빌드
echo ==========================================
echo.

REM 1) 필수 패키지 설치
echo [1/2] 필수 패키지 설치 중...
pip install pypdf pdfplumber openpyxl pyinstaller
echo.

REM 2) exe 빌드
echo [2/2] exe 빌드 중... (1~2분 소요)
pyinstaller --onefile --windowed --name "세금계산서추출기" pdf_invoice_extractor_and_excel_filler.py
echo.

echo ==========================================
echo  빌드 완료!
echo  dist\세금계산서추출기.exe 파일을 확인하세요.
echo ==========================================
echo.

REM dist 폴더 열기
explorer dist
pause