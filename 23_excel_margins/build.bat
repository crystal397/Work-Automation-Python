@echo off
chcp 65001 > nul
echo ===================================
echo  엑셀 여백 조정 EXE 빌드
echo ===================================

pyinstaller ^
  --onefile ^
  --windowed ^
  --name "엑셀_여백_조정" ^
  --hidden-import openpyxl ^
  adjust_excel_margins.py

echo.
if exist "dist\엑셀_여백_조정.exe" (
    echo [완료] dist\엑셀_여백_조정.exe 생성됨
) else (
    echo [오류] 빌드 실패
)
pause
