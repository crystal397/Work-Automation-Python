@echo off
REM ============================================================
REM  탐색기 우클릭 메뉴에서 제거
REM ============================================================

setlocal

echo.
echo 우클릭 메뉴에서 "PDF 변환 + 200MB 분할" 항목을 제거합니다.
echo.
choice /C YN /M "계속할까요"
if errorlevel 2 (
    echo 취소되었습니다.
    pause
    exit /b 0
)

reg delete "HKCU\Software\Classes\Directory\shell\ConvertAndSplitPDF" /f >nul 2>&1
reg delete "HKCU\Software\Classes\Directory\Background\shell\ConvertAndSplitPDF" /f >nul 2>&1

echo.
echo 제거되었습니다.
pause
endlocal