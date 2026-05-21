@echo off
REM ============================================================
REM  탐색기 우클릭 메뉴에 등록 (현재 사용자만, 관리자 권한 불필요)
REM ============================================================
REM  같은 폴더의 convert_and_split.exe를 우클릭 메뉴에 추가합니다.
REM  관리자 권한 없이 HKEY_CURRENT_USER 에 등록되므로 안전합니다.
REM ============================================================

setlocal enabledelayedexpansion
cd /d "%~dp0"

set "EXE_PATH=%~dp0convert_and_split.exe"

if not exist "%EXE_PATH%" (
    echo.
    echo [오류] convert_and_split.exe를 찾을 수 없습니다.
    echo        이 파일을 EXE와 같은 폴더에 두고 다시 실행하세요.
    echo        현재 찾는 경로: %EXE_PATH%
    pause
    exit /b 1
)

REM 레지스트리 경로의 백슬래시는 .reg에서 두 번씩 써야 하므로 변환
set "EXE_REG=%EXE_PATH:\=\\%"

echo.
echo 다음 경로의 EXE를 우클릭 메뉴에 등록합니다:
echo   %EXE_PATH%
echo.
choice /C YN /M "계속할까요"
if errorlevel 2 (
    echo 취소되었습니다.
    pause
    exit /b 0
)

REM 1) 폴더 자체를 우클릭했을 때
reg add "HKCU\Software\Classes\Directory\shell\ConvertAndSplitPDF" /ve /d "PDF 변환 + 200MB 분할" /f >nul
reg add "HKCU\Software\Classes\Directory\shell\ConvertAndSplitPDF" /v "Icon" /d "\"%EXE_PATH%\"" /f >nul
reg add "HKCU\Software\Classes\Directory\shell\ConvertAndSplitPDF\command" /ve /d "\"%EXE_PATH%\" \"%%1\"" /f >nul

REM 2) 폴더 안 빈 공간을 우클릭했을 때 (현재 폴더 대상)
reg add "HKCU\Software\Classes\Directory\Background\shell\ConvertAndSplitPDF" /ve /d "PDF 변환 + 200MB 분할" /f >nul
reg add "HKCU\Software\Classes\Directory\Background\shell\ConvertAndSplitPDF" /v "Icon" /d "\"%EXE_PATH%\"" /f >nul
reg add "HKCU\Software\Classes\Directory\Background\shell\ConvertAndSplitPDF\command" /ve /d "\"%EXE_PATH%\" \"%%V\"" /f >nul

echo.
echo ============================================================
echo  등록 완료!
echo ============================================================
echo  탐색기에서 폴더를 우클릭하면
echo  "PDF 변환 + 200MB 분할" 메뉴가 보입니다.
echo.
echo  Windows 11에서는 우클릭 후 [추가 옵션 표시] 를 눌러야
echo  나타날 수 있습니다 (Shift+우클릭으로도 가능).
echo.
echo  제거하려면 unregister_context_menu.bat 를 실행하세요.
echo ============================================================
pause
endlocal