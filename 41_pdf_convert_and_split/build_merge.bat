@echo off
REM ============================================================
REM  convert_and_merge.exe 빌드 스크립트
REM ============================================================

setlocal
cd /d "%~dp0"

echo.
echo [1/3] 필수 패키지 설치 중...
echo ------------------------------------------------------------
python -m pip install --upgrade pip
python -m pip install pyinstaller pywin32 pypdf pillow reportlab tkinterdnd2 simple-hwp2pdf
if errorlevel 1 (
    echo.
    echo [오류] 패키지 설치에 실패했습니다. Python 설치를 확인하세요.
    pause
    exit /b 1
)

echo.
echo [2/3] 기존 빌드 산출물 정리...
echo ------------------------------------------------------------
if exist build rmdir /s /q build
if exist dist rmdir /s /q dist
if exist convert_and_merge.spec del /q convert_and_merge.spec

echo.
echo [3/3] PyInstaller로 단일 EXE 빌드 중...
echo ------------------------------------------------------------
python -m PyInstaller ^
    --onefile ^
    --windowed ^
    --name "convert_and_merge" ^
    --collect-all tkinterdnd2 ^
    --hidden-import win32com.client ^
    --hidden-import win32timezone ^
    --hidden-import pypdf ^
    --hidden-import PIL ^
    --hidden-import reportlab ^
    --hidden-import reportlab.pdfbase ^
    --hidden-import reportlab.pdfbase.ttfonts ^
    --hidden-import simple_hwp2pdf ^
    convert_and_merge.py

if errorlevel 1 (
    echo.
    echo [오류] PyInstaller 빌드에 실패했습니다.
    pause
    exit /b 1
)

echo.
echo ============================================================
echo  완료!  결과물:  dist\convert_and_merge.exe
echo ============================================================
pause
endlocal