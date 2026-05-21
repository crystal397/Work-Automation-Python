@echo off
REM ============================================================
REM  convert_and_split.exe 빌드 스크립트
REM ============================================================
REM  요구사항: Python 3.10+ 가 PATH에 있어야 함
REM  실행 위치: 이 build.bat가 있는 폴더에서 더블클릭 또는 cmd 실행
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
if exist convert_and_split.spec del /q convert_and_split.spec

echo.
echo [3/3] PyInstaller로 단일 EXE 빌드 중...
echo ------------------------------------------------------------
REM --windowed:  실행 시 검은 콘솔 창을 띄우지 않음 (GUI 앱)
REM --onefile:   단일 EXE로 패키징
REM --name:      산출물 이름
REM --collect-all tkinterdnd2: tkinterdnd2의 DLL/리소스 포함
REM --hidden-import: PyInstaller가 자동 탐지 못 하는 모듈 명시
python -m PyInstaller ^
    --onefile ^
    --windowed ^
    --name "convert_and_split" ^
    --collect-all tkinterdnd2 ^
    --hidden-import win32com.client ^
    --hidden-import win32timezone ^
    --hidden-import pypdf ^
    --hidden-import PIL ^
    --hidden-import reportlab ^
    --hidden-import reportlab.pdfbase ^
    --hidden-import reportlab.pdfbase.ttfonts ^
    --hidden-import simple_hwp2pdf ^
    convert_and_split.py

if errorlevel 1 (
    echo.
    echo [오류] PyInstaller 빌드에 실패했습니다.
    pause
    exit /b 1
)

echo.
echo ============================================================
echo  완료!  결과물:  dist\convert_and_split.exe
echo ============================================================
echo.
echo  배포 시에는 dist\convert_and_split.exe 파일만 복사하면 됩니다.
echo  같은 폴더에 register_context_menu.reg 를 함께 두면
echo  우클릭 메뉴 등록도 가능합니다.
echo.
pause
endlocal