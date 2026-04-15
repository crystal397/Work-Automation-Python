@echo off
chcp 65001 > nul
echo.
echo ========================================
echo   귀책분석 자동화 시스템 빌드
echo ========================================
echo.

REM ── 1. PyInstaller 빌드 ──────────────────────────────────────
echo [1/3] PyInstaller 빌드 중...
pyinstaller 귀책분석_자동화.spec --noconfirm
if errorlevel 1 (
    echo.
    echo [오류] 빌드 실패. 위 오류 메시지를 확인하세요.
    pause
    exit /b 1
)

REM ── 2. 필수 데이터 파일 복사 ─────────────────────────────────
echo.
echo [2/3] 데이터 파일 복사 중...

set DIST=dist\귀책분석_자동화

copy /Y "귀책분석_패턴집.md" "%DIST%\귀책분석_패턴집.md" > nul
if errorlevel 1 (
    echo [경고] 귀책분석_패턴집.md 복사 실패
)

if not exist "%DIST%\output" mkdir "%DIST%\output"
if exist "output\reference_patterns.md" (
    copy /Y "output\reference_patterns.md" "%DIST%\output\reference_patterns.md" > nul
    echo   reference_patterns.md 복사 완료
) else (
    echo   [경고] output\reference_patterns.md 없음 — 배포 후 실행 시 경고 표시됨
    echo         먼저 python main.py learn 을 실행하세요
)

REM ── 3. 완료 안내 ─────────────────────────────────────────────
echo.
echo [3/3] 완료
echo.
echo   배포 폴더: %DIST%\
echo.
echo   배포 시 이 폴더 전체를 압축하세요:
echo     - 귀책분석_자동화.exe
echo     - 귀책분석_패턴집.md
echo     - output\reference_patterns.md
echo     - _internal\  (런타임 라이브러리)
echo.
echo   사용자가 reference\ 폴더를 이 폴더 안에 직접 넣어야 합니다.
echo.
pause
