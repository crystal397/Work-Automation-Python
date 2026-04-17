# PDF 합치기 도구 - EXE 빌드 스크립트 (PowerShell)
# 실행: PowerShell에서 .\build.ps1

$OutputEncoding = [System.Text.Encoding]::UTF8
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8
$ErrorActionPreference = "Stop"

Write-Host "================================================"
Write-Host " PDF Merger - EXE Build Script"
Write-Host "================================================"
Write-Host ""

# 1. 패키지 설치
Write-Host "[1/3] Installing packages..."
pip install reportlab pypdf pywin32 pyinstaller
if ($LASTEXITCODE -ne 0) {
    Write-Host "[ERROR] Package installation failed."
    Read-Host "Press Enter to exit"
    exit 1
}

# 2. spec 파일 생성
Write-Host ""
Write-Host "[2/3] Creating spec file..."

$specPath = Join-Path $PWD "pdf_merger.spec"

$specContent = @'
# -*- mode: python ; coding: utf-8 -*-
from PyInstaller.utils.hooks import collect_all

rl_datas,  rl_binaries,  rl_hiddens  = collect_all("reportlab")
pdf_datas, pdf_binaries, pdf_hiddens = collect_all("pypdf")

a = Analysis(
    ["pdf_merger_v4.py"],
    pathex=[],
    binaries=rl_binaries + pdf_binaries,
    datas=rl_datas + pdf_datas,
    hiddenimports=rl_hiddens + pdf_hiddens + [
        "win32com.client", "win32api", "pythoncom", "pywintypes",
        "tkinter", "tkinter.filedialog", "tkinter.scrolledtext", "tkinter.messagebox",
    ],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[],
    noarchive=False,
)
pyz = PYZ(a.pure)
exe = EXE(
    pyz, a.scripts, a.binaries, a.datas,
    name="pdf_merger",
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=False,
    console=False,
    onefile=True,
)
'@

[System.IO.File]::WriteAllText($specPath, $specContent, [System.Text.Encoding]::ASCII)
Write-Host "spec file created: $specPath"

# 3. EXE 빌드
Write-Host ""
Write-Host "[3/3] Building EXE... (2~5 min)"
pyinstaller $specPath

if ($LASTEXITCODE -ne 0) {
    Write-Host "[ERROR] Build failed."
    Read-Host "Press Enter to exit"
    exit 1
}

Write-Host ""
Write-Host "================================================"
Write-Host " Done! Output: dist\pdf_merger.exe"
Write-Host "================================================"
Write-Host ""
Write-Host " - Double-click dist\pdf_merger.exe to launch GUI"
Write-Host " - Or use CLI: .\dist\pdf_merger.exe <folder> [output.pdf]"
Write-Host ""
Read-Host "Press Enter to exit"
