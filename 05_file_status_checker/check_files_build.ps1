# ================================================================
# Build check_files.py -> check_files.exe
# Usage: Right-click -> Run with PowerShell
# ================================================================

$ScriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$PythonScript = Join-Path $ScriptDir "check_files.py"

Write-Host ""
Write-Host "=======================================" -ForegroundColor Cyan
Write-Host "  Build: check_files.exe" -ForegroundColor Cyan
Write-Host "=======================================" -ForegroundColor Cyan
Write-Host ""

if (-not (Test-Path $PythonScript)) {
    Write-Host "[ERROR] check_files.py not found." -ForegroundColor Red
    Write-Host "        Place check_files.py in the same folder as this script." -ForegroundColor Red
    Read-Host "Press Enter to exit"
    exit 1
}

Write-Host "[1/4] Checking Python..." -ForegroundColor Yellow
try {
    $PythonVersion = python --version 2>&1
    Write-Host "      $PythonVersion found" -ForegroundColor Green
} catch {
    Write-Host "[ERROR] Python not found or not in PATH." -ForegroundColor Red
    Read-Host "Press Enter to exit"
    exit 1
}

Write-Host ""
Write-Host "[2/4] Installing dependencies..." -ForegroundColor Yellow

$Packages = @("openpyxl", "python-docx", "python-pptx", "PyMuPDF", "olefile", "py7zr", "rarfile", "xlrd", "pyinstaller")
foreach ($pkg in $Packages) {
    Write-Host "      pip install $pkg ..." -NoNewline
    $result = python -m pip install $pkg 2>&1 | Out-String
    if ($LASTEXITCODE -eq 0) {
        Write-Host " OK" -ForegroundColor Green
    } else {
        Write-Host " FAIL" -ForegroundColor Red
        Write-Host "      $result" -ForegroundColor Red
    }
}

Write-Host ""
Write-Host "[3/4] Building exe (this may take a while)..." -ForegroundColor Yellow

$BuildCmd = "pyinstaller " +
            "--onefile " +
            "--name check_files " +
            "--distpath `"$ScriptDir\dist`" " +
            "--workpath `"$ScriptDir\build`" " +
            "--specpath `"$ScriptDir`" " +
            "`"$PythonScript`""

Write-Host "      CMD: $BuildCmd" -ForegroundColor Gray
Invoke-Expression $BuildCmd

if ($LASTEXITCODE -ne 0) {
    Write-Host ""
    Write-Host "[ERROR] Build failed. Check the error messages above." -ForegroundColor Red
    Read-Host "Press Enter to exit"
    exit 1
}

Write-Host ""
Write-Host "[4/4] Verifying output..." -ForegroundColor Yellow

$ExePath = Join-Path $ScriptDir "dist\check_files.exe"
if (Test-Path $ExePath) {
    $ExeSize = [math]::Round((Get-Item $ExePath).Length / 1MB, 1)
    Write-Host ""
    Write-Host "=======================================" -ForegroundColor Green
    Write-Host "  Build succeeded!" -ForegroundColor Green
    Write-Host "  Output : $ExePath" -ForegroundColor Green
    Write-Host "  Size   : $ExeSize MB" -ForegroundColor Green
    Write-Host "=======================================" -ForegroundColor Green
    Write-Host ""
    Write-Host "* Distribute dist\check_files.exe only." -ForegroundColor Cyan
    Write-Host "* build\ folder and .spec file can be deleted." -ForegroundColor Cyan
    Start-Process explorer.exe -ArgumentList (Join-Path $ScriptDir "dist")
} else {
    Write-Host "[ERROR] exe file was not created." -ForegroundColor Red
}

Write-Host ""
Read-Host "Press Enter to exit"