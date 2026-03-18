# ============================================================
#  Build Script - Classification of Cost Item Groups
#  Run: Right-click > Run with PowerShell
# ============================================================

chcp 65001 | Out-Null
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8

$SRC_NAME = "classification_of_cost_item_groups"
$EXE_NAME = "classification_of_cost_item_groups"

Write-Host ""
Write-Host "============================================"
Write-Host "  EXE Build - Cost Item Group Classification"
Write-Host "============================================"
Write-Host ""

# --- 1. Python ---
Write-Host "[1/4] Checking Python..."
try {
    $pyVer = python --version 2>&1
    Write-Host "  OK: $pyVer"
} catch {
    Write-Host "  ERROR: Python not found."
    Write-Host "  Install from https://www.python.org/downloads/"
    Write-Host "  (Check 'Add Python to PATH')"
    pause
    exit 1
}

# --- 2. Packages ---
Write-Host ""
Write-Host "[2/4] Installing packages..."
pip install pyinstaller openpyxl --quiet 2>&1 | Out-Null
if ($LASTEXITCODE -ne 0) {
    pip install pyinstaller openpyxl
}
Write-Host "  OK: pyinstaller, openpyxl"

# --- 3. Source file ---
Write-Host ""
Write-Host "[3/4] Checking source file..."
$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$srcFile = Join-Path $scriptDir "$SRC_NAME.py"

if (-not (Test-Path $srcFile)) {
    Write-Host "  ERROR: $SRC_NAME.py not found."
    Write-Host "  Place build.ps1 and $SRC_NAME.py in the same folder."
    pause
    exit 1
}
Write-Host "  OK: $srcFile"

# --- 4. Build ---
Write-Host ""
Write-Host "[4/4] Building EXE... (1-2 min)"

Push-Location $scriptDir
pyinstaller --onefile --name $EXE_NAME --console --noconfirm --clean "$srcFile"
Pop-Location

# --- Result ---
$exePath = Join-Path $scriptDir "dist\$EXE_NAME.exe"

Write-Host ""
if (Test-Path $exePath) {
    $size = [math]::Round((Get-Item $exePath).Length / 1MB, 1)
    Write-Host "============================================"
    Write-Host "  BUILD SUCCESS!"
    Write-Host "============================================"
    Write-Host ""
    Write-Host "  EXE: $exePath"
    Write-Host "  Size: ${size}MB"
    Write-Host ""
    Write-Host "  [How to use]"
    Write-Host "  1. Copy $EXE_NAME.exe to any folder"
    Write-Host "  2. Drag & drop Excel file onto the exe"
    Write-Host "     Or CMD: $EXE_NAME.exe input.xlsx"
    Write-Host "  3. Output: input_classified.xlsx"
    Write-Host ""

    $destExe = Join-Path $scriptDir "$EXE_NAME.exe"
    Copy-Item $exePath $destExe -Force
    Write-Host "  Copied to: $destExe"

    # Cleanup
    $buildDir = Join-Path $scriptDir "build"
    $specFile = Join-Path $scriptDir "$EXE_NAME.spec"
    if (Test-Path $buildDir) { Remove-Item $buildDir -Recurse -Force }
    if (Test-Path $specFile) { Remove-Item $specFile -Force }
    Write-Host "  Temp files cleaned."
} else {
    Write-Host "============================================"
    Write-Host "  BUILD FAILED"
    Write-Host "============================================"
    Write-Host "  Check error messages above."
}

Write-Host ""
pause