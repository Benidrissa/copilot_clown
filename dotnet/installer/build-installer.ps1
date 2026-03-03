# Build script for Copilot Clown installer
# Prerequisites: .NET SDK, Inno Setup (iscc.exe in PATH or default location)
#
# Usage: .\build-installer.ps1

$ErrorActionPreference = "Stop"
$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$dotnetDir = Split-Path -Parent $scriptDir

Write-Host "`n  [1/3] Building .NET project (Release)..." -ForegroundColor Cyan
Push-Location $dotnetDir
dotnet build -c Release
if ($LASTEXITCODE -ne 0) { Write-Host "Build failed!" -ForegroundColor Red; exit 1 }
Pop-Location

Write-Host "`n  [2/3] Compiling installer with Inno Setup..." -ForegroundColor Cyan

# Find iscc.exe
$iscc = $null
$locations = @(
    "C:\Program Files (x86)\Inno Setup 6\ISCC.exe",
    "C:\Program Files\Inno Setup 6\ISCC.exe",
    "C:\Program Files (x86)\Inno Setup 5\ISCC.exe"
)
foreach ($loc in $locations) {
    if (Test-Path $loc) { $iscc = $loc; break }
}
if (-not $iscc) {
    # Try PATH
    $iscc = Get-Command iscc -ErrorAction SilentlyContinue | Select-Object -ExpandProperty Source
}
if (-not $iscc) {
    Write-Host "Error: Inno Setup not found. Download from https://jrsoftware.org/isinfo.php" -ForegroundColor Red
    exit 1
}

& $iscc "$scriptDir\CopilotClown.iss"
if ($LASTEXITCODE -ne 0) { Write-Host "Installer compilation failed!" -ForegroundColor Red; exit 1 }

Write-Host "`n  [3/3] Done!" -ForegroundColor Green
$output = Join-Path $scriptDir "Output\CopilotClownSetup.exe"
$size = [math]::Round((Get-Item $output).Length / 1MB, 1)
Write-Host "  Output: $output ($size MB)`n" -ForegroundColor White
