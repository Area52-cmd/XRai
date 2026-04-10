# XRai Installation Script
# Run as: powershell -ExecutionPolicy Bypass -File install.ps1

$ErrorActionPreference = "Stop"

Write-Host ""
Write-Host "========================================" -ForegroundColor Cyan
Write-Host "  XRai Installer" -ForegroundColor Cyan
Write-Host "  AI-powered Excel automation" -ForegroundColor Cyan
Write-Host "========================================" -ForegroundColor Cyan
Write-Host ""

# Step 1: Check prerequisites
Write-Host "[1/5] Checking prerequisites..." -ForegroundColor Yellow

# .NET 8+
$dotnetVersion = dotnet --version 2>$null
if (-not $dotnetVersion) {
    Write-Host "  FAIL: .NET SDK not found. Install from https://dotnet.microsoft.com/download" -ForegroundColor Red
    exit 1
}
Write-Host "  .NET SDK: $dotnetVersion" -ForegroundColor Green

# Excel
$excel = Get-Process EXCEL -ErrorAction SilentlyContinue
if ($excel) {
    Write-Host "  Excel: Running (PID $($excel[0].Id))" -ForegroundColor Green
} else {
    Write-Host "  Excel: Not running (will need to start it before using XRai)" -ForegroundColor Yellow
}

# Step 2: Build solution
Write-Host ""
Write-Host "[2/5] Building XRai..." -ForegroundColor Yellow
$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
Push-Location $scriptDir

dotnet build XRai.sln -c Release --verbosity quiet
if ($LASTEXITCODE -ne 0) {
    Write-Host "  FAIL: Build failed" -ForegroundColor Red
    Pop-Location
    exit 1
}
Write-Host "  Build: Succeeded" -ForegroundColor Green

# Step 3: Build demo add-in
Write-Host ""
Write-Host "[3/5] Building demo add-in..." -ForegroundColor Yellow
dotnet build demo/XRai.Demo.PortfolioAddin/XRai.Demo.PortfolioAddin.csproj -c Release --verbosity quiet 2>$null
if ($LASTEXITCODE -eq 0) {
    Write-Host "  Demo add-in: Built" -ForegroundColor Green
    $xllPath = Get-ChildItem -Path "demo/XRai.Demo.PortfolioAddin/bin/Release" -Filter "*AddIn64.xll" -Recurse | Select-Object -First 1
    if ($xllPath) {
        Write-Host "  XLL: $($xllPath.FullName)" -ForegroundColor Green
    }
} else {
    Write-Host "  Demo add-in: Skipped (build warning — non-fatal)" -ForegroundColor Yellow
}

# Step 4: Install as global dotnet tool
Write-Host ""
Write-Host "[4/5] Installing xrai global tool..." -ForegroundColor Yellow
dotnet tool uninstall -g XRai.Tool 2>$null
dotnet pack src/XRai.Tool/XRai.Tool.csproj -c Release --verbosity quiet -o ./nupkg 2>$null
if ($LASTEXITCODE -eq 0) {
    $nupkg = Get-ChildItem -Path "./nupkg" -Filter "*.nupkg" | Sort-Object LastWriteTime -Descending | Select-Object -First 1
    if ($nupkg) {
        dotnet tool install -g --add-source ./nupkg XRai.Tool 2>$null
        if ($LASTEXITCODE -eq 0) {
            Write-Host "  Global tool: Installed (run 'xrai' from anywhere)" -ForegroundColor Green
        } else {
            Write-Host "  Global tool: Install failed (use 'dotnet run --project src/XRai.Tool' instead)" -ForegroundColor Yellow
        }
    }
} else {
    Write-Host "  Global tool: Pack failed (use 'dotnet run --project src/XRai.Tool' instead)" -ForegroundColor Yellow
}

# Step 5: Run doctor
Write-Host ""
Write-Host "[5/5] Running xrai doctor..." -ForegroundColor Yellow
dotnet run --project src/XRai.Tool -c Release -- doctor
Write-Host ""

# Step 6: Run tests
Write-Host ""
Write-Host "Running unit tests..." -ForegroundColor Yellow
dotnet test tests/XRai.Tests.Unit/XRai.Tests.Unit.csproj -c Release --verbosity minimal
Write-Host ""

Pop-Location

Write-Host "========================================" -ForegroundColor Cyan
Write-Host "  Installation complete!" -ForegroundColor Green
Write-Host "========================================" -ForegroundColor Cyan
Write-Host ""
Write-Host "Quick start:" -ForegroundColor White
Write-Host "  1. Open Excel with a workbook" -ForegroundColor Gray
Write-Host "  2. Run: dotnet run --project src/XRai.Tool -- --wait" -ForegroundColor Gray
Write-Host '  3. Send: {"cmd":"read","ref":"A1"}' -ForegroundColor Gray
Write-Host ""
Write-Host "Demo add-in:" -ForegroundColor White
Write-Host "  Load the .xll file in Excel > File > Options > Add-ins" -ForegroundColor Gray
if ($xllPath) {
    Write-Host "  Path: $($xllPath.FullName)" -ForegroundColor Gray
}
Write-Host ""
Write-Host "Test harness:" -ForegroundColor White
Write-Host "  Run: powershell -File test-harness.ps1" -ForegroundColor Gray
Write-Host ""
