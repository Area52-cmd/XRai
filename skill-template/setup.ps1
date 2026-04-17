# XRai Setup — run once after extracting the zip
# Usage: Right-click this file > "Run with PowerShell"
#   OR: powershell -ExecutionPolicy Bypass -File setup.ps1

$ErrorActionPreference = "Stop"

Write-Host ""
Write-Host "  =======================================================" -ForegroundColor Cyan
Write-Host "    XRai Setup" -ForegroundColor Cyan
Write-Host "    Think Lovable, but for desktop apps" -ForegroundColor Cyan
Write-Host "  =======================================================" -ForegroundColor Cyan
Write-Host ""

# Step 1: Install files
$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$skillDir  = Join-Path $env:USERPROFILE ".claude\skills\xrai-excel"

if ($scriptDir -ne $skillDir) {
    Write-Host "  [1/3] Installing to $skillDir ..." -ForegroundColor Yellow
    if (Test-Path $skillDir) { Remove-Item $skillDir -Recurse -Force }
    Copy-Item $scriptDir $skillDir -Recurse -Force
    Write-Host "        Done." -ForegroundColor Green
} else {
    Write-Host "  [1/3] Already installed." -ForegroundColor Green
}

# Step 2: Add bin to PATH
$binDir = Join-Path $skillDir "bin"
Write-Host "  [2/3] Checking PATH ..." -ForegroundColor Yellow
$userPath = [Environment]::GetEnvironmentVariable("PATH", "User")
if ($userPath -notlike "*$binDir*") {
    [Environment]::SetEnvironmentVariable("PATH", "$userPath;$binDir", "User")
    Write-Host "        Added xrai to PATH (new terminals will see it)." -ForegroundColor Green
} else {
    Write-Host "        Already on PATH." -ForegroundColor Green
}

# Step 3: Auto-launch Studio on every Claude Code session
Write-Host "  [3/3] Configuring auto-launch ..." -ForegroundColor Yellow
$settingsPath = Join-Path $env:USERPROFILE ".claude\settings.json"

$settings = @{}
if (Test-Path $settingsPath) {
    try {
        $raw = Get-Content $settingsPath -Raw
        if ($raw.Trim().Length -gt 0) {
            $settings = $raw | ConvertFrom-Json -AsHashtable
        }
    } catch {
        Copy-Item $settingsPath "$settingsPath.bak" -ErrorAction SilentlyContinue
        $settings = @{}
    }
}

if (-not $settings.ContainsKey("hooks")) { $settings["hooks"] = @{} }
if (-not $settings["hooks"].ContainsKey("SessionStart")) { $settings["hooks"]["SessionStart"] = @() }

$already = $false
foreach ($hook in $settings["hooks"]["SessionStart"]) {
    if ($hook -is [hashtable] -and $hook.ContainsKey("hooks")) {
        foreach ($h in $hook["hooks"]) {
            if ($h -is [hashtable] -and $h["command"] -like "*xrai*studio*") {
                $already = $true
            }
        }
    }
}

if (-not $already) {
    $settings["hooks"]["SessionStart"] += @(
        @{
            "matcher" = "xrai-studio-autolaunch"
            "hooks" = @(
                @{
                    "type" = "command"
                    "command" = "xrai --studio --no-browser 2>NUL &"
                }
            )
        }
    )
    $settingsDir = Split-Path $settingsPath
    if (-not (Test-Path $settingsDir)) { New-Item -ItemType Directory -Path $settingsDir -Force | Out-Null }
    Set-Content $settingsPath ($settings | ConvertTo-Json -Depth 10) -Encoding UTF8
    Write-Host "        Studio will auto-launch on every Claude Code session." -ForegroundColor Green
} else {
    Write-Host "        Already configured." -ForegroundColor Green
}

# Done
Write-Host ""
Write-Host "  =======================================================" -ForegroundColor Green
Write-Host "    Setup complete!" -ForegroundColor Green
Write-Host "  =======================================================" -ForegroundColor Green
Write-Host ""
Write-Host "  Next steps:" -ForegroundColor White
Write-Host "    1. Open a NEW terminal (so PATH takes effect)" -ForegroundColor Gray
Write-Host "    2. cd into any project folder" -ForegroundColor Gray
Write-Host "    3. Type: claude" -ForegroundColor Gray
Write-Host "    4. Say: Build me an Excel add-in that does X" -ForegroundColor Gray
Write-Host ""
Write-Host "  Studio + IDE + Excel will open automatically." -ForegroundColor Gray
Write-Host "  You watch everything happen live." -ForegroundColor Gray
Write-Host ""

if ($Host.Name -eq "ConsoleHost") {
    Write-Host "  Press any key to close ..." -ForegroundColor DarkGray
    $null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
}
