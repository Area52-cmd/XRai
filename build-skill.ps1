# build-skill.ps1 — Assemble the XRai Claude Code skill distribution.
#
# Output:
#   dist/xrai-excel-skill/          <- the skill directory, ready to drop into ~/.claude/skills/
#   dist/xrai-excel-skill.zip       <- the distributable zip
#
# Run:
#   powershell -ExecutionPolicy Bypass -File build-skill.ps1

$ErrorActionPreference = "Stop"
$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
Push-Location $scriptDir

Write-Host ""
Write-Host "  ===============================================" -ForegroundColor Cyan
Write-Host "    Building XRai Claude Code Skill Distribution" -ForegroundColor Cyan
Write-Host "  ===============================================" -ForegroundColor Cyan
Write-Host ""

# ── Step 1: Build the .NET solution (Release) ─────────────────────────
Write-Host "[1/6] Building XRai.sln in Release..." -ForegroundColor Yellow
dotnet build XRai.sln -c Release --verbosity quiet
if ($LASTEXITCODE -ne 0) {
    Write-Host "  FAIL: build failed" -ForegroundColor Red
    Pop-Location
    exit 1
}
Write-Host "  OK" -ForegroundColor Green

# ── Step 2: Publish self-contained XRai.Tool binary ────────────────────
Write-Host ""
Write-Host "[2/6] Publishing self-contained XRai.Tool..." -ForegroundColor Yellow
$binDir = "dist/xrai-bin"
if (Test-Path $binDir) { Remove-Item $binDir -Recurse -Force }

dotnet publish src/XRai.Tool/XRai.Tool.csproj `
    -c Release `
    -r win-x64 `
    --self-contained true `
    -p:PublishSingleFile=false `
    -o $binDir `
    --verbosity quiet
if ($LASTEXITCODE -ne 0) {
    Write-Host "  FAIL: publish failed" -ForegroundColor Red
    Pop-Location
    exit 1
}
Write-Host "  OK ($(Get-ChildItem $binDir | Measure-Object).Count files)" -ForegroundColor Green

# ── Step 3: Pack XRai.Hooks NuGet ──────────────────────────────────────
Write-Host ""
Write-Host "[3/6] Packing XRai.Hooks NuGet package..." -ForegroundColor Yellow
if (-not (Test-Path "nupkg")) { New-Item -ItemType Directory -Path "nupkg" | Out-Null }
dotnet pack src/XRai.Hooks/XRai.Hooks.csproj -c Release -o nupkg --verbosity quiet
if ($LASTEXITCODE -ne 0) {
    Write-Host "  FAIL: pack failed" -ForegroundColor Red
    Pop-Location
    exit 1
}
$nupkgFile = Get-ChildItem "nupkg/XRai.Hooks.*.nupkg" | Sort-Object LastWriteTime -Descending | Select-Object -First 1
Write-Host "  OK: $($nupkgFile.Name)" -ForegroundColor Green

# ── Step 4: Assemble the skill directory ───────────────────────────────
Write-Host ""
Write-Host "[4/6] Assembling skill directory..." -ForegroundColor Yellow
$skillDir = "dist/xrai-excel-skill"
if (Test-Path $skillDir) { Remove-Item $skillDir -Recurse -Force }
New-Item -ItemType Directory -Path $skillDir | Out-Null
New-Item -ItemType Directory -Path "$skillDir/bin" | Out-Null
New-Item -ItemType Directory -Path "$skillDir/packages" | Out-Null
New-Item -ItemType Directory -Path "$skillDir/docs" | Out-Null
New-Item -ItemType Directory -Path "$skillDir/templates" | Out-Null
New-Item -ItemType Directory -Path "$skillDir/reference" | Out-Null

# Copy binary
Copy-Item "$binDir/*" "$skillDir/bin/" -Recurse -Force

# Copy NuGet package
Copy-Item $nupkgFile.FullName "$skillDir/packages/" -Force

# Copy SKILL.md, docs, reference, templates, and setup script from skill-template
Copy-Item "skill-template/SKILL.md" "$skillDir/SKILL.md" -Force
Copy-Item "skill-template/docs/*" "$skillDir/docs/" -Recurse -Force
Copy-Item "skill-template/reference/*" "$skillDir/reference/" -Recurse -Force
Copy-Item "skill-template/templates/*" "$skillDir/templates/" -Recurse -Force
if (Test-Path "skill-template/setup.ps1") {
    Copy-Item "skill-template/setup.ps1" "$skillDir/setup.ps1" -Force
}

# Auto-generate commands.md from the freshly built binary so docs are never stale.
# This replaces the hand-written skill-template/docs/commands.md copy.
Write-Host "  Auto-generating commands.md from binary..." -ForegroundColor DarkGray
& "$skillDir/bin/XRai.Tool.exe" dump-commands | Out-File "$skillDir/docs/commands.md" -Encoding UTF8
if ((Get-Content "$skillDir/docs/commands.md" | Measure-Object -Line).Lines -lt 10) {
    Write-Host "  WARNING: commands.md auto-generation produced empty file" -ForegroundColor Yellow
} else {
    Write-Host "  OK: commands.md generated" -ForegroundColor DarkGray
}

$binCount = (Get-ChildItem "$skillDir/bin" -File).Count
$docCount = (Get-ChildItem "$skillDir/docs" -File).Count
$tmplCount = (Get-ChildItem "$skillDir/templates" -File).Count
Write-Host "  OK: $binCount bin files, $docCount docs, $tmplCount templates" -ForegroundColor Green

# ── Step 5: Zip it ─────────────────────────────────────────────────────
Write-Host ""
Write-Host "[5/6] Creating distributable zip..." -ForegroundColor Yellow
$zipPath = "dist/xrai-excel-skill.zip"
if (Test-Path $zipPath) { Remove-Item $zipPath -Force }
Compress-Archive -Path "$skillDir" -DestinationPath $zipPath -Force
$zipSize = [math]::Round((Get-Item $zipPath).Length / 1MB, 1)
Write-Host "  OK: $zipPath ($zipSize MB)" -ForegroundColor Green

# ── Step 6: Summary ────────────────────────────────────────────────────
Write-Host ""
Write-Host "[6/6] Summary" -ForegroundColor Yellow
$skillSize = [math]::Round(((Get-ChildItem $skillDir -Recurse | Measure-Object Length -Sum).Sum / 1MB), 1)
Write-Host "  Skill directory: $skillDir ($skillSize MB)" -ForegroundColor Green
Write-Host "  Zip archive:     $zipPath ($zipSize MB)" -ForegroundColor Green
Write-Host ""
Write-Host "  ===============================================" -ForegroundColor Cyan
Write-Host "    Build complete!" -ForegroundColor Green
Write-Host "  ===============================================" -ForegroundColor Cyan
Write-Host ""
# ── Step 7: Auto-install locally ──────────────────────────────────────
Write-Host ""
Write-Host "[7/7] Installing to local skills directory..." -ForegroundColor Yellow
$dest = "$env:USERPROFILE\.claude\skills\xrai-excel"

# Kill any running XRai processes that would lock DLLs in the skill bin dir
$xraiProcs = Get-Process -Name "XRai.Tool","XRai.Mcp" -ErrorAction SilentlyContinue
if ($xraiProcs) {
    Write-Host "  Stopping running XRai processes..." -ForegroundColor DarkYellow
    $xraiProcs | Stop-Process -Force -ErrorAction SilentlyContinue
    Start-Sleep -Seconds 2
}

# Remove old skill dir (retry if DLLs are still locked)
if (Test-Path $dest) {
    $retries = 3
    for ($i = 0; $i -lt $retries; $i++) {
        try {
            Remove-Item $dest -Recurse -Force -ErrorAction Stop
            break
        } catch {
            if ($i -eq $retries - 1) {
                Write-Host "  WARNING: Could not remove old skill dir (DLLs locked). Overwriting in place..." -ForegroundColor Yellow
            } else {
                Start-Sleep -Seconds 2
            }
        }
    }
}

# Copy fresh skill
Copy-Item -Recurse -Force $skillDir $dest

# Also publish and deploy the MCP server if the project exists
$mcpProj = "src/XRai.Mcp/XRai.Mcp.csproj"
if (Test-Path $mcpProj) {
    Write-Host "  Publishing MCP server..." -ForegroundColor DarkGray
    $mcpBin = "dist/xrai-mcp-bin"
    if (Test-Path $mcpBin) { Remove-Item $mcpBin -Recurse -Force }
    dotnet publish $mcpProj -c Release -r win-x64 --self-contained true -o $mcpBin --verbosity quiet 2>&1 | Out-Null
    if (Test-Path "$mcpBin/XRai.Mcp.exe") {
        Copy-Item "$mcpBin/XRai.Mcp.exe" "$dest/bin/XRai.Mcp.exe" -Force
        Write-Host "  MCP server deployed to skill bin" -ForegroundColor DarkGray
    }
}

Write-Host "  Installed to: $dest" -ForegroundColor Green

# ── Drop an xrai.cmd shim so `xrai ...` works from any shell once the
#    bin directory is on PATH. The shim forwards all args to the real exe.
#    Using a .cmd file (not a .bat) so PowerShell execution-policy rules
#    don't apply — .cmd is treated as native by cmd.exe / PowerShell both.
$shimPath = "$dest/bin/xrai.cmd"
$shimContent = @"
@echo off
"%~dp0XRai.Tool.exe" %*
"@
Set-Content -Path $shimPath -Value $shimContent -Encoding ASCII -NoNewline
Write-Host "  Shim:           xrai.cmd -> XRai.Tool.exe" -ForegroundColor Green

# ── Add the bin directory to the user PATH if not already on it, so
#    plain `xrai --studio` works from any new PowerShell / cmd.exe shell.
#    Existing shells won't pick it up until they're restarted — we print
#    a hint about that.
$binDir = "$dest\bin"
$userPath = [Environment]::GetEnvironmentVariable("PATH", "User")
$pathSegments = $userPath -split ';' | Where-Object { $_ }
$alreadyOnPath = $pathSegments -contains $binDir

if ($alreadyOnPath) {
    Write-Host "  PATH:           already contains $binDir" -ForegroundColor Green
} else {
    try {
        $newPath = if ($userPath) { "$userPath;$binDir" } else { $binDir }
        [Environment]::SetEnvironmentVariable("PATH", $newPath, "User")
        Write-Host "  PATH:           added $binDir to user PATH" -ForegroundColor Green
        Write-Host "                  (restart your shell for xrai to be found)" -ForegroundColor DarkYellow
    } catch {
        $errMsg = $_.Exception.Message
        Write-Host "  PATH:           WARNING - could not add to user PATH: $errMsg" -ForegroundColor Yellow
        Write-Host "                  Run this manually in PowerShell:" -ForegroundColor DarkYellow
        Write-Host ('                  [Environment]::SetEnvironmentVariable("PATH", $env:PATH + ";" + "' + $binDir + '", "User")') -ForegroundColor DarkGray
    }
}

# Verify
$toolExe = Get-Item "$dest/bin/XRai.Tool.exe"
$hooksPkg = Get-ChildItem "$dest/packages/*.nupkg" | Sort-Object LastWriteTime -Descending | Select-Object -First 1
Write-Host "  XRai.Tool.exe:  $($toolExe.LastWriteTime)" -ForegroundColor Green
Write-Host "  Hooks package:  $($hooksPkg.Name)" -ForegroundColor Green

Write-Host ""
Write-Host "  ===============================================" -ForegroundColor Cyan
Write-Host "    Build + install complete!" -ForegroundColor Green
Write-Host "  ===============================================" -ForegroundColor Cyan
Write-Host ""
Write-Host "  Run `xrai --studio` from any new terminal." -ForegroundColor DarkGray
Write-Host "  (existing terminals need a restart to see the PATH change)" -ForegroundColor DarkGray
Write-Host ""
Write-Host "Restart Claude Code to pick up the new build." -ForegroundColor White
Write-Host ""

Pop-Location
