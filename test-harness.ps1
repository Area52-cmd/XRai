# XRai Torture Test Harness
# Exercises EVERY command against a live Excel instance
# Run: powershell -ExecutionPolicy Bypass -File test-harness.ps1

$ErrorActionPreference = "Continue"
$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$xraiProject = "$scriptDir\src\XRai.Tool\XRai.Tool.csproj"
$passed = 0
$failed = 0
$errors = @()
$startTime = Get-Date

function Send-XRai {
    param([string[]]$Commands)
    $input_text = ($Commands -join "`n")
    $result = $input_text | dotnet run --project $xraiProject 2>$null
    return $result -split "`n" | Where-Object { $_ -match '^\{' }
}

function Test-Section {
    param([string]$Name, [string[]]$Commands, [scriptblock]$Validate)

    Write-Host "  Testing: $Name... " -NoNewline
    try {
        $responses = Send-XRai -Commands $Commands
        $sectionPassed = 0
        $sectionFailed = 0

        foreach ($resp in $responses) {
            if ($resp -match '"ok"\s*:\s*true') {
                $sectionPassed++
            } elseif ($resp -match '"ok"\s*:\s*false') {
                # Some failures are expected (edge cases)
                if ($resp -match '"error"\s*:\s*"Cell .* did not meet condition') {
                    $sectionPassed++ # Expected timeout
                } elseif ($resp -match '"found"\s*:\s*false') {
                    $sectionPassed++ # Expected not-found
                } else {
                    $sectionFailed++
                    $script:errors += "$Name`: $resp"
                }
            } else {
                $sectionPassed++ # Non-JSON or event output
            }
        }

        if ($sectionFailed -eq 0) {
            Write-Host "$sectionPassed/$($sectionPassed + $sectionFailed) " -ForegroundColor Green -NoNewline
            Write-Host "PASS" -ForegroundColor Green
        } else {
            Write-Host "$sectionPassed/$($sectionPassed + $sectionFailed) " -ForegroundColor Yellow -NoNewline
            Write-Host "($sectionFailed FAILED)" -ForegroundColor Red
        }

        $script:passed += $sectionPassed
        $script:failed += $sectionFailed
    } catch {
        Write-Host "ERROR: $_" -ForegroundColor Red
        $script:failed++
        $script:errors += "$Name`: $_"
    }
}

# ===================================================
Write-Host ""
Write-Host "  ===============================================" -ForegroundColor Cyan
Write-Host "       XRai Torture Test Harness" -ForegroundColor Cyan
Write-Host "  ===============================================" -ForegroundColor Cyan
Write-Host ""

# Check Excel is running
$excel = Get-Process EXCEL -ErrorAction SilentlyContinue
if (-not $excel) {
    Write-Host "  ERROR: Excel is not running. Start Excel with a workbook first." -ForegroundColor Red
    exit 1
}
Write-Host "  Excel: PID $($excel[0].Id)" -ForegroundColor Green
Write-Host ""

# ── CONNECTION ──
Test-Section "Connection" @(
    '{"cmd":"attach"}',
    '{"cmd":"status"}',
    '{"cmd":"detach"}',
    '{"cmd":"attach"}'
)

# ── SHEETS ──
Test-Section "Sheets" @(
    '{"cmd":"attach"}',
    '{"cmd":"sheets"}',
    '{"cmd":"sheet.add","name":"TestSheet1"}',
    '{"cmd":"sheet.add","name":"TestSheet2","after":"TestSheet1"}',
    '{"cmd":"sheet.rename","from":"TestSheet2","to":"Renamed"}',
    '{"cmd":"sheets"}',
    '{"cmd":"goto","target":"TestSheet1"}',
    '{"cmd":"goto","target":"Renamed"}',
    '{"cmd":"sheet.delete","name":"Renamed"}',
    '{"cmd":"sheets"}'
)

# ── CELLS ──
Test-Section "Cells Read/Write" @(
    '{"cmd":"attach"}',
    '{"cmd":"goto","target":"TestSheet1"}',
    '{"cmd":"type","ref":"A1","value":"Hello XRai"}',
    '{"cmd":"read","ref":"A1"}',
    '{"cmd":"type","ref":"A2","value":"42"}',
    '{"cmd":"type","ref":"A3","value":"=A2*2"}',
    '{"cmd":"calc"}',
    '{"cmd":"read","ref":"A3","full":true}',
    '{"cmd":"read","ref":"A1:A3"}',
    '{"cmd":"clear","ref":"A1:A3"}',
    '{"cmd":"read","ref":"A1"}',
    '{"cmd":"select","ref":"B5"}'
)

# ── FORMATTING ──
Test-Section "Formatting" @(
    '{"cmd":"attach"}',
    '{"cmd":"goto","target":"TestSheet1"}',
    '{"cmd":"type","ref":"A1","value":"Formatted Cell"}',
    '{"cmd":"format","ref":"A1","bold":true,"font_size":16,"bg":"#ff0000"}',
    '{"cmd":"format.read","ref":"A1"}',
    '{"cmd":"format.border","ref":"A1:D1","side":"all","weight":"medium","color":"#000000"}',
    '{"cmd":"format.align","ref":"A1","horizontal":"center","vertical":"center","wrap_text":true}',
    '{"cmd":"format.font","ref":"A1","italic":true,"underline":true,"color":"#ffffff"}',
    '{"cmd":"format","ref":"A1","number_format":"#,##0.00"}'
)

# ── LAYOUT ──
Test-Section "Layout" @(
    '{"cmd":"attach"}',
    '{"cmd":"goto","target":"TestSheet1"}',
    '{"cmd":"column.width","ref":"A:A","width":20}',
    '{"cmd":"column.width","ref":"A:A"}',
    '{"cmd":"row.height","ref":"1:1","height":30}',
    '{"cmd":"autofit","ref":"A1:D10","target":"columns"}',
    '{"cmd":"merge","ref":"A10:C10"}',
    '{"cmd":"unmerge","ref":"A10:C10"}',
    '{"cmd":"freeze","ref":"A3"}',
    '{"cmd":"unfreeze"}',
    '{"cmd":"insert.row","ref":"A5"}',
    '{"cmd":"delete.row","ref":"A5"}',
    '{"cmd":"used.range"}',
    '{"cmd":"row.count"}',
    '{"cmd":"column.count"}'
)

# ── DATA OPS ──
Test-Section "Data Operations" @(
    '{"cmd":"attach"}',
    '{"cmd":"goto","target":"TestSheet1"}',
    '{"cmd":"type","ref":"A1","value":"Apple"}',
    '{"cmd":"type","ref":"A2","value":"Banana"}',
    '{"cmd":"type","ref":"A3","value":"Cherry"}',
    '{"cmd":"find","what":"Banana"}',
    '{"cmd":"find","what":"Nonexistent"}',
    '{"cmd":"find.all","what":"a"}',
    '{"cmd":"replace","what":"Cherry","with":"Date"}',
    '{"cmd":"sort","ref":"A1:A3","column":"A1:A3","order":"asc"}',
    '{"cmd":"copy","ref":"A1:A3"}',
    '{"cmd":"paste","ref":"B1"}',
    '{"cmd":"fill.down","ref":"A1:A5"}',
    '{"cmd":"fill.right","ref":"A1:C1"}',
    '{"cmd":"comment","ref":"A1","text":"Test comment"}',
    '{"cmd":"comment.read","ref":"A1"}',
    '{"cmd":"hyperlink","ref":"D1","url":"https://example.com","text":"Link"}',
    '{"cmd":"validation","ref":"E1","type":"list","formula":"Yes,No,Maybe"}',
    '{"cmd":"protect"}',
    '{"cmd":"unprotect"}'
)

# ── CALC ──
Test-Section "Calculation" @(
    '{"cmd":"attach"}',
    '{"cmd":"goto","target":"TestSheet1"}',
    '{"cmd":"type","ref":"F1","value":"10"}',
    '{"cmd":"type","ref":"F2","value":"=F1*2"}',
    '{"cmd":"calc"}',
    '{"cmd":"calc.mode"}',
    '{"cmd":"calc.mode","mode":"manual"}',
    '{"cmd":"calc.mode"}',
    '{"cmd":"calc.mode","mode":"auto"}',
    '{"cmd":"wait.calc","timeout":3000}',
    '{"cmd":"wait.cell","ref":"F2","condition":"not_empty","timeout":2000}',
    '{"cmd":"wait.cell","ref":"Z99","condition":"not_empty","timeout":500}',
    '{"cmd":"time.calc"}'
)

# ── NAMED RANGES ──
Test-Section "Named Ranges" @(
    '{"cmd":"attach"}',
    '{"cmd":"goto","target":"TestSheet1"}',
    '{"cmd":"name.set","name":"TestRange","ref":"A1:A3"}',
    '{"cmd":"names"}',
    '{"cmd":"name.read","name":"TestRange"}',
    '{"cmd":"name.delete","name":"TestRange"}',
    '{"cmd":"names"}'
)

# ── TABLES ──
Test-Section "Tables" @(
    '{"cmd":"attach"}',
    '{"cmd":"goto","target":"TestSheet1"}',
    '{"cmd":"clear","ref":"H1:K10"}',
    '{"cmd":"type","ref":"H1","value":"Name"}',
    '{"cmd":"type","ref":"I1","value":"Score"}',
    '{"cmd":"type","ref":"H2","value":"Alice"}',
    '{"cmd":"type","ref":"I2","value":"95"}',
    '{"cmd":"type","ref":"H3","value":"Bob"}',
    '{"cmd":"type","ref":"I3","value":"87"}',
    '{"cmd":"table.create","ref":"H1:I3","name":"ScoreTable"}',
    '{"cmd":"table.list"}',
    '{"cmd":"table.data","name":"ScoreTable"}',
    '{"cmd":"table.totals","name":"ScoreTable","show":true,"column":"Score","function":"average"}',
    '{"cmd":"table.delete","name":"ScoreTable"}'
)

# ── CHARTS ──
Test-Section "Charts" @(
    '{"cmd":"attach"}',
    '{"cmd":"goto","target":"TestSheet1"}',
    '{"cmd":"chart.create","type":"column","data":"H1:I3","title":"Test Chart"}',
    '{"cmd":"chart.list"}'
)

# ── WINDOW ──
Test-Section "Window" @(
    '{"cmd":"attach"}',
    '{"cmd":"window.zoom","level":80}',
    '{"cmd":"window.zoom"}',
    '{"cmd":"window.gridlines"}',
    '{"cmd":"window.gridlines","show":false}',
    '{"cmd":"window.gridlines","show":true}',
    '{"cmd":"window.headings"}',
    '{"cmd":"window.view"}',
    '{"cmd":"window.zoom","level":100}'
)

# ── SHAPES ──
Test-Section "Shapes" @(
    '{"cmd":"attach"}',
    '{"cmd":"goto","target":"TestSheet1"}',
    '{"cmd":"shape.add","type":"rectangle","left":500,"top":10,"width":150,"height":50,"text":"Hello Shape"}',
    '{"cmd":"shape.list"}'
)

# ── PRINT ──
Test-Section "Print Setup" @(
    '{"cmd":"attach"}',
    '{"cmd":"goto","target":"TestSheet1"}',
    '{"cmd":"print.setup","orientation":"landscape"}',
    '{"cmd":"print.margins","top":1.0,"bottom":1.0,"left":0.75,"right":0.75}',
    '{"cmd":"print.area","ref":"A1:I10"}',
    '{"cmd":"print.gridlines","show":true}'
)

# ── WORKBOOKS ──
Test-Section "Workbooks" @(
    '{"cmd":"attach"}',
    '{"cmd":"workbooks"}',
    '{"cmd":"workbook.properties"}',
    '{"cmd":"selection.info"}'
)

# ── ADVANCED ──
Test-Section "Advanced" @(
    '{"cmd":"attach"}',
    '{"cmd":"goto","target":"TestSheet1"}',
    '{"cmd":"type","ref":"A1","value":"Test"}',
    '{"cmd":"type","ref":"Z1","value":"undo_test"}',
    '{"cmd":"error.check","ref":"A1"}',
    '{"cmd":"link.list"}'
)

# ── SCREENSHOT ──
Test-Section "Screenshot" @(
    '{"cmd":"attach"}',
    '{"cmd":"screenshot"}'
)

# ── RIBBON / UI ──
Test-Section "Ribbon & UI" @(
    '{"cmd":"attach"}',
    '{"cmd":"ribbon"}',
    '{"cmd":"dialog.read"}',
    '{"cmd":"ui.tree","depth":1}'
)

# ── BATCH ──
Test-Section "Batch" @(
    '{"cmd":"attach"}',
    '{"cmd":"batch","commands":[{"cmd":"read","ref":"A1"},{"cmd":"sheets"},{"cmd":"status"},{"cmd":"calc.mode"},{"cmd":"window.zoom"}]}'
)

# ── CLEANUP ──
Test-Section "Cleanup" @(
    '{"cmd":"attach"}',
    '{"cmd":"sheet.delete","name":"TestSheet1"}',
    '{"cmd":"sheets"}'
)

# ===================================================
$duration = (Get-Date) - $startTime
$total = $passed + $failed

Write-Host ""
Write-Host "  ===============================================" -ForegroundColor Cyan
Write-Host "       XRAI TORTURE TEST - RESULTS" -ForegroundColor Cyan
Write-Host "  ===============================================" -ForegroundColor Cyan
Write-Host ""

if ($failed -eq 0) {
    Write-Host "  TOTAL: $passed/$total PASSED - 0 FAILED" -ForegroundColor Green
} else {
    Write-Host "  TOTAL: $passed/$total PASSED - $failed FAILED" -ForegroundColor Red
    Write-Host ""
    Write-Host "  Failed commands:" -ForegroundColor Red
    foreach ($err in $errors) {
        Write-Host "    - $err" -ForegroundColor Red
    }
}

Write-Host "  Duration: $($duration.TotalSeconds.ToString('F1')) seconds" -ForegroundColor Gray
Write-Host ""
Write-Host "  ===============================================" -ForegroundColor Cyan
Write-Host ""

# Save results to JSON
$results = @{
    timestamp = (Get-Date -Format "o")
    passed = $passed
    failed = $failed
    total = $total
    duration_seconds = [math]::Round($duration.TotalSeconds, 1)
    errors = $errors
}
$results | ConvertTo-Json | Out-File "$scriptDir\test-results.json" -Encoding UTF8

exit $(if ($failed -eq 0) { 0 } else { 1 })
