# Testing & Assertions Guide

XRai includes a built-in testing framework with assertions, screenshot comparison, OCR, and report generation. This enables fully automated testing of Excel add-ins without any human interaction.

## Assertions

### assert.cell -- Assert Cell Values

Verify that a cell contains an expected value, formula, or format:

```json
{"cmd":"assert.cell","ref":"A1","condition":"equals","value":"Hello"}
{"cmd":"assert.cell","ref":"B1","condition":"equals","value":150}
{"cmd":"assert.cell","ref":"C1","condition":"not_empty"}
{"cmd":"assert.cell","ref":"D1","condition":"not_error"}
{"cmd":"assert.cell","ref":"E1","condition":"formula","value":"=SUM(A:A)"}
```

Returns `{"ok":true,"passed":true}` on success, or `{"ok":true,"passed":false,"expected":...,"actual":...}` on failure.

### assert.pane -- Assert Control State

Verify that a task pane control has an expected value or state:

```json
{"cmd":"assert.pane","control":"StatusLabel","property":"value","value":"Ready"}
{"cmd":"assert.pane","control":"CalcButton","property":"enabled","value":true}
{"cmd":"assert.pane","control":"ErrorCheck","property":"checked","value":false}
```

### assert.model -- Assert ViewModel Properties

Verify that a ViewModel property has an expected value:

```json
{"cmd":"assert.model","property":"Spot","value":105.0}
{"cmd":"assert.model","property":"Status","value":"Calculated"}
```

## Test Runs

Structure tests into named runs with individual steps for organized reporting.

### test.start -- Begin a Test Run

```json
{"cmd":"test.start","name":"Portfolio Smoke Test"}
```

### test.step -- Record a Step

```json
{"cmd":"test.step","name":"Verify initial state","status":"pass"}
{"cmd":"test.step","name":"Click refresh button","status":"pass"}
{"cmd":"test.step","name":"Check updated values","status":"fail","message":"Expected 150, got 0"}
```

Status values: `pass`, `fail`, `skip`.

### test.end -- End the Test Run

```json
{"cmd":"test.end"}
```

### test.report -- Generate a Report

```json
{"cmd":"test.report","format":"html","path":"C:\\Reports\\test-results.html"}
{"cmd":"test.report","format":"junit-xml","path":"C:\\Reports\\test-results.xml"}
```

Supported formats:

- `html` -- visual HTML report with pass/fail summary and step details
- `junit-xml` -- JUnit XML format compatible with CI/CD systems (Azure DevOps, GitHub Actions, Jenkins)

## Screenshot Comparison

Visual regression testing by comparing screenshots against baselines.

### screenshot.baseline -- Save a Baseline

```json
{"cmd":"screenshot.baseline","name":"portfolio-grid","path":"C:\\Baselines"}
```

Captures the current Excel window and saves it as a named baseline image.

### screenshot.compare -- Compare Against Baseline

```json
{"cmd":"screenshot.compare","name":"portfolio-grid","threshold":5}
```

Compares the current state against the saved baseline. The `threshold` parameter sets the maximum allowed pixel difference percentage (default: 5%). Returns:

```json
{
  "ok": true,
  "match": true,
  "difference_percent": 1.2,
  "threshold": 5,
  "baseline_path": "C:\\Baselines\\portfolio-grid.png",
  "current_path": "C:\\Baselines\\portfolio-grid-current.png"
}
```

## OCR

Extract text from the screen or specific elements using optical character recognition.

### ocr.screen -- OCR the Excel Window

```json
{"cmd":"ocr.screen"}
{"cmd":"ocr.screen","region":{"x":100,"y":200,"width":400,"height":300}}
```

Returns recognized text from the entire window or a specific region.

### ocr.element -- OCR a Specific Element

```json
{"cmd":"ocr.element","ref":"A1:D10"}
```

Captures and OCRs just the specified range or element.

## Wait Commands

Block until a condition is met, with configurable timeouts.

### wait.element -- Wait for UI Element

```json
{"cmd":"wait.element","query":"StatusBar","timeout":5000}
```

### wait.window -- Wait for Window

```json
{"cmd":"wait.window","title":"Microsoft Excel","timeout":10000}
```

### wait.property -- Wait for Property Value

```json
{"cmd":"wait.property","control":"StatusLabel","property":"value","value":"Done","timeout":5000}
```

Equivalent to `pane.wait`.

### wait.gone -- Wait for Element to Disappear

```json
{"cmd":"wait.gone","query":"LoadingSpinner","timeout":10000}
```

### wait.idle -- Wait for Application Idle

```json
{"cmd":"wait.idle","timeout":5000}
```

## Example: Full Smoke Test

A complete automated smoke test for an add-in:

```json
{"cmd":"batch","commands":[
  {"cmd":"test.start","name":"Portfolio Add-In Smoke Test"},

  {"cmd":"connect"},
  {"cmd":"test.step","name":"Connected to Excel","status":"pass"},

  {"cmd":"pane"},
  {"cmd":"test.step","name":"Pane controls discovered","status":"pass"},

  {"cmd":"pane.type","control":"TickerInput","value":"AAPL"},
  {"cmd":"pane.click","control":"AddButton"},
  {"cmd":"pane.wait","control":"StatusLabel","property":"value","value":"Added","timeout":5000},
  {"cmd":"test.step","name":"Added AAPL to portfolio","status":"pass"},

  {"cmd":"assert.pane","control":"HoldingsGrid","property":"rows","value":1},
  {"cmd":"test.step","name":"Grid shows 1 holding","status":"pass"},

  {"cmd":"pane.click","control":"RefreshButton"},
  {"cmd":"pane.wait","control":"StatusLabel","property":"value","value":"Refreshed","timeout":10000},
  {"cmd":"assert.cell","ref":"B2","condition":"not_empty"},
  {"cmd":"test.step","name":"Refresh populated cells","status":"pass"},

  {"cmd":"screenshot"},
  {"cmd":"test.step","name":"Screenshot captured","status":"pass"},

  {"cmd":"test.end"},
  {"cmd":"test.report","format":"html","path":"C:\\Reports\\smoke-test.html"}
]}
```

## CI/CD Integration

Use `test.report` with `junit-xml` format to integrate XRai tests into your CI/CD pipeline:

```powershell
# Run tests and generate JUnit XML
echo '{"cmd":"batch","commands":[...]}' | XRai.Tool.exe

# Azure DevOps: publish test results
- task: PublishTestResults@2
  inputs:
    testResultsFormat: 'JUnit'
    testResultsFiles: '**/test-results.xml'
```
