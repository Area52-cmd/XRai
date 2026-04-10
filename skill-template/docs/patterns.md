# Token-Efficient Patterns for XRai

XRai is designed for AI agents. Minimize round-trips, maximize information per response.

## Pattern 1: Always Batch

**Bad** (4 round-trips, 4 tokens of overhead each):
```json
{"cmd":"attach"}
{"cmd":"status"}
{"cmd":"sheets"}
{"cmd":"read","ref":"A1:D10"}
```

**Good** (1 round-trip):
```json
{"cmd":"batch","commands":[
  {"cmd":"attach"},
  {"cmd":"status"},
  {"cmd":"sheets"},
  {"cmd":"read","ref":"A1:D10"}
]}
```

## Pattern 2: Read Ranges, Not Cells

**Bad** (100 cells = 100 responses):
```json
{"cmd":"read","ref":"A1"}
{"cmd":"read","ref":"A2"}
...
```

**Good** (1 response, all 100 values):
```json
{"cmd":"read","ref":"A1:A100"}
```

## Pattern 3: Discover Once, Target by Name

**Bad** (guessing control names):
```json
{"cmd":"pane.click","control":"RefreshButton"}  // might fail
{"cmd":"pane.click","control":"refresh_btn"}   // might fail
{"cmd":"pane.click","control":"BtnRefresh"}    // might fail
```

**Good** (discover all controls once, then target):
```json
{"cmd":"pane"}
// → full list of controls with names
{"cmd":"pane.click","control":"RefreshButton"}  // now you know the real name
```

## Pattern 4: Use `model` for State Snapshots

Instead of reading individual properties:
```json
{"cmd":"pane.read","control":"SpotInput"}
{"cmd":"pane.read","control":"VolInput"}
{"cmd":"pane.read","control":"ResultLabel"}
{"cmd":"pane.read","control":"StatusLabel"}
```

Use one call that returns everything (including collections):
```json
{"cmd":"model"}
// → {"ok":true,"properties":{"Spot":100,"Vol":0.25,"Result":42.5,"Status":"Ready","Trades":[...]}}
```

## Pattern 5: Skip `full` Unless Needed

**Bad** (fetches formula + font + format for every cell):
```json
{"cmd":"read","ref":"A1:Z100","full":true}
```

**Good** (values only, lighter response):
```json
{"cmd":"read","ref":"A1:Z100"}
```

Only use `full:true` when you specifically need formulas or formatting.

## Pattern 6: Orient Before Acting

Always start with a state snapshot:
```json
{"cmd":"batch","commands":[
  {"cmd":"attach"},
  {"cmd":"status"},
  {"cmd":"sheets"},
  {"cmd":"used.range"},
  {"cmd":"workbooks"}
]}
```

Now you know:
- Excel version
- Hooks status
- Active sheet + all sheets
- Data bounds
- Workbook count

Proceed from an informed state instead of guessing.

## Pattern 7: Use `wait.cell` for Async Results

**Bad** (polling loop):
```json
{"cmd":"read","ref":"A1"}  // empty
{"cmd":"read","ref":"A1"}  // empty
{"cmd":"read","ref":"A1"}  // empty
{"cmd":"read","ref":"A1"}  // finally populated
```

**Good** (single blocking wait):
```json
{"cmd":"wait.cell","ref":"A1","condition":"not_empty","timeout":10000}
```

## Pattern 8: Calc Mode for Performance

Running many updates? Switch to manual calc:
```json
{"cmd":"batch","commands":[
  {"cmd":"calc.mode","mode":"manual"},
  {"cmd":"type","ref":"A1","value":"=SUM(B:B)"},
  {"cmd":"type","ref":"A2","value":"=AVERAGE(B:B)"},
  {"cmd":"type","ref":"A3","value":"=MAX(B:B)"},
  {"cmd":"type","ref":"A4","value":"=MIN(B:B)"},
  {"cmd":"calc"},
  {"cmd":"calc.mode","mode":"auto"}
]}
```

## Pattern 9: Use `functions` to List UDFs

Don't guess UDF names or parameters — ask the add-in:
```json
{"cmd":"functions"}
// → [{"name":"PRICE","description":"...","parameters":[...]}, ...]
```

## Pattern 10: Screenshot for Visual Verification

When state is ambiguous, capture the visual:
```json
{"cmd":"screenshot"}
// → {"path":"...","base64":"..."}
```

Attach the image to your analysis when reporting to the user.

## Common Compound Operations

### Build a dashboard
```json
{"cmd":"batch","commands":[
  {"cmd":"attach"},
  {"cmd":"clear","ref":"A1:Z100"},
  {"cmd":"type","ref":"A1","value":"Dashboard Title"},
  {"cmd":"format","ref":"A1:F1","bold":true,"font_size":18,"bg":"#1a1a2e"},
  {"cmd":"merge","ref":"A1:F1"},
  {"cmd":"type","ref":"A3","value":"Header1"},
  {"cmd":"type","ref":"B3","value":"Header2"},
  ...
]}
```

### Run a full add-in smoke test
```json
{"cmd":"batch","commands":[
  {"cmd":"attach"},
  {"cmd":"status"},
  {"cmd":"functions"},
  {"cmd":"pane"},
  {"cmd":"model"},
  {"cmd":"pane.type","control":"InputBox","value":"test"},
  {"cmd":"pane.click","control":"SubmitButton"},
  {"cmd":"pane.read","control":"ResultLabel"},
  {"cmd":"screenshot"}
]}
```

### Diagnose hooks issues
```json
{"cmd":"batch","commands":[
  {"cmd":"status"},
  {"cmd":"workbooks"},
  {"cmd":"functions"},
  {"cmd":"pane"},
  {"cmd":"ui.tree","depth":2}
]}
```
