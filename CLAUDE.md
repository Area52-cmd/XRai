# XRai — Excel Automation for Claude Code

You have access to XRai, a CLI tool that gives you full control over a live Excel instance. Send JSON commands on stdin, get JSON responses on stdout.

## Running XRai

```bash
# Attach to running Excel and start sending commands:
echo '{"cmd":"attach"}' | path/to/XRai.Tool.exe

# Or pipe multiple commands:
echo '{"cmd":"attach"}
{"cmd":"read","ref":"A1:D10"}
{"cmd":"sheets"}' | path/to/XRai.Tool.exe
```

## Essential Commands

### Read/Write Cells
```json
{"cmd":"read","ref":"A1"}                    → value
{"cmd":"read","ref":"A1:D10"}                → all cells in range
{"cmd":"read","ref":"A1","full":true}        → value + formula + format
{"cmd":"type","ref":"A1","value":"Hello"}    → write string
{"cmd":"type","ref":"A1","value":"=SUM(B:B)"} → write formula
{"cmd":"clear","ref":"A1:D10"}              → clear range
```

### Sheets
```json
{"cmd":"sheets"}                             → list all sheets
{"cmd":"sheet.add","name":"NewSheet"}        → add sheet
{"cmd":"goto","target":"Sheet2"}             → activate sheet
```

### Formatting
```json
{"cmd":"format","ref":"A1","bold":true,"font_size":14,"bg":"#ff0000","number_format":"$#,##0.00"}
{"cmd":"format.border","ref":"A1:D1","side":"all","weight":"medium","color":"#000000"}
```

### Charts & Tables
```json
{"cmd":"chart.create","type":"column","data":"A1:B10","title":"My Chart"}
{"cmd":"table.create","ref":"A1:D10","name":"MyTable"}
{"cmd":"table.data","name":"MyTable"}        → read as JSON array
```

### Task Pane (if add-in has XRai.Hooks)
```json
{"cmd":"pane"}                               → list all controls
{"cmd":"pane.click","control":"CalcButton"}  → click button
{"cmd":"pane.type","control":"Input","value":"100"} → type in textbox
{"cmd":"pane.read","control":"Result"}       → read control value
{"cmd":"pane.tab","control":"Tabs","tab":"Settings"} → switch tab
{"cmd":"pane.grid.read","control":"DataGrid"} → read grid data
```

### ViewModel
```json
{"cmd":"model"}                              → all ViewModel properties
{"cmd":"model.set","property":"Spot","value":105.0}
{"cmd":"functions"}                          → list registered UDFs
```

### Screenshots & UI
```json
{"cmd":"screenshot"}                         → capture Excel as PNG
{"cmd":"ribbon"}                             → list ribbon tabs
{"cmd":"ui.tree","depth":2}                  → UI automation tree
```

## Token-Saving Rules

1. **Use `batch`**: `{"cmd":"batch","commands":[{...},{...},{...}]}` — one round-trip
2. **Read ranges**: `A1:Z50` not individual cells
3. **Use `pane` once** to discover controls, then target by name
4. **Use `model`** for full state in one call
5. **Skip `full` flag** unless you need formula/format details

## Response Format

All: `{"ok":true,...data...}` or `{"ok":false,"error":"message"}`
