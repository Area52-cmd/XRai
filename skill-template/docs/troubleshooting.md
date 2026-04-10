# XRai Troubleshooting Guide

## Connection Errors

### `GetActiveObject failed` / `MK_E_UNAVAILABLE`
**Cause**: Excel isn't running, or is sitting on the start screen with no workbook open.

**Fix**: Start Excel and open/create a workbook. Then retry `{"cmd":"attach"}`.

```bash
# Verify Excel is running
tasklist | grep -i excel

# If running but on start screen, use PowerShell to create a workbook:
powershell -Command "[Runtime.InteropServices.Marshal]::GetActiveObject('Excel.Application').Workbooks.Add()"
```

### `No active workbook`
**Cause**: Excel is running but no workbook is open.

**Fix**: Create/open a workbook via COM or have the user click "Blank workbook".

### `RPC server unavailable` (0x800706BA)
**Cause**: Excel crashed or COM connection was lost.

**Fix**: Kill any orphan EXCEL processes, restart Excel, re-attach.
```bash
taskkill //F //IM EXCEL.EXE
start excel
```

## Hooks Errors

### `status` returns `"hooks": false`
**Cause**: Either the add-in isn't loaded, or the add-in doesn't call `Pilot.Start()`.

**Fix**:
1. Verify add-in is loaded: `{"cmd":"functions"}` — if any UDFs return, the .xll is loaded
2. Check `AddInEntry.cs` (or whichever class implements `IExcelAddIn`):
   ```csharp
   public void AutoOpen()
   {
       Pilot.Start();  // ← must be present
       // ...
   }
   ```
3. Rebuild and reload the .xll

### `pane` returns empty `controls: []`
**Cause**: `Pilot.Expose()` was never called on a WPF control, or was called before the visual tree was rendered.

**Fix**: Call `Pilot.Expose()` after the task pane UserControl is fully created. If creating the pane asynchronously, expose inside the same callback:

```csharp
ExcelAsyncUtil.QueueAsMacro(() =>
{
    var pane = new MyTaskPane();
    // ... set up the pane, display it ...
    Pilot.Expose(pane);  // ← here, after pane exists
});
```

### `Control not found: X`
**Cause**: No control with `x:Name="X"` in the exposed visual tree.

**Fix**:
1. Add `x:Name="X"` to the control in XAML
2. Use `{"cmd":"pane"}` to list all discoverable controls and find the actual name
3. Check that the control is inside the `FrameworkElement` passed to `Pilot.Expose()`

### `The calling thread cannot access this object because a different thread owns it`
**Cause**: WPF dispatcher threading issue — trying to read a control from the wrong thread.

**Fix**: This is handled internally by XRai.Hooks' `InvokeOnUI`. If it still happens, ensure `Pilot.Expose()` runs on the same thread that created the WPF controls (typically the Excel UI thread or an STA thread you manage).

### `model` returns `"No model exposed"`
**Cause**: `Pilot.ExposeModel()` was never called.

**Fix**: Add to `AutoOpen()`:
```csharp
Pilot.ExposeModel(myViewModel, "Portfolio");
```

## Build & Deployment Errors

### NuGet install fails: "Package XRai.Hooks is not compatible with net8.0"
**Cause**: Your add-in targets `net8.0` instead of `net8.0-windows`. XRai.Hooks needs Windows APIs.

**Fix**: Update `.csproj`:
```xml
<TargetFramework>net8.0-windows</TargetFramework>
```

### File locked: `XRai.Hooks.dll` cannot be copied
**Cause**: Excel has the add-in loaded, locking the DLL.

**Fix**: Unload the add-in first:
```powershell
$excel = [Runtime.InteropServices.Marshal]::GetActiveObject('Excel.Application')
$excel.RegisterXLL('path\to\your.xll')  # toggle off
```
Or close Excel entirely before rebuilding.

### `PackAsTool does not support TargetPlatformIdentifier being set`
**Cause**: Trying to pack an Excel-related project as a dotnet global tool.

**Fix**: XRai.Tool is distributed via the skill directory (`~/.claude/skills/xrai-excel/bin/`), not as a dotnet tool. Don't add `PackAsTool` to Windows-targeted projects.

## COM Cleanup

### Zombie `EXCEL.EXE` after tests
**Cause**: COM objects weren't released properly.

**Fix**: Kill them, then fix the underlying code to wrap every COM interaction in `ComGuard`:
```bash
taskkill //F //IM EXCEL.EXE
```

Every COM object must be tracked:
```csharp
using var guard = new ComGuard();
var sheet = guard.Track(_session.GetActiveSheet());
var range = guard.Track(sheet.Range["A1"]);
// All released on Dispose, in reverse order
```

## Document Recovery Blocks Add-in Load

### Symptom: `xll loaded: true` but `pane` returns empty and `QueueAsMacro` never fires

**Cause**: Excel's Document Recovery pane is blocking macro execution.

**Fix**:
1. Close the Document Recovery pane via the X button (or via FlaUI):
   ```json
   {"cmd":"ribbon.click","button":"Close"}
   ```
2. Unload and reload the add-in:
   ```powershell
   $excel.RegisterXLL($xllPath)  # off
   $excel.RegisterXLL($xllPath)  # on
   ```

## General Debugging

### See what XRai sees
```json
{"cmd":"status"}                 → connection state
{"cmd":"sheets"}                 → active sheets
{"cmd":"workbooks"}              → open workbooks
{"cmd":"used.range"}             → data bounds on active sheet
{"cmd":"pane"}                   → all WPF controls (if hooks connected)
{"cmd":"model"}                  → all ViewModel properties (if hooks connected)
{"cmd":"functions"}              → registered UDFs (if hooks connected)
{"cmd":"ribbon"}                 → ribbon tabs (FlaUI)
{"cmd":"ui.tree","depth":2}      → full UI automation tree
```

### Capture state for diagnosis
```json
{"cmd":"batch","commands":[
  {"cmd":"status"},
  {"cmd":"sheets"},
  {"cmd":"workbooks"},
  {"cmd":"used.range"},
  {"cmd":"pane"},
  {"cmd":"model"},
  {"cmd":"functions"},
  {"cmd":"screenshot"}
]}
```

One round-trip dumps the entire state.
