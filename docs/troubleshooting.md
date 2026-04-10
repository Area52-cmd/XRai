# Troubleshooting

## Error Table

| Error | Cause | Fix |
|-------|-------|-----|
| `GetActiveObject failed` / `MK_E_UNAVAILABLE` | Excel not running or on start screen | Start Excel, open a workbook, then `{"cmd":"connect"}` |
| `No active workbook` | Excel running but no workbook open | `{"cmd":"connect"}` auto-creates a workbook |
| `hooks: false` | Stale .xll or `Pilot.Start()` not called | Recovery B |
| `Control not found: X` | Missing `x:Name` or `Expose` not called | Recovery C |
| `RPC server unavailable` (0x800706BA) | Excel crashed or hung | Recovery A |
| Empty `controls` array | `Pilot.Expose()` not called | Recovery C |
| `workbook.open` hangs | Unhandled NUIDialog pattern | Recovery D |
| OLE "waiting for another application" | Cross-app OLE contention | `{"cmd":"excel.autodismiss","enabled":true}` |
| Multiple Excel processes | Zombies from previous runs | Recovery E |
| `File is being used by another process` | Excel has .xll locked | Kill Excel first, then rebuild |
| `The calling thread cannot access this object` | WPF dispatcher threading issue | Ensure `Pilot.Expose()` runs on the WPF UI thread |
| `model` returns "No model exposed" | `Pilot.ExposeModel()` not called | Add `Pilot.ExposeModel(vm)` to `AutoOpen()` |
| NuGet "Package not compatible with net8.0" | Add-in targets `net8.0` instead of `net8.0-windows` | Change TargetFramework in `.csproj` |
| `PackAsTool does not support TargetPlatformIdentifier` | Trying to pack a Windows project as dotnet tool | XRai ships as self-contained binary, not a dotnet tool |
| `.xll loaded: true` but pane is empty | Document Recovery pane blocking macro execution | Dismiss Document Recovery, then reload .xll |

## Recovery A: Excel Not Responding

1. Check if Excel is running: `tasklist | grep -i excel`
2. If hung, kill it: `XRai.Tool.exe kill-excel`
3. Wait 2 seconds for cleanup
4. Start Excel: `start excel`
5. Wait for Excel to initialize (5 seconds)
6. Reconnect: `{"cmd":"connect"}`

## Recovery B: Hooks Not Connecting

The `hooks: false` status means the XRai.Hooks named pipe server is not running inside Excel. This happens when the add-in is not loaded or is running a stale version.

1. Kill Excel to release the DLL lock: `XRai.Tool.exe kill-excel`
2. Rebuild the add-in: `dotnet build`
3. Start Excel: `start excel`
4. Load the `.xll` add-in
5. Reconnect: `{"cmd":"connect"}` -- verify `hooks: true`

Key rule: Excel MUST be killed before rebuilding, or the DLL lock prevents the build from replacing the output files.

## Recovery C: Empty Pane Controls

The `pane` command returns controls but the list is empty.

1. Check `AutoOpen()` in your `IExcelAddIn` class
2. Verify `Pilot.Expose(taskPane)` is called after the pane is created
3. If the pane is created inside `ExcelAsyncUtil.QueueAsMacro`, the `Expose` call must be in the same callback:

```csharp
ExcelAsyncUtil.QueueAsMacro(() =>
{
    var pane = new MyTaskPane();
    // ... set up and display the pane ...
    Pilot.Expose(pane);  // Must be here, after pane exists
});
```

4. Rebuild and reload using Recovery B steps

## Recovery D: workbook.open Hangs

This is rare because `workbook.open` has a built-in NUIDialog watchdog that auto-dismisses common dialogs. If it still hangs:

1. Open a second XRai shell (daemon mode supports multiple clients)
2. List dialogs: `{"cmd":"win32.dialog.list"}`
3. Inspect the blocking dialog: `{"cmd":"win32.dialog.read","title":"<fragment>"}`
4. Dismiss it: `{"cmd":"win32.dialog.click","button":"<safest button>","title":"<fragment>"}`

Safe-button priority: `Don't Update` > `No` > `Close` > `Cancel` > `OK` > `Yes` > `Continue` > `Enable Editing`

## Recovery E: Multiple Excel Zombies

Zombie Excel processes are caused by unreleased COM objects.

1. Kill all Excel processes: `XRai.Tool.exe kill-excel`
2. Wait 2 seconds
3. Start fresh: `start excel`
4. Reconnect: `{"cmd":"connect"}`

## Version Hygiene

Use `Version="1.0.*"` in your `.csproj` PackageReference for XRai.Hooks:

```xml
<PackageReference Include="XRai.Hooks" Version="1.0.*" />
```

The `rebuild` command handles NuGet cache clearing, restore, and version resolution automatically. If `connect` returns `hooks_stale: true`, the loaded hooks DLL is outdated -- rebuild the add-in.

## Daemon Mode

For faster command execution, run XRai as a persistent daemon:

```bash
XRai.Tool.exe --daemon
```

Benefits: no OLE attach races, no per-command startup cost, hooks pipe stays alive, approximately 50ms per command.

| Command | Description |
|---------|-------------|
| `XRai.Tool.exe daemon-status` | Check if daemon is running |
| `XRai.Tool.exe daemon-stop` | Stop the daemon |
| `XRai.Tool.exe --no-daemon` | Force direct mode (bypass daemon) |

Daemon staleness is auto-detected: if the running daemon was built from an older binary, the next CLI invocation auto-stops it and falls back to direct mode.

## General Debugging

Dump the full state in one call:

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

This returns connection state, all sheets, open workbooks, data bounds, pane controls, ViewModel properties, registered UDFs, and a screenshot -- everything needed to diagnose most issues.
