# Troubleshooting Reference

## Error Table

| Error | Cause | Fix |
|-------|-------|-----|
| `GetActiveObject failed` / `MK_E_UNAVAILABLE` | Excel not running | `start excel`, then `{"cmd":"connect"}` |
| `No active workbook` | Excel on start screen | `{"cmd":"connect"}` auto-creates a workbook |
| `hooks: false` | Stale .xll or XRai.Hooks not wired | Recovery B |
| `Control not found: X` | Missing `x:Name` or Expose not called | Recovery C |
| `RPC server unavailable` (0x800706BA) | Excel crashed or hung | Recovery A |
| Empty `controls` array | `Pilot.Expose` not called | Recovery C |
| `workbook.open` hangs | Unknown NUIDialog pattern | Recovery D |
| OLE "waiting for another application" | Cross-app OLE contention | `{"cmd":"excel.autodismiss","enabled":true}` |
| Multiple Excel processes | Zombies from previous runs | Recovery E |
| `File is being used by another process` | Excel has .xll locked | `kill-excel` first, then rebuild |

## Recovery A: connect fails / Excel not responding

1. `tasklist | grep -i excel`
2. If hung: `XRai.Tool.exe kill-excel`
3. Wait 2 seconds
4. `start excel`
5. Wait 5 seconds
6. `{"cmd":"connect"}`

## Recovery B: hooks: false after adding XRai.Hooks

1. `XRai.Tool.exe kill-excel` — release .xll lock
2. `dotnet build` — rebuild with Excel closed
3. `start excel`
4. User loads .xll (or MSBuild auto-launch target does it)
5. `{"cmd":"connect"}` — verify `hooks: true`

**Key rule**: Excel MUST be killed before rebuilding or the DLL lock blocks the build.

## Recovery C: pane returns empty controls

1. Check `AutoOpen()` in the `IExcelAddIn` class
2. Verify `Pilot.Expose(taskPane)` is called after pane is created
3. If pane is inside `ExcelAsyncUtil.QueueAsMacro`, Expose must be in the same callback
4. See `templates/pilot-wiring.cs` for patterns
5. Rebuild and reload (Recovery B steps)

## Recovery D: workbook.open hangs

**Usually not needed** — `workbook.open` has built-in NUIDialog defense with auto-dismiss watchdog.

If it still hangs:

1. Open a second XRai shell (daemon mode: same pipe, new client)
2. `{"cmd":"win32.dialog.list"}` — find the blocking dialog
3. `{"cmd":"win32.dialog.read","title":"<fragment>"}` — inspect structure
4. `{"cmd":"win32.dialog.click","button":"<safest button>","title":"<fragment>"}` — dismiss
5. Record title + class_name + buttons for watchdog extension

Safe-button priority: `Don't Update` > `No` > `Close` > `Cancel` > `OK` > `Yes` > `Continue` > `Enable Editing`

## Recovery E: Multiple Excel zombies

1. `XRai.Tool.exe kill-excel`
2. Wait 2 seconds
3. `start excel`
4. `{"cmd":"connect"}`

## Version hygiene

Use `Version="1.0.0-*"` in your `.csproj` PackageReference for XRai.Hooks. The `rebuild` command handles NuGet cache clearing, restore, and version resolution automatically. If `connect` returns `hooks_stale: true`, the loaded hooks DLL is outdated — rebuild the add-in.

## Daemon mode

```bash
"$HOME/.claude/skills/xrai-excel/bin/XRai.Tool.exe" --daemon
```

Benefits: no OLE races, no per-command attach cost, hooks pipe stays alive, ~50ms per command.

| Command | Description |
|---|---|
| `XRai.Tool.exe daemon-status` | Check if daemon is running |
| `XRai.Tool.exe daemon-stop` | Stop the daemon |
| `XRai.Tool.exe --no-daemon` | Force direct mode (bypass daemon) |

Daemon staleness is auto-detected: if the running daemon was built from an older binary, the next CLI invocation auto-stops it and falls back to direct mode.
