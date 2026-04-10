# Architecture

XRai is a bridge between AI coding agents and live Windows desktop applications. It provides a JSON-over-stdin/stdout protocol that agents use to send commands and receive structured responses.

## Three-Layer Architecture

```
AI Agent (Claude Code / Codex / Cursor)
    |
    |  pipes JSON commands on stdin
    v
XRai.Tool.exe (CLI)
    |
    |  routes commands to the appropriate driver
    v
+----------------+-----------------+------------------+
| XRai.Com       | XRai.Hooks      | XRai.UI          |
| (COM interop)  | (named pipe)    | (FlaUI + Win32)  |
|                |                 |                  |
| cells,         | WPF/WinForms    | ribbon, dialogs, |
| formulas,      | controls,       | screenshots,     |
| charts,        | ViewModel,      | folder pickers,  |
| tables,        | UDFs            | UI tree          |
| pivots, VBA    |                 |                  |
+-------+--------+--------+--------+---------+--------+
        |                 |                  |
        +-----------------+------------------+
                    Excel / Windows App
```

### COM Layer (XRai.Com)

Direct COM interop with Excel via `Microsoft.Office.Interop.Excel`. Handles cells, formulas, formatting, charts, tables, pivots, filters, shapes, print setup, workbook management, and more. This is the most comprehensive layer with 130+ commands.

All COM objects are tracked via `ComGuard`, an IDisposable wrapper that:

- Records every COM object handed to it
- Releases them all via `Marshal.ReleaseComObject` on Dispose
- Releases in reverse order of acquisition
- Calls `GC.Collect` and `GC.WaitForPendingFinalizers` on session close

This prevents zombie Excel processes. Every COM interaction must use ComGuard -- never write two-dot expressions that create untracked intermediate objects.

### Hooks Layer (XRai.Hooks)

An in-process library that ships as a NuGet package (`XRai.Hooks`). Add-in developers reference it and call three methods in `AutoOpen()`:

- `Pilot.Start()` -- starts a named pipe server inside the add-in process
- `Pilot.Expose(element)` -- walks the WPF/WinForms visual tree and registers named controls
- `Pilot.ExposeModel(vm)` -- exposes `INotifyPropertyChanged` properties for direct read/write

The pipe name is `xrai_{excel_pid}`, making it deterministic and auto-discoverable. XRai.Tool connects as a named pipe client.

The hooks layer has zero heavyweight dependencies -- just `System.IO.Pipes` and `System.Text.Json`. It adds minimal overhead to the add-in.

### UI Layer (XRai.UI)

FlaUI (UIA3) automation for native Excel UI elements that cannot be reached through COM or hooks: ribbon tabs, ribbon buttons, modal dialogs, folder pickers, and the UI automation tree.

Win32 APIs (`EnumWindows`, `SendMessage`, `WM_SETTEXT`) handle native dialogs that UIA cannot reach, including `#32770` class dialogs and WinForms message boxes.

This layer also provides screenshot capture via Win32 `PrintWindow`.

## JSON Protocol

### Transport

One JSON object per line on stdin (commands) and stdout (responses). No framing protocol needed -- newline delimits messages.

### Command Format

```json
{"cmd": "read", "ref": "A1"}
{"cmd": "pane.click", "control": "CalcButton"}
{"cmd": "batch", "commands": [{"cmd": "status"}, {"cmd": "sheets"}]}
```

### Response Format

Every response includes an `ok` field:

```json
{"ok": true, "value": 42.5}
{"ok": false, "error": "Cell reference out of range: ZZ99999"}
```

### Event Format

Events are pushed to stdout without being requested. They have an `event` field instead of `cmd`:

```json
{"event": "error", "exception": "NullReferenceException", "message": "...", "stack": "...", "timestamp": "..."}
{"event": "log", "message": "Calc started", "source": "Debug", "timestamp": "..."}
{"event": "functions_registered", "functions": [...], "timestamp": "..."}
```

## Threading Model

### COM Thread

Excel COM interop requires an STA (Single-Threaded Apartment) thread. XRai runs all COM operations on a dedicated STA worker thread. The `IOleMessageFilter` is installed to handle RPC retry/rejection when Excel is busy.

### Hooks Thread

All hooks operations marshal to the WPF UI thread via `Dispatcher.Invoke`. This is mandatory because WPF controls can only be accessed from the thread that created them. The named pipe server runs on a background thread and dispatches to the UI thread for each command.

### Event Forwarding

Events (errors, logs, function registrations) flow from the hooks pipe server to XRai and are forwarded to stdout. The agent sees them interleaved with command responses.

## Named Pipe Protocol

Pipe name: `xrai_{pid}` where `pid` is the Excel process ID.

The hooks library (inside the add-in) runs a `NamedPipeServerStream`. XRai connects as a `NamedPipeClientStream`.

Messages are newline-delimited JSON, using the same format as the stdin/stdout protocol.

Command flow: `XRai.Tool -> pipe -> Hooks (execute on UI thread) -> pipe -> XRai.Tool -> stdout`

Event flow: `Hooks -> pipe -> XRai.Tool -> stdout (forwarded to agent)`

## The Skill System

XRai is distributed as a Claude Code skill -- a directory structure that Claude Code auto-discovers:

```
~/.claude/skills/xrai-excel/
  SKILL.md            Instructions Claude Code reads on every session
  bin/XRai.Tool.exe   Self-contained CLI binary
  packages/           XRai.Hooks NuGet package
  templates/          Project scaffolding templates
  docs/               Reference documentation
```

The `SKILL.md` file tells Claude Code how to use XRai: when to attach, how to send commands, how to interpret responses, and common patterns for add-in development.

## The MCP Server

XRai.Mcp exposes XRai commands as an MCP (Model Context Protocol) server. This enables agents that support MCP -- Codex, Cursor, Windsurf, VS Code Copilot -- to use XRai without the skill system.

The MCP server wraps the same command router used by the CLI, so all 283 commands are available through both interfaces.

## Project Structure

```
src/
  XRai.Core/          Command router, JSON protocol, response helpers, event stream
  XRai.Com/           Excel COM interop (17 Ops files, 130+ commands)
  XRai.Hooks/         In-process NuGet library (Pilot, PipeServer, ControlAdapter)
  XRai.HooksClient/   Named pipe client for hooks communication
  XRai.UI/            FlaUI ribbon/dialog driver, Win32 dialog driver
  XRai.Vision/        Win32 PrintWindow screenshot capture, OCR
  XRai.Tool/          CLI entry point, doctor, init, reload orchestrator
  XRai.Mcp/           MCP server for Codex/Cursor/Windsurf
tests/
  XRai.Tests.Unit/    Unit tests (no Excel required)
demo/
  XRai.Demo.PortfolioAddin/   Sample add-in with 30+ controls
  XRai.Demo.MacroGuard/       VBA management add-in (13 tabs, 50+ controls)
```

## Design Decisions

- **Target `net8.0-windows`** -- Windows APIs required for COM, WPF, and Win32
- **System.Text.Json** everywhere -- no Newtonsoft dependency
- **One JSON per line** -- simple, no framing protocol needed
- **Flat, terse commands** -- `{"cmd":"read","ref":"A1"}` not deeply nested structures
- **Responses always have `ok` field** -- easy programmatic parsing
- **Events have `event` field** -- easily distinguished from command responses
- **Deterministic pipe name** -- `xrai_{pid}` enables auto-discovery
- **All hooks operations marshal to UI thread** -- mandatory for WPF thread safety
- **XRai.Hooks has zero heavyweight dependencies** -- just System.IO.Pipes and System.Text.Json
- **FlaUI is fallback, hooks are primary** -- route to hooks first, FlaUI only for native UI
- **Self-contained publish** -- no .NET SDK required on target machines
