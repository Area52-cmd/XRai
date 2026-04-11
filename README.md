# XRai -- AI-Powered Desktop Application Development Kit

> Give any AI agent structured, instant access to running Windows desktop applications.
> Build, test, and ship Excel add-ins without touching a mouse.

## What is XRai?

XRai is a development kit that lets AI coding agents (Claude Code, Codex, Cursor, Windsurf) programmatically control live Windows desktop applications. It provides 283 commands for:

- **Excel automation** -- cells, formulas, charts, tables, pivots, Power Query, DAX, slicers, VBA, conditional formatting, data validation, sparklines, shapes, print layout
- **Task pane control** -- click buttons, type in TextBoxes, read DataGrids, select ListBox items, expand ComboBox dropdowns, toggle checkboxes, switch tabs -- via in-process WPF and WinForms hooks
- **ViewModel binding** -- read and write MVVM ViewModel properties directly, without touching the UI
- **Ribbon & dialog automation** -- click ribbon buttons, dismiss dialogs, navigate folder pickers, drive Win32/WinForms modals
- **Desktop automation** -- clipboard, process management, window control, raw keyboard/mouse, app launch/attach
- **Testing** -- assertions (cell/pane/model), screenshot diff, OCR, intelligent waits, HTML/JUnit XML test reporting
- **Screenshots** -- capture Excel windows, task panes, specific regions

## Quick Start

### Install (one time, any machine)

Download the latest release and extract:

```powershell
# PowerShell
Expand-Archive xrai-excel-skill.zip -DestinationPath "$env:USERPROFILE\.claude\skills\"
```

Or manually extract `xrai-excel-skill.zip` to `~/.claude/skills/xrai-excel/`.

### Use

Open Claude Code (or any AI agent with XRai configured) in your Excel-DNA project and say:

> "Build me a stock portfolio tracker Excel add-in with a dark theme task pane"

Claude scaffolds the project, builds it, launches Excel, connects hooks, interacts with the pane, takes screenshots, and verifies everything works -- all autonomously.

### Watch it happen live with XRai Studio

Run this in any terminal:

```powershell
xrai --studio
```

A browser tab opens to a localhost dashboard showing:

- **A friendly onboarding overlay** that detects your installed IDEs (VS Code, Visual Studio 2022/2026, JetBrains Rider) and asks which one you want to follow your code in.
- **Live agent activity feed** -- every message Claude sends, every file edit, every tool call streams in real time as the session happens. Tails Claude Code's session transcript directly from disk -- zero token cost, zero API impact.
- **Live screenshot of Excel** at 4 fps so you watch the add-in materialize.
- **Auto IDE follow** -- when the agent edits a file, Studio automatically opens that file in your real IDE at the right line. You watch the code land in your own editor. Studio never replaces VS Code / VS 2026 / Rider -- it just points your editor at the right place at the right moment.
- **Auto-attach to Excel** -- the moment you open Excel, Studio attaches and the screenshot panel comes alive. No commands required.
- **File touched panel, build progress panel, app state panel**, all live, all in plain English.

Studio is a passive viewer. It does not edit files, does not invoke any AI model, does not consume any tokens. It tails files Claude Code is already writing. If you turn Studio off, your existing workflow is 100% unchanged.

For headless / RDP use:

```powershell
xrai --studio --no-browser
```

The dashboard URL is printed prominently in the terminal so you can copy-paste it into a browser.

### For existing add-ins

1. Add `XRai.Hooks` to your project: `dotnet add package XRai.Hooks --version "1.0.0-*"`
2. Wire `Pilot.Start()` / `Pilot.Expose(pane)` / `Pilot.ExposeModel(vm)` in your `IExcelAddIn.AutoOpen()`
3. Build, load the `.xll` in Excel, and start commanding via XRai

## How It Works

```
AI Agent (Claude / Codex / Cursor)
    | pipes JSON commands
    v
XRai.Tool.exe (CLI)
    | routes to appropriate driver
    v
+----------------+-----------------+------------------+
| XRai.Com       | XRai.Hooks      | XRai.UI          |
| (COM interop)  | (named pipe)    | (FlaUI + Win32)  |
|                |                 |                  |
| cells,         | WPF/WinForms    | ribbon, dialogs, |
| formulas,      | controls,       | screenshots,     |
| charts...      | ViewModel...    | folder pickers...|
+-------+--------+--------+--------+---------+--------+
        |                 |                  |
        +-----------------+------------------+
                  Excel / Any Windows App
```

## 283 Commands

Full command catalog: [docs/commands.md](docs/commands.md)

### Highlights

| Category | Commands | Examples |
|----------|----------|---------|
| Cells & Data | 30+ | `read`, `type`, `format`, `sort`, `find`, `validation` |
| Charts & Sparklines | 15+ | `chart.create`, `chart.trendline`, `sparkline.create` |
| Tables & Pivots | 25+ | `table.create`, `pivot.create`, `pivot.calculated` |
| Power Query & DAX | 10+ | `powerquery.create`, `powerquery.edit`, `slicer.set` |
| VBA | 5 | `vba.list`, `vba.view`, `vba.import`, `vba.update` |
| Task Pane | 30+ | `pane.click`, `pane.read`, `pane.grid.read`, `pane.wait` |
| ViewModel | 3 | `model`, `model.set`, `functions` |
| Ribbon & Dialogs | 20+ | `ribbon.click`, `dialog.dismiss`, `folder.dialog.set_path` |
| Desktop | 24 | `clipboard.read`, `process.list`, `window.focus`, `keys.send` |
| Testing | 16 | `assert.cell`, `test.report`, `screenshot.compare`, `ocr.screen` |

## Supported Agents

| Agent | Integration | Status |
|-------|-------------|--------|
| Claude Code | SKILL.md (auto-loaded) | Primary |
| OpenAI Codex | MCP server | Supported |
| Cursor | MCP server | Supported |
| Windsurf | MCP server | Supported |
| VS Code Copilot | MCP server | Supported |
| Visual Studio 2026 | Community extension + MCP | Supported |

## System Requirements

- Windows 10/11 (x64)
- Excel 2016+ or Microsoft 365
- No .NET SDK required (XRai ships self-contained)

## Documentation

- [Getting Started Guide](docs/getting-started.md)
- [Command Reference (283 commands)](docs/commands.md)
- [Task Pane & ViewModel Guide](docs/pane-controls.md)
- [Dialog & Folder Picker Guide](docs/dialogs.md)
- [Testing & Assertions Guide](docs/testing.md)
- [Visual Studio Integration](docs/vs-integration.md)
- [Troubleshooting](docs/troubleshooting.md)
- [Architecture](docs/architecture.md)
- [Contributing](CONTRIBUTING.md)

## Project Structure

```
src/
  XRai.Core/            Command router, JSON protocol, response helpers
  XRai.Com/             Excel COM interop (cells, charts, tables, pivots, Power Query, VBA)
  XRai.Hooks/           In-process WPF/WinForms hooks (Pilot.Expose, ViewModel binding)
  XRai.HooksClient/     Named pipe client for hooks communication
  XRai.UI/              FlaUI ribbon/dialog automation, Win32 dialog driver
  XRai.Vision/          Screenshots, OCR, visual diff
  XRai.Tool/            CLI entry point (self-contained exe)
  XRai.Mcp/             MCP server for Codex/Cursor/Windsurf
tests/
  XRai.Tests.Unit/      Unit tests
demo/
  XRai.Demo.PortfolioAddin/   Sample add-in (dark theme, 30+ controls)
  XRai.Demo.MacroGuard/       VBA management add-in (13 tabs, 50+ controls)
```

## License

MIT
