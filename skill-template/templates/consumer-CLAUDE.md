# This Project Uses XRai

This is an Excel-DNA add-in with XRai hooks enabled. When the user asks to test, debug, inspect, build, or interact with anything in Excel — cells, task pane controls, ViewModel, UDFs, ribbon — use XRai.

## XRai is available globally

The full XRai skill is at `~/.claude/skills/xrai-excel/` (already loaded in every Claude Code session). It contains:
- The `XRai.Tool.exe` binary at `~/.claude/skills/xrai-excel/bin/XRai.Tool.exe`
- The `XRai.Hooks` NuGet package (already installed in this project)
- The full command reference at `~/.claude/skills/xrai-excel/docs/commands.md`

## How to invoke XRai

```bash
echo '{"cmd":"connect"}' | "$HOME/.claude/skills/xrai-excel/bin/XRai.Tool.exe"
```

## This project's add-in wiring

- `Pilot.Start()` is called in `AutoOpen()`
- `Pilot.Expose(taskPane)` exposes WPF/WinForms controls
- `Pilot.ExposeModel(viewModel)` exposes ViewModel properties

So you have access to:
- **COM**: cells, sheets, formulas, charts, tables, formatting
- **Hooks**: task pane controls (TextBox/Button/DataGrid/TabControl), ViewModel properties, UDFs
- **UI**: ribbon, dialogs, screenshots

## Standard workflow

1. Start by orienting: `{"cmd":"batch","commands":[{"cmd":"connect"},{"cmd":"pane"},{"cmd":"model"}]}`
2. Interact with the pane: `{"cmd":"pane.click","control":"..."}`
3. Read results: `{"cmd":"pane.read","control":"..."}` or `{"cmd":"model"}`
4. Verify in Excel: `{"cmd":"read","ref":"A1:D10"}`
5. Capture visual if needed: `{"cmd":"screenshot"}`

Always use `batch` for multiple commands. Read ranges not cells. Use `model` for full state.

## Rebuilding the add-in

Use the `rebuild` command — it handles everything (NuGet cache clear, restore, kill Excel, build, launch, reconnect):
```json
{"cmd":"rebuild","project":"<path-to-this-project.csproj>"}
```

**ALWAYS rebuild and verify in live Excel after every code change. Never report a change as complete without XRai verification.**

See the global skill at `~/.claude/skills/xrai-excel/SKILL.md` for complete setup, troubleshooting, and the full command catalog.
