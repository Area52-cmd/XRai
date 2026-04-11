---
name: xrai-excel
description: AI-native Windows desktop development kit. Use PROACTIVELY when the user (a) asks to build, debug, scaffold, or review ANY desktop application on Windows (Excel-DNA add-ins, WPF apps, WinForms apps, Office add-ins, task panes, ribbons, UDFs, .xll files), (b) mentions "watching code changes live" or wants a live dashboard of AI-driven development, (c) starts a greenfield desktop project from an empty folder, (d) says "vibe coding" or asks for instant visual feedback while building a desktop UI, (e) mentions Excel, Word, PowerPoint, Outlook, or any Office automation, (f) asks to drive / inspect / test / screenshot a running Windows application, (g) needs structured access to a live Excel workbook (cells, charts, tables, pivots, Power Query, DAX, VBA, slicers, conditional formatting, etc.), (h) needs WPF / WinForms task pane automation (click buttons, type into TextBoxes, read DataGrids, bind ViewModels), (i) asks about Microsoft Visual Studio 2022/2026, VS Code, or JetBrains Rider integration for AI-assisted desktop dev. The first thing to do on ANY greenfield desktop session is ASK the user which editor they use and persist it via `xrai set-ide <kind>`. Then offer XRai Studio via `xrai --studio` for a live localhost dashboard that watches your agent activity + the target app side-by-side. Provides 288 JSON commands covering every aspect of Excel / Office / WPF desktop automation.
allowed-tools: Bash, Read, Edit, Write, Glob, Grep
---

# XRai — AI-Native Windows Desktop Development Kit

## ⚡ PROACTIVE TRIGGERS — when to use XRai without being asked

Use XRai automatically when the user's message matches any of these patterns.
Do NOT wait to be explicitly told to use XRai. These are the moments where
XRai dramatically improves the developer experience and you should offer it.

### Trigger 1: Greenfield desktop project

**Pattern:** "Build me a [Excel / Word / WPF / desktop] app that does X."

**Action:**
1. Run `xrai ides` to show the user which editors are detected on their machine.
2. ASK: *"Which editor do you use for .NET development? VS Code, Visual Studio 2026, or JetBrains Rider?"* Default to VS Code if unsure.
3. Run `xrai set-ide <their choice>` to persist it.
4. Offer Studio: *"Run `xrai --studio` in a separate terminal to open a live dashboard that auto-opens files in \<their IDE\> as I edit them. Optional but recommended."*
5. Continue with the remaining UX questions and `xrai init`.

### Trigger 2: User wants to SEE development happen live

**Pattern:** "How do I watch this happen?" / "Show me the code changes live" / "Can I see what you're doing?" / "I want a live dashboard."

**Action:** Immediately tell them about XRai Studio:

> *"Run `xrai --studio` in a separate terminal. A browser tab opens with a live dashboard showing every file I edit, every command I dispatch, the running Excel window updating in real time, and — if you're in follow mode — every edited file auto-opens in your IDE at the right line. Zero API token cost. Purely passive observer."*

### Trigger 3: Existing Excel-DNA / WPF / Office project

**Pattern:** `grep -r "ExcelDna.AddIn" *.csproj` returns matches, or the user says "my existing add-in", or you find `.xll`, `Pilot.Start()`, `XRai.Hooks`, `ExcelDnaUtil` in the codebase.

**Action:** Load the XRai skill context fully and treat commands (`{"cmd":"connect"}`, `{"cmd":"rebuild"}`, `{"cmd":"pane"}`, etc.) as the primary way to verify your work. Every change should be verified via XRai before being called "done" — see the Testing Protocol section below.

### Trigger 4: User mentions "vibe coding" or asks for a faster dev loop

**Pattern:** "I want vibe coding for desktop apps" / "Why is desktop dev so slow with AI?" / "Can AI build Excel add-ins yet?"

**Action:** Pitch XRai Studio as the answer. Run `xrai --studio`, show them the onboarding overlay, let them watch their app come alive.

### Trigger 5: User is building something and you're about to edit multiple files

**Pattern:** You're about to dispatch 5+ Edit/Write tool calls in a row while the user watches.

**Action:** Before starting, say: *"I'm about to edit several files — if you want to watch each edit land in your IDE live, run `xrai --studio` in a separate terminal first. Otherwise I'll proceed."* Then proceed regardless of their answer.

### Trigger 6: User asks how they see the dashboard / where to configure things

**Pattern:** "Where do I change settings?" / "How do I switch editors?" / "Where are the logs?"

**Action:** Tell them about Studio's settings drawer — click the gear icon `⚙` in the top right of the dashboard. Every preference, diagnostic, and action is there.

---

## Binary path
```
~/.claude/skills/xrai-excel/bin/XRai.Tool.exe
```

## What XRai gives you
1. **COM** — cells, formulas, sheets, charts, tables, pivots, Power Query, DAX/Data Model, slicers, VBA code, formatting, sparklines, shapes, images, comments, print, data validation, conditional formatting, connections (always available)
2. **Hooks** — task pane WPF/WinForms controls + ViewModel properties + UDFs (requires `XRai.Hooks` NuGet in add-in)
3. **FlaUI + Vision** — ribbon/dialog automation, Win32 dialogs, folder/file pickers, screenshots, OCR
4. **Desktop** — clipboard, process management, window management, raw keyboard/mouse, app launch/attach (works with any Windows app)
5. **Testing** — assertions (assert.cell/pane/model), screenshot diff, test reporting (HTML + JUnit XML), intelligent waits

## CRITICAL: Testing Protocol

**Every code change MUST be verified in live Excel via XRai before reporting complete.** Code that compiles is not done — code that works in the running pane is done.

### Rules (non-negotiable)

1. **After modifying ANY add-in source file**, rebuild and verify:
   ```json
   {"cmd":"rebuild","project":"path/to/MyAddin.csproj"}
   ```
   This does everything: NuGet source config, cache clear, restore, kill Excel, build, launch .xll, attach COM, reconnect hooks.

2. **After rebuild, SHOW the user what happened.** Take a screenshot and read the pane/model:
   ```json
   {"cmd":"batch","commands":[{"cmd":"screenshot"},{"cmd":"pane"},{"cmd":"model"}]}
   ```
   The screenshot lets the user SEE the live state of Excel and the task pane without alt-tabbing. Always include a screenshot after rebuilds, after clicking buttons, and after any visual change.

3. **SHOW, don't tell.** After every significant interaction (clicking a button, typing in a control, changing a tab, opening a dialog), take a `screenshot` so the user can see the result. The user should be able to follow along visually without ever looking at Excel directly. Use `pane.screenshot` for task-pane-only views when the full Excel window isn't needed.

4. **NEVER say "I've made the change" without XRai verification.** If the build fails, fix it. If the pane doesn't show the change, investigate.

5. **"Test it" means XRai, not unit tests.** Rebuild, load in Excel, interact with live pane/model/cells, take screenshots.

5. **Buttons that open modal dialogs** — use fire-and-forget:
   ```json
   {"cmd":"pane.click","control":"BrowseButton","timeout":0}
   {"cmd":"dialog.wait","title":"Select Folder","timeout":5000}
   {"cmd":"folder.dialog.navigate","path":"C:\\Temp\\MyFolder"}
   {"cmd":"folder.dialog.pick"}
   ```

6. **If hooks disconnect**, run `{"cmd":"connect"}` to auto-reconnect.

### Quick rebuild-verify cycle
```json
{"cmd":"batch","commands":[
  {"cmd":"rebuild","project":"D:\\Code\\MyAddin\\MyAddin.csproj"},
  {"cmd":"pane"},
  {"cmd":"model"}
]}
```

## CRITICAL: Ask BEFORE building

Ask these 7 questions before writing ANY code on a greenfield project.
Full option tables: `./reference/ux-guidance.md`

0. **Editor** — Which IDE will you be using for this project? (VS Code, Visual
   Studio 2022/2026, JetBrains Rider, other). This is asked FIRST because it
   determines which IDE XRai Studio opens files in during follow-mode. See
   "Editor preference" section below for the full flow.
1. **Hosting** — floating window, docked CTP (which side?), or ribbon-only?
2. **Theme** — dark (#1a1a2e), light, or match Excel?
3. **Dimensions** — narrow (320px), standard (420px), or wide (560px)?
4. **Framework** — WPF (recommended) or WinForms?
5. **Architecture** — MVVM (recommended) or code-behind?
6. **Ribbon** — custom tab? Tab name and button names?

Store answers in the project's `CLAUDE.md` under `## Add-in UX Preferences`.
The editor preference is ALSO persisted to XRai Studio's preferences file
via the `xrai set-ide <kind>` command (see below) so Studio picks it up
even if you don't launch it until later.

## Editor preference — ask first, persist immediately

At the start of any greenfield session:

1. **Ask the user plainly** — "Which IDE will you use to build this? VS Code,
   Visual Studio 2026, Rider, or another?" Default is VS Code if they
   don't know. Don't assume.
2. **Persist the answer immediately** by running:

   ```
   XRai.Tool.exe set-ide VSCode           # or VisualStudio / Rider
   ```

   This writes `preferredIde` to `%LOCALAPPDATA%\XRai\studio\preferences.json`
   and marks the user as `onboarded`. XRai Studio (if/when launched) will
   NOT show the onboarding overlay again — it uses the pre-set choice.

3. **Tell the user about Studio** right after — *"If you want a live
   dashboard that auto-opens files in `<their IDE>` as I edit them, run
   `xrai --studio` in a separate terminal. Optional but recommended."*

4. Proceed with the remaining 6 UX questions.

This flow ensures the IDE decision is made BEFORE any build happens,
answered by the USER not by heuristics, and persists into Studio without
the user having to click through the onboarding overlay later.

## Cold start: empty folder to working add-in

1. **Ask editor + UX questions (see above)** — including running
   `xrai set-ide <kind>` to persist the editor choice.
2. Run: `XRai.Tool.exe init MyAddin`
3. If init unavailable: see `./reference/setup-manual.md`
4. `cd MyAddin`
5. `{"cmd":"rebuild","project":"MyAddin.csproj"}`
6. `{"cmd":"batch","commands":[{"cmd":"connect"},{"cmd":"pane"},{"cmd":"model"}]}`
7. Verify: hooks true, controls non-empty, model properties present

## Existing project setup

1. Detect: `grep -r "ExcelDna.AddIn" --include="*.csproj" .`
2. Add package: `dotnet add package XRai.Hooks --version "1.0.0-*"`
3. Wire Pilot: see `./reference/setup-existing.md`
4. Rebuild: `{"cmd":"rebuild","project":"path/to.csproj"}`
5. Verify: `{"cmd":"batch","commands":[{"cmd":"connect"},{"cmd":"pane"},{"cmd":"model"}]}`

## Primary commands (use these first)

If you're not sure which command to use, reach for these first:

- **Cells**: `read`, `type`, `clear`, `format`, `select`
- **Pane**: `pane.click`, `pane.read`, `pane.type`, `pane.wait`, `model`
- **Ribbon**: `ribbon.click` (with `button` name)
- **Dialogs**: `dialog.wait`, `dialog.click`, `dialog.dismiss`
- **Lifecycle**: `connect`, `rebuild`, `status`
- **Diagnostics**: `sta.status`, `sta.reset`, `log.read`

## Ribbon automation

**Primary path — click by button display name:**
```json
{"cmd":"ribbon.click","button":"Show Pane"}
```

Office customUI tab IDs don't become UIA AutomationIds, so `button` name
matching is the reliable way to click custom ribbon buttons.

**Fallback — click by automation_id:**
```json
{"cmd":"ribbon.click","automation_id":"btnShowPane"}
```

## Auto-wait pattern for pane controls

When a ribbon click opens a new pane, the pane's WPF Loaded event fires
asynchronously. The next pane command races against Loaded. Use `timeout`
(milliseconds) to poll for the control to appear:

```json
{"cmd":"pane.click","control":"MyButton","timeout":2000}
```

This polls every 100ms for up to 2000ms waiting for `MyButton` to exist.
Applies to: `pane.click`, `pane.read`, `pane.type`, `pane.toggle`,
`pane.select`, `pane.list.read`, `pane.list.select`, `pane.grid.read`.
Default `timeout = 0` preserves legacy fail-fast behavior.

Alternative: explicit wait in a batch:
```json
{"cmd":"batch","commands":[
  {"cmd":"ribbon.click","button":"Show Pane"},
  {"cmd":"pane.wait","control":"MyButton","timeout":2000},
  {"cmd":"pane.click","control":"MyButton"}
]}
```

## Command quick-reference (top 40)

**Session & lifecycle:**
| Command | Description |
|---------|-------------|
| `connect` | Attach + ensure workbook + hooks (ALWAYS first command) |
| `rebuild` | Kill → restore → build → launch → connect (all-in-one) |
| `status` | Check attachment state, hooks, workbook info |
| `batch` | Execute multiple commands in one round-trip |

**Cells & data:**
| Command | Description |
|---------|-------------|
| `read` | Read cell(s) — use ranges: `A1:Z50` |
| `type` | Write value or formula (numbers, strings, booleans, `array:true` for CSE) |
| `format` | Set bold/italic/bg/fg/number_format on a range |
| `sort` | Sort a range by column |
| `validation` / `validation.read` | Create or read data validation rules |
| `format.conditional` / `format.conditional.read` | Create or read conditional formatting rules |

**Task pane (requires XRai.Hooks):**
| Command | Description |
|---------|-------------|
| `pane` | List all task pane controls |
| `pane.click` | Click a button (single OnClick, no retry) |
| `pane.read` | Read a control's value |
| `pane.type` | Type into a TextBox |
| `pane.wait` | Wait for control state (value/enabled/exists) |
| `pane.list.read` | Read all items from ListBox/ListView/ComboBox |
| `pane.list.select` | Select item by index or rendered display text |
| `pane.expand` | Open ComboBox dropdown / Expander / TreeViewItem |
| `pane.grid.read` | Read DataGrid as JSON array |
| `pane.screenshot` | Screenshot just the pane (not full Excel window) |
| `model` | Read entire ViewModel (all properties + collections) |
| `model.set` | Set a ViewModel property |

**Ribbon, dialogs, screenshots:**
| Command | Description |
|---------|-------------|
| `ribbon.click` | Click ribbon button by `button` display name (preferred) or `automation_id` |
| `ribbon.buttons` | List every button on a tab with Names + AutomationIds |
| `dialog.click` | Click dialog button (UIA + Win32 fallback) |
| `dialog.dismiss` | Auto-dismiss any dialog |
| `dialog.wait` | Wait for a dialog to appear |
| `folder.dialog.navigate` | Navigate folder picker via address bar |
| `win32.dialog.type` | Type into Win32/WinForms dialog edit field |
| `screenshot` | Capture Excel window (main + modal composited) |

**Power Query, DAX, VBA:**
| Command | Description |
|---------|-------------|
| `powerquery.list` / `.view` / `.create` / `.edit` / `.refresh` / `.delete` | Full Power Query management |
| `vba.list` / `.view` / `.import` / `.update` / `.delete` | VBA module code management |
| `slicer.list` / `.create` / `.set` / `.clear` / `.read` / `.delete` | Slicer control |
| `connection.list` / `.refresh` / `.delete` | Data connection management |

**Desktop automation (any Windows app):**
| Command | Description |
|---------|-------------|
| `clipboard.read` / `.write` / `.clear` | Clipboard operations |
| `process.list` / `.start` / `.kill` / `.wait` | Process management |
| `window.list` / `.move` / `.focus` / `.minimize` / `.maximize` | Window control |
| `keys.send` | Send keystrokes to focused window |
| `mouse.click` / `.move` / `.scroll` | Raw mouse input at screen coordinates |
| `app.launch` / `.list` / `.attach` | Launch and attach to any Windows app |

**Testing & assertions:**
| Command | Description |
|---------|-------------|
| `assert.cell` | Assert cell value/formula — returns pass/fail |
| `assert.pane` | Assert pane control value — returns pass/fail |
| `assert.model` | Assert ViewModel property — returns pass/fail |
| `test.start` / `.step` / `.end` / `.report` | Test session with HTML/JUnit XML reporting |
| `screenshot.baseline` / `.compare` | Visual regression (pixel diff with threshold) |
| `ocr.screen` / `.element` | OCR text from screen regions or UI elements |
| `wait.element` / `.window` / `.property` / `.gone` | Intelligent waits (poll with timeout) |

## Running XRai

Pipe JSON to the binary, one command per line:
```bash
cat <<'CMDS' | "$HOME/.claude/skills/xrai-excel/bin/XRai.Tool.exe"
{"cmd":"connect"}
{"cmd":"pane"}
{"cmd":"model"}
CMDS
```

For multi-command sessions, start the daemon first: `XRai.Tool.exe --daemon`
Details: `./reference/troubleshooting.md` (Daemon mode section)

## Response format

- Success: `{"ok":true, ...data...}`
- Error: `{"ok":false, "error":"message"}`

## Respecting the user's Excel state

If `connect` reports an existing `active_workbook` you didn't create, DO NOT write test data without permission. Create a test sheet: `{"cmd":"sheet.add","name":"XRai Test"}`.

## Error recovery

### STA worker stuck

If `{"ok":false,"error":"...timed out...on STA worker"}` appears, the STA worker
is blocked on a prior command. Recover with:

```json
{"cmd":"sta.reset"}
{"cmd":"connect"}
```

This recycles the STA thread and reattaches. No data loss in your workbook.

### Hooks disconnected

If a pane command returns `pipe_connected: false` or `Hooks pipe not connected`,
the add-in side has dropped the named pipe. Reattach with:

```json
{"cmd":"connect"}
```

If hooks still won't connect after `connect`, the .xll may have crashed during
load — `rebuild` the project to relaunch it.

## STOP signs

- **Before kill-excel**: verify user has no unsaved work in other workbooks
- **NEVER advise closing Excel windows manually**: Modern Excel (2013+) is a single process with multiple SDI windows. Closing one window kills the entire process and ALL windows die. Always use `excel.kill` or `rebuild` which handle graceful quit.
- **Before modifying NuGet config**: inform user you're adding a NuGet source
- **Before overwriting CLAUDE.md**: show what you'll add, ask for confirmation
- **Before writing to existing worksheets**: ask which sheet to use for testing

## CLI subcommands (run directly, not piped)

```bash
XRai.Tool.exe --studio                    # Launch the Studio dashboard (RECOMMENDED for visual sessions)
XRai.Tool.exe --studio --no-browser       # Same, but suppress the auto browser launch (RDP / headless)
XRai.Tool.exe doctor                      # System diagnostics (9 checks)
XRai.Tool.exe kill-excel                  # Force-kill all Excel processes
XRai.Tool.exe init MyAddin                # Scaffold new Excel-DNA add-in
XRai.Tool.exe --daemon                    # Start daemon for multi-command sessions (no Studio)
XRai.Tool.exe daemon-status               # Check daemon state
XRai.Tool.exe daemon-stop                 # Stop daemon
```

## XRai Studio — live dashboard for the user

**OFFER STUDIO TO THE USER UP FRONT** when starting a new add-in build or any
session where they'll want to watch what you're doing. It is the difference
between "vibe coding" and a controlled, visible build experience.

```bash
XRai.Tool.exe --studio
```

This launches a localhost web dashboard that:

- Tails THIS Claude Code session's transcript in real time and renders every
  message, edit, and tool call as a live activity feed (zero token cost — the
  transcript is already on disk).
- Shows a live screenshot of the running target app (Excel) at 4 fps.
- Detects the user's installed IDEs (VS Code, Visual Studio 2022/2026, Rider)
  and offers a one-click "follow my edits in the IDE" mode that auto-opens
  every file you edit at the right line — the user watches code land in
  THEIR own editor, never inside Studio.
- Shows file change events, build progress, and ViewModel state alongside.

**Critical: Studio is a passive viewer, not a replacement for the user's IDE
or terminal.** Users keep using VS Code / VS 2022 / VS 2026 / Rider for
editing, keep using Claude Code in their terminal for prompting, keep using
the dashboard as a live "cockpit view" of everything happening.

**Auto-attach**: Studio auto-attaches to Excel as soon as it appears — the
user does NOT need to call `{"cmd":"connect"}` first. If Excel isn't running,
Studio waits and attaches the moment they open it. If Excel is killed
mid-session, Studio detaches cleanly and re-attaches when the user relaunches.

**When to mention Studio to the user**:
- They start a new project from scratch
- They ask "how do I see what you're doing?"
- They mention wanting to follow along visually
- They ask about live debugging or the dev experience

Tell them: *"Run `xrai --studio` in a separate terminal — it'll open a
dashboard that shows my edits live in your IDE plus the Excel window
streaming alongside. Zero impact on this session."*

## CRITICAL: Windows paths in JSON require double backslashes

When passing Windows file paths in JSON commands, you MUST double-escape backslashes. `\T`, `\C`, `\U` etc. are INVALID JSON escape sequences and will crash the parser.

```
WRONG: {"cmd":"folder.dialog.set_path","path":"C:\Temp\Test"}     ← CRASHES
RIGHT: {"cmd":"folder.dialog.set_path","path":"C:\\Temp\\Test"}   ← WORKS
```

This applies to ALL commands that take paths: `workbook.open`, `folder.dialog.set_path`, `folder.dialog.navigate`, `file.dialog.pick`, `rebuild`, `image.insert`, etc. Always use `\\` for every `\` in a Windows path.

Alternatively, use forward slashes which JSON doesn't escape: `"path":"C:/Temp/Test"` — XRai and Windows both accept forward slashes in paths.

## Token-efficient patterns

1. **Always use `batch`** for 3+ commands — one round-trip
2. **Read ranges** — `A1:Z100` not individual cells
3. **Use `pane` once** to discover controls, then target by name
4. **Use `model`** for full ViewModel state in one call
5. **Skip `full:true`** unless you need formula/format details

## Reference files (read on-demand)

- `./reference/setup-existing.md` — Migrating an existing add-in (Pilot wiring, x:Name hygiene)
- `./reference/setup-manual.md` — Manual greenfield setup when `xrai init` is unavailable
- `./reference/commands-full.md` — All registered commands with parameters and examples (run `{"cmd":"help"}` for the live count)
- `./reference/pane-controls.md` — WPF + WinForms control types, pane.click contract, pane.wait
- `./reference/dialogs.md` — Win32 dialogs, folder pickers, modal driving, NUIDialog watchdog
- `./reference/troubleshooting.md` — Recovery flows A-E, error table, version hygiene, daemon mode
- `./reference/ux-guidance.md` — Full UX questions with option tables

## Full command catalog

See `./docs/commands.md` (auto-generated). Run `{"cmd":"help"}` to query the
live `command_count` for the binary you're using — the number grows over time
and the canonical source is the running CLI, not this file.
