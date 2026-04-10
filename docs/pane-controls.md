# Task Pane & ViewModel Guide

XRai provides in-process access to WPF and WinForms task pane controls through the `XRai.Hooks` library. This guide covers how control discovery works, which controls are supported, and how to interact with them.

## How Pilot.Expose Works

When you call `Pilot.Expose(element)` in your add-in's `AutoOpen()`, XRai.Hooks walks the WPF visual tree (or WinForms control tree) recursively and registers every control that has a name. For WPF, this means controls with `x:Name="..."` in XAML. For WinForms, controls with a non-empty `Name` property.

```csharp
Pilot.Start();                    // Start the named pipe server
Pilot.Expose(myTaskPane);         // Walk tree, register controls
Pilot.ExposeModel(myViewModel);   // Expose INotifyPropertyChanged properties
```

Controls nested inside UserControls, ContentPresenters, ItemsControls, and DataTemplates are all discovered. The `pane` command returns the full list of discovered controls.

## WPF Control Types Supported

| Control | Read | Write | Click | Commands |
|---------|------|-------|-------|----------|
| `TextBox` | yes | yes | -- | `pane.type`, `pane.read` |
| `PasswordBox` | -- | yes | -- | `pane.type` (write-only) |
| `RichTextBox` | yes | yes | -- | `pane.type`, `pane.read` (plain text) |
| `Button` | -- | -- | yes | `pane.click` |
| `ToggleButton` | yes | -- | yes | `pane.click` (toggles) |
| `RepeatButton` | -- | -- | yes | `pane.click` |
| `CheckBox` | yes | yes | yes | `pane.toggle`, `pane.read` |
| `RadioButton` | yes | yes | -- | `pane.click`, `pane.read` |
| `Label` | yes | -- | -- | `pane.read` |
| `TextBlock` | yes | -- | -- | `pane.read` |
| `ComboBox` | yes | yes | -- | `pane.select`, `pane.read` |
| `ListBox` | yes | yes | -- | `pane.select`, `pane.read` |
| `ListView` | yes | yes | -- | `pane.select`, `pane.read` |
| `DataGrid` | yes | -- | yes | `pane.grid.read`, `pane.grid.cell`, `pane.grid.select` |
| `TabControl` | yes | yes | -- | `pane.tab` (by index or header name) |
| `TreeView` | yes | -- | -- | `pane.tree.expand` (by path) |
| `DatePicker` | yes | yes | -- | `pane.type` (yyyy-MM-dd format) |
| `Slider` | yes | yes | -- | `pane.read`, `pane.set` |
| `ProgressBar` | yes | -- | -- | `pane.read` (value, min, max, percentage) |
| `Expander` | yes | yes | -- | `pane.toggle` (expand/collapse) |
| `ScrollViewer` | -- | -- | -- | `pane.scroll` |

## WinForms Control Types Supported

| Control | Read | Write | Click | Commands |
|---------|------|-------|-------|----------|
| `Button` | -- | -- | yes | `pane.click` |
| `TextBox` | yes | yes | -- | `pane.type`, `pane.read` |
| `CheckBox` | yes | yes | -- | `pane.toggle` |
| `ComboBox` | yes | yes | -- | `pane.select` |
| `DataGridView` | yes | -- | -- | `pane.grid.read` |
| `Label` | yes | -- | -- | `pane.read` |
| `TabControl` | yes | yes | -- | `pane.tab` |

## The pane.click Contract

`pane.click` on a WPF `ButtonBase` invokes `ButtonBase.OnClick` exactly once via reflection. There is no retry, no pre-click focus change, and no post-click observation. Virtual dispatch honors subclass overrides (ToggleButton, CheckBox, RadioButton).

Response:

```json
{
  "ok": true,
  "clicked": true,
  "method": "ButtonBase.OnClick",
  "has_command": true,
  "command_can_execute": true,
  "command_executed": true
}
```

- `command_can_execute` -- snapshot of `ICommand.CanExecute` at click time
- `command_executed` -- mirrors `command_can_execute` (only the synchronous signal is available)
- Verification is the caller's responsibility: read `model`, `pane.read`, or cells after clicking

### Click priority order

1. **`pane.click`** -- primary, always preferred
2. **`pane.focus` + `pane.key keys:"Enter"`** -- fallback if click silently no-ops
3. **`model.set`** -- only if the add-in exposes Commands as properties

## pane.wait -- Wait for Control State

Wait for a control to reach a specific state before proceeding:

```json
{"cmd":"pane.wait","control":"StatusLabel","property":"value","value":"Done","timeout":5000}
{"cmd":"pane.wait","control":"GoButton","property":"enabled","value":true,"timeout":3000}
{"cmd":"pane.wait","control":"NewPanel","property":"exists","value":true,"timeout":5000}
```

## pane.screenshot -- Capture the Task Pane

```json
{"cmd":"pane.screenshot"}
```

Captures just the WPF pane as a PNG, excluding the rest of the Excel window.

## pane.context_menu -- Right-Click Context Menu

```json
{"cmd":"pane.context_menu","control":"DataGrid","item":"Delete Row"}
```

Right-clicks the control and selects the named item from the resulting context menu.

## pane.drag -- Drag Between Controls

```json
{"cmd":"pane.drag","from":"SourceItem","to":"TargetArea"}
```

## DataGrid Operations

```json
{"cmd":"pane.grid.read","control":"HoldingsGrid"}
```

Returns all data as a 2D array with column headers:

```json
{
  "ok": true,
  "name": "HoldingsGrid",
  "info": {"rows": 8, "columns": 6, "column_headers": ["Symbol","Qty","Price","Value","P&L","%"]},
  "data": [["AAPL","50","178.25","8912.5","1787.5","25.1"], ...]
}
```

Read a specific cell or select a row:

```json
{"cmd":"pane.grid.cell","control":"HoldingsGrid","row":0,"col":2}
{"cmd":"pane.grid.select","control":"HoldingsGrid","row":2}
```

## TabControl Switching

```json
{"cmd":"pane.tab","control":"MainTabs","tab":"Holdings"}
{"cmd":"pane.tab","control":"MainTabs","tab":"1"}
```

Switch by header name or zero-based index.

## TreeView Navigation

```json
{"cmd":"pane.tree.expand","control":"SectorTree","path":"Technology/Semiconductors"}
```

Expand nodes by specifying the full path with `/` separators.

## Keyboard and Focus

```json
{"cmd":"pane.focus","control":"SpotInput"}
{"cmd":"pane.key","control":"SpotInput","keys":"Enter"}
{"cmd":"pane.key","control":"SpotInput","keys":"Control+A"}
```

Both `keys` (canonical) and `key` (alias) parameter names are accepted.

## Mouse Simulation

```json
{"cmd":"pane.double_click","control":"HoldingsGrid"}
{"cmd":"pane.right_click","control":"CellValue"}
{"cmd":"pane.hover","control":"HelpIcon"}
```

## Modal-Opening Buttons

Buttons that open modal dialogs block the UI thread. Use `timeout:0` for fire-and-forget:

```json
{"cmd":"pane.click","control":"BrowseButton","timeout":0}
{"cmd":"dialog.wait","title":"Select Folder","timeout":5000}
{"cmd":"folder.dialog.navigate","path":"C:\\Temp"}
{"cmd":"folder.dialog.pick"}
```

## Synthetic Naming for Unnamed Controls

Controls without `x:Name` receive synthetic names:

- `_unnamed_Button_0` (no content text)
- `_unnamed_Button_Save_0` (has `Content="Save"`)

Best practice: always add `x:Name` to interactive controls in XAML.

## ViewModel Binding

`Pilot.ExposeModel(vm)` exposes all public properties of an `INotifyPropertyChanged` object:

```json
{"cmd":"model"}
```

Returns every property including collections:

```json
{
  "ok": true,
  "properties": {
    "Spot": 100.0,
    "Vol": 0.25,
    "Result": 42.5,
    "Status": "Ready",
    "Trades": [{"Symbol":"AAPL","Qty":50}, ...]
  }
}
```

Set a property:

```json
{"cmd":"model.set","property":"Spot","value":105.0}
```

List registered UDFs:

```json
{"cmd":"functions"}
```

## Rendered Text Matching

`pane.read` returns the rendered text of a control -- what the user sees on screen. For TextBlocks with bindings, this is the resolved value, not the binding expression.
