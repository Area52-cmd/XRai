# Pane Controls Reference

## pane.click contract

`pane.click` on a WPF `ButtonBase` invokes `ButtonBase.OnClick` exactly once via reflection. No retry, no pre-click focus, no post-click observation. Virtual dispatch honors subclass overrides (ToggleButton, CheckBox, RadioButton).

Response:
```json
{"ok":true,"clicked":true,"method":"ButtonBase.OnClick",
 "has_command":true,"command_can_execute":true,"command_executed":true}
```

- `command_can_execute` — snapshot of `CanExecute` at click time
- `command_executed` — mirrors `command_can_execute` (only synchronous signal available)
- Verification is the caller's responsibility: read `model`, `pane.read`, or cells after clicking

### Click priority order

1. **`pane.click`** — primary, always preferred
2. **`pane.focus` + `pane.key keys:"Enter"`** — fallback if click silently no-ops
3. **`model.set`** — only if add-in exposes Commands as properties

## pane.wait — wait for control state

```json
{"cmd":"pane.wait","control":"StatusLabel","property":"value","value":"Done","timeout":5000}
{"cmd":"pane.wait","control":"GoButton","property":"enabled","value":true,"timeout":3000}
{"cmd":"pane.wait","control":"NewPanel","property":"exists","value":true,"timeout":5000}
```

## pane.screenshot — capture just the WPF pane

```json
{"cmd":"pane.screenshot"}
```

## pane.context_menu — right-click context menu

```json
{"cmd":"pane.context_menu","control":"DataGrid","item":"Delete Row"}
```

## pane.drag — drag between controls

```json
{"cmd":"pane.drag","from":"SourceItem","to":"TargetArea"}
```

## WPF control types supported

| Type | Commands |
|---|---|
| Button, ToggleButton, RepeatButton | `pane.click` |
| CheckBox, RadioButton | `pane.click` (toggles), `pane.toggle` |
| TextBox, PasswordBox, RichTextBox | `pane.type`, `pane.read` |
| ComboBox | `pane.select`, `pane.read` |
| ListBox, ListView | `pane.select`, `pane.read` |
| DataGrid | `pane.grid.read`, `pane.grid.cell`, `pane.grid.click` |
| TabControl | `pane.tab` |
| Slider | `pane.set`, `pane.read` |
| TextBlock, Label | `pane.read` |
| Expander | `pane.click` (toggles expand) |
| TreeView | `pane.tree.read`, `pane.tree.expand` |

## WinForms control types supported

| Type | Commands |
|---|---|
| Button | `pane.click` |
| TextBox | `pane.type`, `pane.read` |
| CheckBox | `pane.toggle` |
| ComboBox | `pane.select` |
| DataGridView | `pane.grid.read` |
| Label | `pane.read` |
| TabControl | `pane.tab` |

## Synthetic naming for unnamed controls

Controls without `x:Name` get synthetic names:
- `_unnamed_Button_0` (no content text)
- `_unnamed_Button_Save_0` (has `Content="Save"`)

Best practice: always add `x:Name` to interactive controls in XAML.

## Deep nesting discovery

`pane` walks the full visual tree recursively. Controls nested inside UserControls, ContentPresenters, ItemsControls, and DataTemplates are all discovered. Use `pane` output to find the exact name before targeting.

## Rendered text matching

`pane.read` returns the rendered text of a control — what the user sees on screen. For TextBlocks with bindings, this is the resolved value, not the binding expression.

## Modal-opening buttons: use timeout:0

Buttons that open modal dialogs block the UI thread. Use fire-and-forget:

```json
{"cmd":"pane.click","control":"BrowseButton","timeout":0}
{"cmd":"dialog.wait","title":"Select Folder","timeout":5000}
{"cmd":"folder.dialog.navigate","path":"C:\\Temp"}
{"cmd":"folder.dialog.pick"}
```

## pane.key parameter naming

Both `keys` (canonical) and `key` (alias) work:
```json
{"cmd":"pane.key","control":"InputBox","keys":"Enter"}
{"cmd":"pane.key","control":"InputBox","key":"Enter"}
```
