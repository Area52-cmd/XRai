# Dialog & Folder Picker Guide

XRai provides a two-tier dialog automation system: UIA (FlaUI) for standard modal dialogs and Win32 for native dialogs. Both tiers are tried automatically -- commands attempt UIA first, then fall back to Win32.

## UIA Dialogs

### dialog.read -- Inspect an Open Dialog

```json
{"cmd":"dialog.read"}
```

Returns the dialog title, button labels, edit controls, and structure.

### dialog.click -- Click a Dialog Button

```json
{"cmd":"dialog.click","button":"OK"}
{"cmd":"dialog.click","button":"Cancel"}
{"cmd":"dialog.click","button":"Yes"}
```

### dialog.dismiss -- Auto-Dismiss with Safest Button

```json
{"cmd":"dialog.dismiss"}
```

Automatically picks the safest button using a priority list:

`Don't Update` > `No` > `Close` > `Cancel` > `OK` > `Yes` > `Continue` > `Enable Editing` > `Enable Content`

### dialog.wait -- Wait for a Dialog to Appear

```json
{"cmd":"dialog.wait","title":"Select Folder","timeout":5000}
```

Blocks until a dialog with the specified title appears, or times out.

### dialog.auto_click -- Persistent Auto-Click

```json
{"cmd":"dialog.auto_click","button":"OK","title":"Update Links"}
```

Sets up a background watcher that automatically clicks the specified button whenever a matching dialog appears.

## Win32 Dialogs

For native Win32 dialogs (`#32770` class), WinForms dialogs, and NUIDialog windows that UIA cannot reach.

### win32.dialog.list -- Enumerate All Dialogs

```json
{"cmd":"win32.dialog.list"}
```

Returns all dialog-class windows with title, class_name, hwnd, and is_likely_dialog flag.

### win32.dialog.read -- Full Dialog Structure

```json
{"cmd":"win32.dialog.read","title":"Save Snapshot"}
```

Returns title, class_name, edit_controls (with control_id, class_name, current_text), and buttons array.

### win32.dialog.type -- Type into a Dialog Edit Control

```json
{"cmd":"win32.dialog.type","title":"Save Snapshot","text":"my-snapshot","submit":true}
```

Uses `WM_SETTEXT` directly -- no keystroke simulation. Works on:

- Native Win32 Edit controls
- WinForms TextBox (`WindowsForms10.EDIT.*`)
- Rich Edit controls (`RichEdit20W`, `RICHEDIT50W`)
- ComboBoxEx32 editable fields

Target a specific control when a dialog has multiple edit fields:

```json
{"cmd":"win32.dialog.type","title":"Save Snapshot","text":"value","control_id":1001,"submit":true}
{"cmd":"win32.dialog.type","title":"Save Snapshot","text":"value","index":0}
```

### win32.dialog.click -- Click a Dialog Button

```json
{"cmd":"win32.dialog.click","button":"OK","title":"Save Snapshot"}
```

## Folder Picker Dialogs

Navigate a folder picker to a specific path and confirm the selection:

```json
{"cmd":"folder.dialog.navigate","path":"C:\\Temp\\MyFolder"}
{"cmd":"folder.dialog.pick"}
```

For combined set-and-pick:

```json
{"cmd":"folder.dialog.set_path","path":"C:\\Temp\\MyFolder"}
```

## File Picker Dialogs

```json
{"cmd":"file.dialog.pick","path":"C:\\Data\\input.xlsx"}
```

## NUIDialog Watchdog

The `workbook.open` command runs a background watchdog that auto-dismisses NUIDialog windows. These are the modal dialogs Excel shows when opening files with external links, protected views, format mismatches, and similar conditions.

The watchdog uses a safe-button priority list:

`Don't Update` > `No` > `Close` > `Cancel` > `OK` > `Yes` > `Continue` > `Enable Editing` > `Enable Content`

This means `workbook.open` handles most dialog interruptions automatically. No manual intervention is needed in the common case.

## excel.autodismiss -- OLE Dialog Suppression

For the persistent "waiting for another application to complete an OLE action" message box:

```json
{"cmd":"excel.autodismiss","enabled":true}
```

Starts a background loop that continuously monitors for and auto-dismisses `#32770` class OLE dialogs. Disable with:

```json
{"cmd":"excel.autodismiss","enabled":false}
```

## Modal-Opening Button Pattern

When a task pane button opens a modal dialog, the WPF UI thread blocks. Use `timeout:0` on the `pane.click` to fire-and-forget, then handle the dialog separately:

```json
{"cmd":"pane.click","control":"BrowseButton","timeout":0}
{"cmd":"dialog.wait","title":"Select Folder","timeout":5000}
{"cmd":"folder.dialog.navigate","path":"C:\\Temp"}
{"cmd":"folder.dialog.pick"}
```

This pattern applies to any pane button that calls `MessageBox.Show()`, `FolderBrowserDialog.ShowDialog()`, `OpenFileDialog.ShowDialog()`, or similar blocking calls.

## Response Source Indicator

Dialog commands include a `source` field in their response indicating which tier handled the request:

```json
{"ok": true, "source": "uia", "button": "OK", "clicked": true}
{"ok": true, "source": "win32", "button": "OK", "clicked": true}
```

This is useful for debugging when a dialog command succeeds but produces unexpected results.
