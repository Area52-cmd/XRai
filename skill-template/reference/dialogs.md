# Dialog Driving Reference

## dialog.click and dialog.dismiss — two-tier (UIA + Win32)

Both commands try UIA first (FlaUI ModalWindows), then fall through to Win32 `EnumWindows`. Response includes `"source":"uia"` or `"source":"win32"`.

```json
{"cmd":"dialog.click","button":"OK"}
{"cmd":"dialog.dismiss"}
```

`dialog.dismiss` auto-picks the safest button using a priority list.

## dialog.wait — wait for a dialog to appear

```json
{"cmd":"dialog.wait","title":"Select Folder","timeout":5000}
```

## dialog.read — inspect open dialog

```json
{"cmd":"dialog.read"}
```

Returns title, buttons, edit controls, and structure.

## Win32 dialog commands

For native Win32 dialogs (#32770), WinForms dialogs, and NUIDialog windows.

### win32.dialog.list — enumerate all dialogs

```json
{"cmd":"win32.dialog.list"}
```

Returns all dialog-class windows with title, class_name, hwnd, is_likely_dialog.

### win32.dialog.read — full dialog structure

```json
{"cmd":"win32.dialog.read","title":"Save Snapshot"}
```

Returns title, class_name, edit_controls (with control_id, class_name, current_text), and buttons array.

### win32.dialog.type — type into edit control

```json
{"cmd":"win32.dialog.type","title":"Save Snapshot","text":"my-snapshot","submit":true}
```

Uses `WM_SETTEXT` directly — no keystroke simulation. Works on:
- Native Win32 Edit controls
- WinForms TextBox (`WindowsForms10.EDIT.*`)
- Rich Edit (`RichEdit20W`, `RICHEDIT50W`)
- ComboBoxEx32 editable fields

Target by `control_id` or `index` for multiple edit controls:
```json
{"cmd":"win32.dialog.type","title":"Save Snapshot","text":"value","control_id":1001,"submit":true}
{"cmd":"win32.dialog.type","title":"Save Snapshot","text":"value","index":0}
```

### win32.dialog.click — click dialog button

```json
{"cmd":"win32.dialog.click","button":"OK","title":"Save Snapshot"}
```

## Folder picker dialogs

```json
{"cmd":"folder.dialog.navigate","path":"C:\\Temp\\MyFolder"}
{"cmd":"folder.dialog.pick"}
```

## File picker dialogs

```json
{"cmd":"file.dialog.pick","path":"C:\\Data\\input.xlsx"}
```

## NUIDialog watchdog

`workbook.open` runs a background watchdog that auto-dismisses NUIDialog windows (update-links, protected view, format mismatch, etc.). Safe-button priority:

`Don't Update` > `No` > `Close` > `Cancel` > `OK` > `Yes` > `Continue` > `Enable Editing` > `Enable Content`

## excel.autodismiss — persistent OLE dismissal

For the "waiting for another application to complete an OLE action" message box:

```json
{"cmd":"excel.autodismiss","enabled":true}
```

Starts a background loop that auto-dismisses `#32770` OLE dialogs.

## Modal-opening button pattern

When a pane button opens a modal, use `timeout:0` so the pipe does not block:

```json
{"cmd":"pane.click","control":"BrowseButton","timeout":0}
{"cmd":"dialog.wait","title":"Select Folder","timeout":5000}
{"cmd":"folder.dialog.navigate","path":"C:\\Temp"}
{"cmd":"folder.dialog.pick"}
```
