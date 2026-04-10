# XRai Command Reference

Complete catalog of all 283 JSON commands XRai accepts on stdin. Responses come on stdout as `{"ok":true,...}` or `{"ok":false,"error":"..."}`.

## Connection (5)

| Command | Args | Description |
|---------|------|-------------|
| `attach` | `[pid]` | Connect to running Excel (optional specific PID) |
| `wait` | -- | Block until Excel appears, then attach |
| `detach` | -- | Disconnect from Excel |
| `status` | -- | Attached state, version, hooks status, pipe name |
| `doctor` | -- | Diagnose system requirements (CLI subcommand) |

## Cells (13)

| Command | Args | Description |
|---------|------|-------------|
| `read` | `ref [full]` | Read cell or range. `full:true` adds formula/format |
| `type` | `ref value` | Write value (string/number) or formula (starts with `=`) |
| `type` | `ref values[]` | Write array of values to range |
| `clear` | `ref` | Clear cell contents |
| `select` | `ref` | Select range in UI |
| `format` | `ref props` | Set bold, font_name, font_size, bg, number_format |
| `format.read` | `ref` | Read all format properties |
| `format.border` | `ref side weight color style` | Set borders |
| `format.align` | `ref horizontal vertical wrap_text text_rotation indent` | Alignment |
| `format.font` | `ref italic underline strikethrough color` | Font details |
| `format.style` | `ref style` | Apply named style |
| `format.conditional` | `ref type operator value format` | Conditional format rule |
| `format.conditional.clear` | `ref` | Clear conditional rules |

## Sheets (5)

| Command | Args | Description |
|---------|------|-------------|
| `sheets` | -- | List all sheets with index, visibility, active |
| `sheet.add` | `name [after]` | Add new sheet |
| `sheet.rename` | `from to` | Rename sheet |
| `sheet.delete` | `name` | Delete (suppresses confirm dialog) |
| `goto` | `target` | Activate sheet or jump to named range |

## Named Ranges (4)

| Command | Args | Description |
|---------|------|-------------|
| `names` | -- | List all named ranges |
| `name.set` | `name ref` | Create named range |
| `name.read` | `name` | Read named range details |
| `name.delete` | `name` | Delete named range |

## Calculation (6)

| Command | Args | Description |
|---------|------|-------------|
| `calc` | -- | Full recalculation |
| `calc.mode` | `[mode]` | Get/set auto/manual/semiautomatic |
| `wait.calc` | `[timeout]` | Wait for calc to finish (ms timeout) |
| `wait.cell` | `ref condition [value] [timeout]` | Wait for cell condition (not_empty/equals/not_error) |
| `time.calc` | -- | Benchmark recalc duration |
| `time.udf` | `function args N` | Benchmark UDF over N iterations |

## Charts (10)

| Command | Args | Description |
|---------|------|-------------|
| `chart.list` | -- | List all charts on active sheet |
| `chart.create` | `type data title` | Create chart (column/bar/line/pie/scatter/area) |
| `chart.delete` | `name` | Delete chart |
| `chart.type` | `name [type]` | Get/set chart type |
| `chart.title` | `name [title]` | Get/set chart title |
| `chart.data` | `name [data]` | Get/set data source |
| `chart.legend` | `name show position` | Legend configuration |
| `chart.axis` | `name which title min max format` | Axis configuration |
| `chart.series` | `name action` | List/add/remove/format series |
| `chart.export` | `name path` | Export as PNG |

## Sparklines (3)

| Command | Args | Description |
|---------|------|-------------|
| `sparkline.create` | `ref data type` | Create sparkline (line/column/winloss) |
| `sparkline.format` | `ref color weight` | Format sparkline |
| `sparkline.delete` | `ref` | Delete sparkline |

## Tables (12)

| Command | Args | Description |
|---------|------|-------------|
| `table.list` | -- | List all tables |
| `table.create` | `ref name [style]` | Create structured table |
| `table.delete` | `name` | Unlist (convert to range) |
| `table.style` | `name style` | Apply table style |
| `table.resize` | `name ref` | Resize to new range |
| `table.totals` | `name show [column] [function]` | Show/hide totals row |
| `table.filter` | `name column value` | Filter by column |
| `table.filter.clear` | `name` | Clear filters |
| `table.sort` | `name column [order]` | Sort table |
| `table.row.add` | `name values` | Append row |
| `table.column.add` | `name header [formula]` | Add column |
| `table.data` | `name` | Read all data as JSON array |

## Filters (5)

| Command | Args | Description |
|---------|------|-------------|
| `filter.on` | `ref` | Enable AutoFilter |
| `filter.off` | -- | Disable AutoFilter |
| `filter.set` | `column criteria [operator]` | Apply filter criteria |
| `filter.clear` | `[column]` | Clear filter |
| `filter.read` | -- | Read current criteria |

## Pivot Tables (7)

| Command | Args | Description |
|---------|------|-------------|
| `pivot.list` | -- | List all pivots |
| `pivot.create` | `source destination name` | Create from data |
| `pivot.refresh` | `name` | Refresh data |
| `pivot.field.add` | `name field area [function]` | Add row/column/value/filter field |
| `pivot.field.remove` | `name field` | Remove field |
| `pivot.style` | `name style` | Apply style |
| `pivot.data` | `name` | Read values |

## Pivot Calculated Fields (3)

| Command | Args | Description |
|---------|------|-------------|
| `pivot.calculated` | `name field_name formula` | Add calculated field |
| `pivot.calculated.list` | `name` | List calculated fields |
| `pivot.calculated.delete` | `name field_name` | Delete calculated field |

## Power Query (4)

| Command | Args | Description |
|---------|------|-------------|
| `powerquery.list` | -- | List all Power Query connections |
| `powerquery.create` | `name formula` | Create M query |
| `powerquery.edit` | `name formula` | Edit existing query |
| `powerquery.refresh` | `name` | Refresh query data |

## DAX & Slicers (4)

| Command | Args | Description |
|---------|------|-------------|
| `slicer.list` | -- | List all slicers |
| `slicer.set` | `name items` | Set slicer selection |
| `slicer.clear` | `name` | Clear slicer |
| `slicer.create` | `pivot field` | Create slicer for pivot field |

## VBA (5)

| Command | Args | Description |
|---------|------|-------------|
| `vba.list` | -- | List VBA modules |
| `vba.view` | `module` | View module source code |
| `vba.import` | `path` | Import .bas/.cls module |
| `vba.update` | `module code` | Replace module source |
| `macro.run` | `name [args]` | Run VBA macro |

## Layout (19)

| Command | Args | Description |
|---------|------|-------------|
| `column.width` | `ref [width]` | Get/set column width |
| `row.height` | `ref [height]` | Get/set row height |
| `autofit` | `ref [target]` | Auto-size columns/rows/both |
| `merge` | `ref` | Merge cells |
| `unmerge` | `ref` | Unmerge |
| `freeze` | `ref` | Freeze panes at cell |
| `unfreeze` | -- | Remove freeze |
| `hide` | `target ref` | Hide rows/columns |
| `unhide` | `target ref` | Show rows/columns |
| `insert.row` | `ref` | Insert row |
| `insert.col` | `ref` | Insert column |
| `delete.row` | `ref` | Delete row |
| `delete.col` | `ref` | Delete column |
| `group` | `ref target` | Group rows/columns |
| `ungroup` | `ref target` | Ungroup |
| `outline.level` | `level target` | Set outline visibility level |
| `used.range` | -- | Get used range bounds |
| `row.count` | -- | Used row count |
| `column.count` | -- | Used column count |

## Data Operations (21)

| Command | Args | Description |
|---------|------|-------------|
| `copy` | `ref` | Copy to clipboard |
| `paste` | `ref` | Paste from clipboard |
| `paste.values` | `ref` | Paste values only |
| `paste.special` | `ref type operation skip_blanks transpose` | Full paste special |
| `sort` | `ref column [order]` | Sort range |
| `find` | `what [ref]` | Find first match |
| `find.all` | `what [ref]` | Find ALL matches (returns array) |
| `replace` | `what with [ref]` | Find and replace |
| `fill.down` | `ref` | Fill down |
| `fill.right` | `ref` | Fill right |
| `fill.series` | `ref type step [stop]` | Fill series (linear/growth/date) |
| `transpose` | `from to` | Transpose range |
| `comment` | `ref text` | Add/replace comment |
| `comment.read` | `ref` | Read cell comment |
| `hyperlink` | `ref url [text]` | Add hyperlink |
| `validation` | `ref type formula` | Data validation (list/whole/decimal/date) |
| `protect` | `[password]` | Protect sheet |
| `unprotect` | `[password]` | Unprotect sheet |
| `undo` | -- | Undo last action |
| `redo` | -- | Redo |
| `error.check` | `ref` | Check cell for errors |

## Workbooks (7)

| Command | Args | Description |
|---------|------|-------------|
| `workbooks` | -- | List open workbooks |
| `workbook.new` | -- | Create blank workbook |
| `workbook.open` | `path [read_only] [password]` | Open file |
| `workbook.save` | -- | Save active workbook |
| `workbook.saveas` | `path [format]` | Save As (xlsx/xlsm/xls/csv/pdf/txt) |
| `workbook.close` | `[save] [name]` | Close workbook |
| `workbook.properties` | `[title author subject]` | Get/set metadata |

## Window (8)

| Command | Args | Description |
|---------|------|-------------|
| `window.zoom` | `[level]` | Get/set zoom (10-400) |
| `window.scroll` | `ref` | Scroll cell into view |
| `window.split` | `[ref]` | Split panes or remove |
| `window.view` | `[mode]` | normal/pagebreak/pagelayout |
| `window.gridlines` | `[show]` | Show/hide gridlines |
| `window.headings` | `[show]` | Show/hide row/column headings |
| `window.statusbar` | -- | Read status bar text |
| `window.fullscreen` | `[on]` | Toggle fullscreen |

## Shapes & Images (9)

| Command | Args | Description |
|---------|------|-------------|
| `shape.list` | -- | List all shapes |
| `shape.add` | `type left top width height [text]` | Add shape (rectangle/oval/line/textbox/callout) |
| `shape.delete` | `name` | Delete shape |
| `shape.text` | `name [text]` | Get/set shape text |
| `shape.move` | `name left top` | Move shape |
| `shape.resize` | `name width height` | Resize |
| `shape.format` | `name fill_color line_color line_weight` | Format |
| `image.insert` | `path left top [width] [height]` | Insert image |
| `image.delete` | `name` | Delete image |

## Print Setup (9)

| Command | Args | Description |
|---------|------|-------------|
| `print.setup` | `orientation paper_size scale fit_wide fit_tall` | Page setup |
| `print.margins` | `top bottom left right header footer` | Margins in inches |
| `print.area` | `[ref]` | Get/set print area |
| `print.area.clear` | -- | Clear print area |
| `print.titles` | `rows columns` | Repeat rows/columns |
| `print.headers` | `left center right` | Header/footer text |
| `print.gridlines` | `[show]` | Print gridlines |
| `print.breaks` | `ref type action` | Page breaks |
| `print.preview` | -- | Get page count + setup |

## External Links (2)

| Command | Args | Description |
|---------|------|-------------|
| `link.list` | -- | List external workbook links |
| `link.update` | -- | Update external links |

## Selection (1)

| Command | Args | Description |
|---------|------|-------------|
| `selection.info` | -- | Active cell + selection range + sheet |

## Task Pane (19)

Requires `XRai.Hooks` NuGet installed in the add-in and `Pilot.Expose()` called on a WPF or WinForms UserControl.

| Command | Args | Description |
|---------|------|-------------|
| `pane` | -- | List ALL controls with types/values/state/details |
| `pane.read` | `control` | Read single control |
| `pane.info` | `control` | Detailed info (DataGrid rows, ComboBox items, etc.) |
| `pane.tree` | -- | Full control tree (same as pane) |
| `pane.type` | `control value` | Type text into TextBox |
| `pane.click` | `control [timeout]` | Click button (timeout:0 for modal-opening buttons) |
| `pane.select` | `control value` | Select ComboBox/ListBox item |
| `pane.toggle` | `control` | Toggle CheckBox/ToggleButton/Expander |
| `pane.tab` | `control tab` | Switch TabControl tab (by index or name) |
| `pane.double_click` | `control` | Double-click control |
| `pane.right_click` | `control` | Right-click (trigger context menu) |
| `pane.hover` | `control` | Hover (trigger tooltip) |
| `pane.focus` | `control` | Set keyboard focus |
| `pane.key` | `control keys` | Send key presses (Enter, Escape, Tab, arrows) |
| `pane.scroll` | `control offset` | Scroll scrollable control |
| `pane.grid.read` | `control` | Read DataGrid as 2D array |
| `pane.grid.cell` | `control row col` | Read specific grid cell |
| `pane.grid.select` | `control row` | Select DataGrid row |
| `pane.tree.expand` | `control path` | Expand TreeView node by path |

## Pane Advanced (4)

| Command | Args | Description |
|---------|------|-------------|
| `pane.wait` | `control property value [timeout]` | Wait for control state change |
| `pane.screenshot` | -- | Capture just the WPF task pane |
| `pane.context_menu` | `control item` | Right-click and select context menu item |
| `pane.drag` | `from to` | Drag between controls |

## ViewModel (3)

Requires `Pilot.ExposeModel()` called on an `INotifyPropertyChanged` object.

| Command | Args | Description |
|---------|------|-------------|
| `model` | -- | Read ALL ViewModel properties (including collections) |
| `model.set` | `property value` | Set property value |
| `functions` | -- | List registered UDFs with parameter signatures |

## Ribbon & UI Automation (5)

| Command | Args | Description |
|---------|------|-------------|
| `ribbon` | -- | List all ribbon tabs |
| `ribbon.click` | `button [item]` | Click ribbon button (optional menu item) |
| `dialog.read` | -- | Read open modal dialog |
| `dialog.click` | `button` | Click dialog button |
| `ui.tree` | `[depth]` | Dump UI automation tree as JSON |

## Dialog Automation (10)

| Command | Args | Description |
|---------|------|-------------|
| `dialog.wait` | `title [timeout]` | Wait for dialog to appear |
| `dialog.dismiss` | -- | Auto-dismiss with safest button |
| `dialog.auto_click` | `button [title]` | Auto-click a button whenever a dialog appears |
| `win32.dialog.list` | -- | List all Win32 dialog windows |
| `win32.dialog.read` | `title` | Read Win32 dialog structure |
| `win32.dialog.type` | `title text [control_id] [submit]` | Type into Win32 dialog edit |
| `win32.dialog.click` | `button title` | Click Win32 dialog button |
| `folder.dialog.navigate` | `path` | Navigate folder picker to path |
| `folder.dialog.pick` | -- | Confirm folder selection |
| `file.dialog.pick` | `path` | Select file in file picker |

## Excel Auto-Dismiss (1)

| Command | Args | Description |
|---------|------|-------------|
| `excel.autodismiss` | `enabled` | Auto-dismiss OLE action dialogs |

## Desktop Automation (24)

| Command | Args | Description |
|---------|------|-------------|
| `clipboard.read` | -- | Read clipboard text |
| `clipboard.write` | `text` | Write text to clipboard |
| `clipboard.image` | `[path]` | Read clipboard image |
| `process.list` | `[name]` | List running processes |
| `process.start` | `path [args]` | Launch process |
| `process.kill` | `pid` | Kill process by PID |
| `window.list` | -- | List visible windows |
| `window.focus` | `title` | Bring window to front |
| `window.minimize` | `title` | Minimize window |
| `window.maximize` | `title` | Maximize window |
| `window.restore` | `title` | Restore window |
| `window.close` | `title` | Close window |
| `window.move` | `title x y` | Move window |
| `window.size` | `title width height` | Resize window |
| `keys.send` | `keys [window]` | Send keystrokes to window |
| `mouse.click` | `x y [button]` | Click at screen coordinates |
| `mouse.move` | `x y` | Move mouse |
| `mouse.drag` | `x1 y1 x2 y2` | Drag from point to point |
| `app.launch` | `path [args] [wait]` | Launch application |
| `app.attach` | `title` | Attach to application by title |
| `env.get` | `name` | Read environment variable |
| `env.set` | `name value` | Set environment variable |
| `path.resolve` | `path` | Resolve relative/special path |
| `shell.exec` | `command` | Execute shell command |

## Testing & Assertions (16)

| Command | Args | Description |
|---------|------|-------------|
| `assert.cell` | `ref condition [value]` | Assert cell value/formula/format |
| `assert.pane` | `control property value` | Assert pane control state |
| `assert.model` | `property value` | Assert ViewModel property |
| `test.start` | `name` | Begin named test run |
| `test.step` | `name [status]` | Record test step |
| `test.end` | -- | End test run |
| `test.report` | `[format] [path]` | Generate report (html/junit-xml) |
| `screenshot.baseline` | `name [path]` | Save baseline screenshot |
| `screenshot.compare` | `name [threshold]` | Compare current vs baseline |
| `ocr.screen` | `[region]` | OCR the Excel window |
| `ocr.element` | `ref` | OCR a specific element |
| `wait.element` | `query [timeout]` | Wait for UI element to appear |
| `wait.window` | `title [timeout]` | Wait for window to appear |
| `wait.property` | `control property value [timeout]` | Wait for property value |
| `wait.gone` | `query [timeout]` | Wait for element to disappear |
| `wait.idle` | `[timeout]` | Wait for application idle |

## Vision (3)

| Command | Args | Description |
|---------|------|-------------|
| `screenshot` | `[target] [path]` | Capture Excel window as PNG |
| `screenshot.region` | `x y width height [path]` | Capture screen region |
| `screenshot.element` | `ref [path]` | Capture specific element |

## Meta (4)

| Command | Args | Description |
|---------|------|-------------|
| `batch` | `commands[]` | Execute multiple commands in one call |
| `reload` | `[xll]` | Hot-reload add-in |
| `connect` | -- | Attach + auto-create workbook if needed |
| `help` | -- | List all commands |

## CLI Subcommands

Run directly on the binary (not as JSON):

```bash
XRai.Tool.exe doctor          # System requirements check
XRai.Tool.exe init MyAddin    # Scaffold new Excel-DNA add-in project
XRai.Tool.exe kill-excel      # Kill all Excel processes
XRai.Tool.exe daemon-status   # Check daemon status
XRai.Tool.exe daemon-stop     # Stop daemon
XRai.Tool.exe --wait          # Wait for Excel, auto-attach, enter REPL
XRai.Tool.exe --pid 1234      # Attach to specific Excel PID, enter REPL
XRai.Tool.exe --daemon        # Run as persistent daemon
XRai.Tool.exe --no-daemon     # Force direct mode (bypass daemon)
```
