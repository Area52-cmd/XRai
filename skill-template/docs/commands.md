# XRai Command Reference

Complete catalog of JSON commands XRai accepts on stdin. Responses come on stdout as `{"ok":true,...}` or `{"ok":false,"error":"..."}`.

## Connection

| Command | Args | Description |
|---------|------|-------------|
| `attach` | `[pid]` | Connect to running Excel (optional specific PID) |
| `wait` | — | Block until Excel appears, then attach |
| `detach` | — | Disconnect from Excel |
| `status` | — | Attached state, version, hooks status, pipe name |
| `doctor` | — | Diagnose system requirements (CLI subcommand) |

## Cells (COM)

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

## Sheets

| Command | Args | Description |
|---------|------|-------------|
| `sheets` | — | List all sheets with index, visibility, active |
| `sheet.add` | `name [after]` | Add new sheet |
| `sheet.rename` | `from to` | Rename sheet |
| `sheet.delete` | `name` | Delete (suppresses confirm dialog) |
| `goto` | `target` | Activate sheet or jump to named range |

## Named Ranges

| Command | Args | Description |
|---------|------|-------------|
| `names` | — | List all named ranges |
| `name.set` | `name ref` | Create named range |
| `name.read` | `name` | Read named range details |
| `name.delete` | `name` | Delete named range |

## Calculation

| Command | Args | Description |
|---------|------|-------------|
| `calc` | — | Full recalculation |
| `calc.mode` | `[mode]` | Get/set auto/manual/semiautomatic |
| `wait.calc` | `[timeout]` | Wait for calc to finish (ms timeout) |
| `wait.cell` | `ref condition [value] [timeout]` | Wait for cell condition (not_empty/equals/not_error) |
| `time.calc` | — | Benchmark recalc duration |
| `time.udf` | `function args N` | Benchmark UDF over N iterations |

## Charts

| Command | Args | Description |
|---------|------|-------------|
| `chart.list` | — | List all charts on active sheet |
| `chart.create` | `type data title` | Create chart (column/bar/line/pie/scatter/area) |
| `chart.delete` | `name` | Delete chart |
| `chart.type` | `name [type]` | Get/set chart type |
| `chart.title` | `name [title]` | Get/set chart title |
| `chart.data` | `name [data]` | Get/set data source |
| `chart.legend` | `name show position` | Legend configuration |
| `chart.axis` | `name which title min max format` | Axis configuration |
| `chart.series` | `name action` | List/add/remove/format series |
| `chart.export` | `name path` | Export as PNG |

## Tables (ListObjects)

| Command | Args | Description |
|---------|------|-------------|
| `table.list` | — | List all tables |
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

## Filters

| Command | Args | Description |
|---------|------|-------------|
| `filter.on` | `ref` | Enable AutoFilter |
| `filter.off` | — | Disable AutoFilter |
| `filter.set` | `column criteria [operator]` | Apply filter criteria |
| `filter.clear` | `[column]` | Clear filter |
| `filter.read` | — | Read current criteria |

## Pivot Tables

| Command | Args | Description |
|---------|------|-------------|
| `pivot.list` | — | List all pivots |
| `pivot.create` | `source destination name` | Create from data |
| `pivot.refresh` | `name` | Refresh data |
| `pivot.field.add` | `name field area [function]` | Add row/column/value/filter field |
| `pivot.field.remove` | `name field` | Remove field |
| `pivot.style` | `name style` | Apply style |
| `pivot.data` | `name` | Read values |

## Layout

| Command | Args | Description |
|---------|------|-------------|
| `column.width` | `ref [width]` | Get/set column width |
| `row.height` | `ref [height]` | Get/set row height |
| `autofit` | `ref [target]` | Auto-size columns/rows/both |
| `merge` | `ref` | Merge cells |
| `unmerge` | `ref` | Unmerge |
| `freeze` | `ref` | Freeze panes at cell |
| `unfreeze` | — | Remove freeze |
| `hide` | `target ref` | Hide rows/columns |
| `unhide` | `target ref` | Show rows/columns |
| `insert.row` | `ref` | Insert row |
| `insert.col` | `ref` | Insert column |
| `delete.row` | `ref` | Delete row |
| `delete.col` | `ref` | Delete column |
| `group` | `ref target` | Group rows/columns |
| `ungroup` | `ref target` | Ungroup |
| `outline.level` | `level target` | Set outline visibility level |
| `used.range` | — | Get used range bounds |
| `row.count` | — | Used row count |
| `column.count` | — | Used column count |

## Data Operations

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
| `undo` | — | Undo last action |
| `redo` | — | Redo |

## Workbooks

| Command | Args | Description |
|---------|------|-------------|
| `workbooks` | — | List open workbooks |
| `workbook.new` | — | Create blank workbook |
| `workbook.open` | `path [read_only] [password]` | Open file |
| `workbook.save` | — | Save active workbook |
| `workbook.saveas` | `path [format]` | Save As (xlsx/xlsm/xls/csv/pdf/txt) |
| `workbook.close` | `[save] [name]` | Close workbook |
| `workbook.properties` | `[title author subject]` | Get/set metadata |

## Window

| Command | Args | Description |
|---------|------|-------------|
| `window.zoom` | `[level]` | Get/set zoom (10-400) |
| `window.scroll` | `ref` | Scroll cell into view |
| `window.split` | `[ref]` | Split panes or remove |
| `window.view` | `[mode]` | normal/pagebreak/pagelayout |
| `window.gridlines` | `[show]` | Show/hide gridlines |
| `window.headings` | `[show]` | Show/hide row/column headings |
| `window.statusbar` | — | Read status bar text |
| `window.fullscreen` | `[on]` | Toggle fullscreen |

## Shapes & Images

| Command | Args | Description |
|---------|------|-------------|
| `shape.list` | — | List all shapes |
| `shape.add` | `type left top width height [text]` | Add shape (rectangle/oval/line/textbox/callout) |
| `shape.delete` | `name` | Delete shape |
| `shape.text` | `name [text]` | Get/set shape text |
| `shape.move` | `name left top` | Move shape |
| `shape.resize` | `name width height` | Resize |
| `shape.format` | `name fill_color line_color line_weight` | Format |
| `image.insert` | `path left top [width] [height]` | Insert image |
| `image.delete` | `name` | Delete image |

## Print Setup

| Command | Args | Description |
|---------|------|-------------|
| `print.setup` | `orientation paper_size scale fit_wide fit_tall` | Page setup |
| `print.margins` | `top bottom left right header footer` | Margins in inches |
| `print.area` | `[ref]` | Get/set print area |
| `print.area.clear` | — | Clear print area |
| `print.titles` | `rows columns` | Repeat rows/columns |
| `print.headers` | `left center right` | Header/footer text |
| `print.gridlines` | `[show]` | Print gridlines |
| `print.breaks` | `ref type action` | Page breaks |
| `print.preview` | — | Get page count + setup |

## Advanced

| Command | Args | Description |
|---------|------|-------------|
| `macro.run` | `name [args]` | Run VBA macro |
| `selection.info` | — | Active cell + selection range + sheet |
| `error.check` | `ref` | Check cell for errors |
| `link.list` | — | List external workbook links |
| `link.update` | — | Update external links |

## Task Pane (Hooks)

Requires `XRai.Hooks` NuGet installed in the add-in and `Pilot.Expose()` called on a WPF UserControl.

| Command | Args | Description |
|---------|------|-------------|
| `pane` | — | List ALL controls with types/values/state/details |
| `pane.read` | `control` | Read single control |
| `pane.info` | `control` | Detailed info (DataGrid rows, ComboBox items, etc.) |
| `pane.tree` | — | Full control tree (same as pane) |
| `pane.type` | `control value` | Type text into TextBox |
| `pane.click` | `control` | Click button |
| `pane.select` | `control value` | Select ComboBox item |
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

## ViewModel (Hooks)

Requires `Pilot.ExposeModel()` called on an `INotifyPropertyChanged` object.

| Command | Args | Description |
|---------|------|-------------|
| `model` | — | Read ALL ViewModel properties (including collections) |
| `model.set` | `property value` | Set property value |
| `functions` | — | List registered UDFs with parameter signatures |

## Ribbon & Dialogs (FlaUI)

| Command | Args | Description |
|---------|------|-------------|
| `ribbon` | — | List all ribbon tabs |
| `ribbon.click` | `button` | Click ribbon button by name |
| `dialog.read` | — | Read open modal dialog |
| `dialog.click` | `button` | Click dialog button |
| `ui.tree` | `[depth]` | Dump UI automation tree as JSON |

## Vision

| Command | Args | Description |
|---------|------|-------------|
| `screenshot` | `[target] [path]` | Capture Excel window as PNG |

## Meta

| Command | Args | Description |
|---------|------|-------------|
| `batch` | `commands[]` | Execute multiple commands in one call |
| `reload` | `[xll]` | Hot-reload add-in |

## CLI Subcommands (not JSON)

Run directly on the binary:
```bash
XRai.Tool.exe doctor        # System requirements check
XRai.Tool.exe init MyAddin  # Scaffold new Excel-DNA add-in project
XRai.Tool.exe --wait        # Wait for Excel, auto-attach, enter REPL
XRai.Tool.exe --pid 1234    # Attach to specific Excel PID, enter REPL
```
