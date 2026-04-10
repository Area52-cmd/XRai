# WPF Control Support Matrix

XRai.Hooks walks the WPF visual tree via `Pilot.Expose(element)` and registers every control with `x:Name="..."`. Controls without `x:Name` are invisible.

## Supported Controls

| Control | Read | Write | Click | Notes |
|---------|------|-------|-------|-------|
| `TextBox` | ✅ | ✅ | — | `Text` property |
| `Button` | — | — | ✅ | Click event |
| `Label` | ✅ | — | — | `Content` property |
| `ComboBox` | ✅ | ✅ | — | `SelectedItem` + items list |
| `CheckBox` | ✅ | ✅ | ✅ | `IsChecked` toggle |
| `RadioButton` | ✅ | ✅ | — | `IsChecked` within group |
| `DataGrid` | ✅ | — | ✅ | Cell read, row select, full data export |
| `TabControl` | ✅ | ✅ | — | Tab switching by index or header name |
| `ListView` | ✅ | ✅ | — | Item selection |
| `ListBox` | ✅ | ✅ | — | Item selection |
| `TreeView` | ✅ | — | — | Node expansion by path |
| `DatePicker` | ✅ | ✅ | — | `SelectedDate` (yyyy-MM-dd) |
| `Slider` | ✅ | ✅ | — | Numeric value |
| `ProgressBar` | ✅ | — | — | Value, min, max, percentage |
| `Expander` | ✅ | ✅ | — | `IsExpanded` toggle |
| `RichTextBox` | ✅ | ✅ | — | Plain text extraction |
| `PasswordBox` | — | ✅ | — | Write-only (`Password`) |
| `ScrollViewer` | — | — | — | Scroll offset control |

## Interaction Commands

All `pane.*` commands target controls by their `x:Name`:

```json
{"cmd":"pane.read","control":"SpotInput"}
{"cmd":"pane.type","control":"SpotInput","value":"105.0"}
{"cmd":"pane.click","control":"CalcButton"}
{"cmd":"pane.toggle","control":"EnableCheck"}
{"cmd":"pane.select","control":"InstrumentCombo","value":"Option"}
```

## Mouse Simulation

```json
{"cmd":"pane.double_click","control":"HoldingsGrid"}
{"cmd":"pane.right_click","control":"CellValue"}
{"cmd":"pane.hover","control":"HelpIcon"}
```

## Keyboard & Focus

```json
{"cmd":"pane.focus","control":"SpotInput"}
{"cmd":"pane.key","control":"SpotInput","keys":"Enter"}
{"cmd":"pane.key","control":"SpotInput","keys":"Control+A"}
```

## DataGrid Operations

```json
{"cmd":"pane.grid.read","control":"HoldingsGrid"}
→ {"ok":true,"name":"HoldingsGrid","info":{"rows":8,"columns":6,"column_headers":["Symbol","Qty","Price","Value","P&L","%"]},"data":[["AAPL","50","178.25","8912.5","1787.5","25.1"], ...]}

{"cmd":"pane.grid.cell","control":"HoldingsGrid","row":0,"col":2}
→ {"ok":true,"value":"178.25"}

{"cmd":"pane.grid.select","control":"HoldingsGrid","row":2}
```

## TabControl Switching

```json
{"cmd":"pane.tab","control":"MainTabs","tab":"Holdings"}  // by header name
{"cmd":"pane.tab","control":"MainTabs","tab":"1"}          // by index
```

## TreeView Navigation

```json
{"cmd":"pane.tree.expand","control":"SectorTree","path":"Technology/Semiconductors"}
```

## Critical: Name Your Controls

```xml
<!-- ✅ XRai can find this -->
<TextBox x:Name="SpotInput" Text="{Binding Spot}" />
<Button x:Name="CalcButton" Content="Calculate" Click="Calc_Click" />
<DataGrid x:Name="TradesGrid" ItemsSource="{Binding Trades}" />
<TabControl x:Name="MainTabs">
    <TabItem Header="Holdings">...</TabItem>
    <TabItem Header="Trade">...</TabItem>
</TabControl>

<!-- ❌ XRai cannot see this (no x:Name) -->
<TextBox Text="{Binding Spot}" />
```

Auto-generated control names like `templateRoot`, `PART_SelectedContentHost`, `contentPresenter` are visible but noisy — always give your interactive controls explicit names.

## Detailed Info Responses

`pane.info control` returns structured info per control type:

```json
// DataGrid
{"rows":8,"columns":6,"selected_index":-1,"column_headers":["Symbol","Qty",...]}

// TabControl
{"selected_index":0,"tab_count":4,"tabs":["Holdings","Trade","History","Settings"]}

// ComboBox
{"selected_index":0,"item_count":4,"items":["Personal","IRA","401k","Joint"]}

// Slider
{"value":50,"minimum":0,"maximum":100}

// ProgressBar
{"value":75,"minimum":0,"maximum":100,"percentage":75}

// ListBox
{"item_count":10,"selected_index":3,"items":[...]}
```
