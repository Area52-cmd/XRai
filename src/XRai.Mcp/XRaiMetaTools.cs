using System.ComponentModel;
using ModelContextProtocol.Server;
using XRai.Core;

namespace XRai.Mcp;

// ═══════════════════════════════════════════════════════════════════════
// Tool 1: XRaiConnect
// ═══════════════════════════════════════════════════════════════════════
[McpServerToolType]
public sealed class XRaiConnect : XRaiMetaToolBase
{
    public XRaiConnect(CommandRouter router) : base(router) { }

    [McpServerTool(Name = "xrai_connect")]
    [Description("Connect to Excel, check status, manage session lifecycle. Start every session with 'connect'. Commands: connect, attach, detach, wait, status, ensure.workbook, excel.kill, help, commands, reload, rebuild")]
    public string Execute(
        [Description("Sub-command name (connect, attach, detach, wait, status, ensure.workbook, excel.kill, help, commands, reload, rebuild)")] string command,
        [Description("Command arguments as JSON object, e.g. {\"timeout\":30000}")] string? args = null)
    {
        return Dispatch(command, args);
    }
}

// ═══════════════════════════════════════════════════════════════════════
// Tool 2: XRaiCells
// ═══════════════════════════════════════════════════════════════════════
[McpServerToolType]
public sealed class XRaiCells : XRaiMetaToolBase
{
    public XRaiCells(CommandRouter router) : base(router) { }

    [McpServerTool(Name = "xrai_cells")]
    [Description("Read, write, clear, and format Excel cells and ranges. Commands: read, type, clear, select, format, format.read")]
    public string Execute(
        [Description("Sub-command name (read, type, clear, select, format, format.read)")] string command,
        [Description("Command arguments as JSON object, e.g. {\"ref\":\"A1:D10\"} or {\"ref\":\"A1\",\"value\":\"Hello\"}")] string? args = null)
    {
        return Dispatch(command, args);
    }
}

// ═══════════════════════════════════════════════════════════════════════
// Tool 3: XRaiSheets
// ═══════════════════════════════════════════════════════════════════════
[McpServerToolType]
public sealed class XRaiSheets : XRaiMetaToolBase
{
    public XRaiSheets(CommandRouter router) : base(router) { }

    [McpServerTool(Name = "xrai_sheets")]
    [Description("List, create, rename, delete sheets and named ranges. Commands: sheets, sheet.add, sheet.rename, sheet.delete, goto, names, name.read, name.set, name.delete")]
    public string Execute(
        [Description("Sub-command name (sheets, sheet.add, sheet.rename, sheet.delete, goto, names, name.read, name.set, name.delete)")] string command,
        [Description("Command arguments as JSON object, e.g. {\"name\":\"Sheet2\"} or {\"target\":\"Sheet2\"}")] string? args = null)
    {
        return Dispatch(command, args);
    }
}

// ═══════════════════════════════════════════════════════════════════════
// Tool 4: XRaiWorkbooks
// ═══════════════════════════════════════════════════════════════════════
[McpServerToolType]
public sealed class XRaiWorkbooks : XRaiMetaToolBase
{
    public XRaiWorkbooks(CommandRouter router) : base(router) { }

    [McpServerTool(Name = "xrai_workbooks")]
    [Description("Open, save, close, and list Excel workbooks. Commands: workbooks, workbook.new, workbook.open, workbook.save, workbook.saveas, workbook.close, workbook.properties")]
    public string Execute(
        [Description("Sub-command name (workbooks, workbook.new, workbook.open, workbook.save, workbook.saveas, workbook.close, workbook.properties)")] string command,
        [Description("Command arguments as JSON object, e.g. {\"path\":\"C:/file.xlsx\"}")] string? args = null)
    {
        return Dispatch(command, args);
    }
}

// ═══════════════════════════════════════════════════════════════════════
// Tool 5: XRaiFormat
// ═══════════════════════════════════════════════════════════════════════
[McpServerToolType]
public sealed class XRaiFormat : XRaiMetaToolBase
{
    public XRaiFormat(CommandRouter router) : base(router) { }

    [McpServerTool(Name = "xrai_format")]
    [Description("Borders, alignment, fonts, styles, conditional formatting. Commands: format.border, format.align, format.font, format.style, format.conditional, format.conditional.read, format.conditional.clear, format.style.list")]
    public string Execute(
        [Description("Sub-command name (format.border, format.align, format.font, format.style, format.conditional, format.conditional.read, format.conditional.clear, format.style.list)")] string command,
        [Description("Command arguments as JSON object, e.g. {\"ref\":\"A1:D1\",\"side\":\"all\",\"weight\":\"medium\"}")] string? args = null)
    {
        return Dispatch(command, args);
    }
}

// ═══════════════════════════════════════════════════════════════════════
// Tool 6: XRaiLayout
// ═══════════════════════════════════════════════════════════════════════
[McpServerToolType]
public sealed class XRaiLayout : XRaiMetaToolBase
{
    public XRaiLayout(CommandRouter router) : base(router) { }

    [McpServerTool(Name = "xrai_layout")]
    [Description("Row/column sizing, merge/split, freeze, insert/delete. Commands: column.width, row.height, autofit, merge, unmerge, freeze, unfreeze, hide, unhide, insert.row, insert.col, delete.row, delete.col, column.count, row.count, used.range")]
    public string Execute(
        [Description("Sub-command name (column.width, row.height, autofit, merge, unmerge, freeze, unfreeze, hide, unhide, insert.row, insert.col, delete.row, delete.col, column.count, row.count, used.range)")] string command,
        [Description("Command arguments as JSON object, e.g. {\"ref\":\"A:C\"} or {\"ref\":\"1:5\"}")] string? args = null)
    {
        return Dispatch(command, args);
    }
}

// ═══════════════════════════════════════════════════════════════════════
// Tool 7: XRaiData
// ═══════════════════════════════════════════════════════════════════════
[McpServerToolType]
public sealed class XRaiData : XRaiMetaToolBase
{
    public XRaiData(CommandRouter router) : base(router) { }

    [McpServerTool(Name = "xrai_data")]
    [Description("Clipboard, sort, find/replace, fill, comments, validation, hyperlinks. Commands: copy, paste, paste.values, paste.special, sort, find, find.all, replace, fill.down, fill.right, fill.series, transpose, protect, unprotect, comment, comment.read, comment.thread, comment.thread.read, validation, validation.read, hyperlink, undo, redo")]
    public string Execute(
        [Description("Sub-command name (copy, paste, paste.values, paste.special, sort, find, find.all, replace, fill.down, fill.right, fill.series, transpose, protect, unprotect, comment, comment.read, comment.thread, comment.thread.read, validation, validation.read, hyperlink, undo, redo)")] string command,
        [Description("Command arguments as JSON object, e.g. {\"ref\":\"A1:D10\",\"text\":\"hello\"}")] string? args = null)
    {
        return Dispatch(command, args);
    }
}

// ═══════════════════════════════════════════════════════════════════════
// Tool 8: XRaiTables
// ═══════════════════════════════════════════════════════════════════════
[McpServerToolType]
public sealed class XRaiTables : XRaiMetaToolBase
{
    public XRaiTables(CommandRouter router) : base(router) { }

    [McpServerTool(Name = "xrai_tables")]
    [Description("Create and manage Excel tables. Commands: table.list, table.create, table.delete, table.style, table.resize, table.totals, table.filter, table.filter.clear, table.sort, table.row.add, table.column.add, table.data")]
    public string Execute(
        [Description("Sub-command name (table.list, table.create, table.delete, table.style, table.resize, table.totals, table.filter, table.filter.clear, table.sort, table.row.add, table.column.add, table.data)")] string command,
        [Description("Command arguments as JSON object, e.g. {\"ref\":\"A1:D10\",\"name\":\"MyTable\"}")] string? args = null)
    {
        return Dispatch(command, args);
    }
}

// ═══════════════════════════════════════════════════════════════════════
// Tool 9: XRaiFilters
// ═══════════════════════════════════════════════════════════════════════
[McpServerToolType]
public sealed class XRaiFilters : XRaiMetaToolBase
{
    public XRaiFilters(CommandRouter router) : base(router) { }

    [McpServerTool(Name = "xrai_filters")]
    [Description("AutoFilter operations on sheets. Commands: filter.on, filter.off, filter.set, filter.clear, filter.read")]
    public string Execute(
        [Description("Sub-command name (filter.on, filter.off, filter.set, filter.clear, filter.read)")] string command,
        [Description("Command arguments as JSON object, e.g. {\"ref\":\"A1:D10\",\"column\":1,\"criteria\":\">=100\"}")] string? args = null)
    {
        return Dispatch(command, args);
    }
}

// ═══════════════════════════════════════════════════════════════════════
// Tool 10: XRaiCharts
// ═══════════════════════════════════════════════════════════════════════
[McpServerToolType]
public sealed class XRaiCharts : XRaiMetaToolBase
{
    public XRaiCharts(CommandRouter router) : base(router) { }

    [McpServerTool(Name = "xrai_charts")]
    [Description("Charts and sparklines. Commands: chart.list, chart.create, chart.delete, chart.type, chart.title, chart.data, chart.legend, chart.axis, chart.series, chart.export, sparkline.list, sparkline.create, sparkline.delete")]
    public string Execute(
        [Description("Sub-command name (chart.list, chart.create, chart.delete, chart.type, chart.title, chart.data, chart.legend, chart.axis, chart.series, chart.export, sparkline.list, sparkline.create, sparkline.delete)")] string command,
        [Description("Command arguments as JSON object, e.g. {\"type\":\"column\",\"data\":\"A1:B10\",\"title\":\"My Chart\"}")] string? args = null)
    {
        return Dispatch(command, args);
    }
}

// ═══════════════════════════════════════════════════════════════════════
// Tool 11: XRaiPivots
// ═══════════════════════════════════════════════════════════════════════
[McpServerToolType]
public sealed class XRaiPivots : XRaiMetaToolBase
{
    public XRaiPivots(CommandRouter router) : base(router) { }

    [McpServerTool(Name = "xrai_pivots")]
    [Description("PivotTables. Commands: pivot.list, pivot.create, pivot.refresh, pivot.field.add, pivot.field.remove, pivot.style, pivot.data")]
    public string Execute(
        [Description("Sub-command name (pivot.list, pivot.create, pivot.refresh, pivot.field.add, pivot.field.remove, pivot.style, pivot.data)")] string command,
        [Description("Command arguments as JSON object, e.g. {\"source\":\"A1:D100\",\"dest\":\"F1\"}")] string? args = null)
    {
        return Dispatch(command, args);
    }
}

// ═══════════════════════════════════════════════════════════════════════
// Tool 12: XRaiPrint
// ═══════════════════════════════════════════════════════════════════════
[McpServerToolType]
public sealed class XRaiPrint : XRaiMetaToolBase
{
    public XRaiPrint(CommandRouter router) : base(router) { }

    [McpServerTool(Name = "xrai_print")]
    [Description("Print layout configuration. Commands: print.setup, print.margins, print.area, print.area.clear, print.titles, print.headers, print.gridlines, print.breaks, print.preview")]
    public string Execute(
        [Description("Sub-command name (print.setup, print.margins, print.area, print.area.clear, print.titles, print.headers, print.gridlines, print.breaks, print.preview)")] string command,
        [Description("Command arguments as JSON object, e.g. {\"orientation\":\"landscape\"}")] string? args = null)
    {
        return Dispatch(command, args);
    }
}

// ═══════════════════════════════════════════════════════════════════════
// Tool 13: XRaiShapes
// ═══════════════════════════════════════════════════════════════════════
[McpServerToolType]
public sealed class XRaiShapes : XRaiMetaToolBase
{
    public XRaiShapes(CommandRouter router) : base(router) { }

    [McpServerTool(Name = "xrai_shapes")]
    [Description("Shapes and images. Commands: shape.list, shape.add, shape.delete, shape.text, shape.move, shape.resize, shape.format, image.insert, image.delete")]
    public string Execute(
        [Description("Sub-command name (shape.list, shape.add, shape.delete, shape.text, shape.move, shape.resize, shape.format, image.insert, image.delete)")] string command,
        [Description("Command arguments as JSON object, e.g. {\"type\":\"rectangle\",\"left\":100,\"top\":100,\"width\":200,\"height\":100}")] string? args = null)
    {
        return Dispatch(command, args);
    }
}

// ═══════════════════════════════════════════════════════════════════════
// Tool 14: XRaiUI
// ═══════════════════════════════════════════════════════════════════════
[McpServerToolType]
public sealed class XRaiUI : XRaiMetaToolBase
{
    public XRaiUI(CommandRouter router) : base(router) { }

    [McpServerTool(Name = "xrai_ui")]
    [Description("Ribbon, dialogs, screenshots, window settings. Commands: ribbon, ribbon.tabs, ribbon.buttons, ribbon.buttons.all, ribbon.activate, ribbon.tab.activate, ribbon.click, dialog.read, dialog.click, dialog.dismiss, dialog.wait, dialog.list, ui.tree, screenshot, window.zoom, window.scroll, window.split, window.view, window.gridlines, window.headings, window.statusbar, window.fullscreen, folder.dialog.pick, folder.dialog.navigate, file.dialog.pick")]
    public string Execute(
        [Description("Sub-command name (ribbon, ribbon.tabs, ribbon.buttons, ribbon.buttons.all, ribbon.activate, ribbon.tab.activate, ribbon.click, dialog.read, dialog.click, dialog.dismiss, dialog.wait, dialog.list, ui.tree, screenshot, window.zoom, window.scroll, window.split, window.view, window.gridlines, window.headings, window.statusbar, window.fullscreen, folder.dialog.pick, folder.dialog.navigate, file.dialog.pick)")] string command,
        [Description("Command arguments as JSON object, e.g. {\"tab\":\"Home\"} or {\"depth\":2}")] string? args = null)
    {
        return Dispatch(command, args);
    }
}

// ═══════════════════════════════════════════════════════════════════════
// Tool 15: XRaiWin32
// ═══════════════════════════════════════════════════════════════════════
[McpServerToolType]
public sealed class XRaiWin32 : XRaiMetaToolBase
{
    public XRaiWin32(CommandRouter router) : base(router) { }

    [McpServerTool(Name = "xrai_win32")]
    [Description("Win32 native dialog automation and autodismiss watchdog. Commands: win32.dialog.list, win32.dialog.dismiss, win32.dialog.click, win32.dialog.type, win32.dialog.read, excel.autodismiss, excel.autodismiss.status")]
    public string Execute(
        [Description("Sub-command name (win32.dialog.list, win32.dialog.dismiss, win32.dialog.click, win32.dialog.type, win32.dialog.read, excel.autodismiss, excel.autodismiss.status)")] string command,
        [Description("Command arguments as JSON object, e.g. {\"button\":\"OK\"}")] string? args = null)
    {
        return Dispatch(command, args);
    }
}

// ═══════════════════════════════════════════════════════════════════════
// Tool 16: XRaiPane
// ═══════════════════════════════════════════════════════════════════════
[McpServerToolType]
public sealed class XRaiPane : XRaiMetaToolBase
{
    public XRaiPane(CommandRouter router) : base(router) { }

    [McpServerTool(Name = "xrai_pane")]
    [Description("Drive WPF/WinForms task pane controls, ViewModel, UDFs. Commands: pane, pane.status, pane.type, pane.click, pane.select, pane.toggle, pane.read, pane.double_click, pane.right_click, pane.hover, pane.focus, pane.key, pane.scroll, pane.info, pane.tree, pane.grid.read, pane.grid.cell, pane.grid.select, pane.tree.expand, pane.tab, pane.list.read, pane.list.select, pane.expand, pane.wait, pane.screenshot, pane.drag, pane.context_menu, model, model.set, functions")]
    public string Execute(
        [Description("Sub-command name (pane, pane.status, pane.type, pane.click, pane.select, pane.toggle, pane.read, pane.double_click, pane.right_click, pane.hover, pane.focus, pane.key, pane.scroll, pane.info, pane.tree, pane.grid.read, pane.grid.cell, pane.grid.select, pane.tree.expand, pane.tab, pane.list.read, pane.list.select, pane.expand, pane.wait, pane.screenshot, pane.drag, pane.context_menu, model, model.set, functions)")] string command,
        [Description("Command arguments as JSON object, e.g. {\"control\":\"CalcButton\"} or {\"property\":\"Spot\",\"value\":105.0}")] string? args = null)
    {
        return Dispatch(command, args);
    }
}

// ═══════════════════════════════════════════════════════════════════════
// Tool 17: XRaiDesktop
// ═══════════════════════════════════════════════════════════════════════
[McpServerToolType]
public sealed class XRaiDesktop : XRaiMetaToolBase
{
    public XRaiDesktop(CommandRouter router) : base(router) { }

    [McpServerTool(Name = "xrai_desktop")]
    [Description("General Windows desktop automation: clipboard, process management, window control, keyboard/mouse input, system info, app launch/attach. Commands: clipboard.read, clipboard.write, clipboard.clear, clipboard.formats, process.list, process.start, process.kill, process.wait, process.info, window.list, window.move, window.minimize, window.maximize, window.restore, window.close, window.focus, keys.send, mouse.click, mouse.move, mouse.scroll, system.info, app.launch, app.list, app.attach")]
    public string Execute(
        [Description("Sub-command name (clipboard.read, clipboard.write, clipboard.clear, clipboard.formats, process.list, process.start, process.kill, process.wait, process.info, window.list, window.move, window.minimize, window.maximize, window.restore, window.close, window.focus, keys.send, mouse.click, mouse.move, mouse.scroll, system.info, app.launch, app.list, app.attach)")] string command,
        [Description("Command arguments as JSON object, e.g. {\"text\":\"hello\"} or {\"pid\":1234}")] string? args = null)
    {
        return Dispatch(command, args);
    }
}

// ═══════════════════════════════════════════════════════════════════════
// Tool 18: XRaiCalc
// ═══════════════════════════════════════════════════════════════════════
[McpServerToolType]
public sealed class XRaiCalc : XRaiMetaToolBase
{
    public XRaiCalc(CommandRouter router) : base(router) { }

    [McpServerTool(Name = "xrai_calc")]
    [Description("Calculation, timing, macros, grouping. Commands: calc, calc.mode, wait.calc, wait.cell, time.calc, time.udf, selection.info, error.check, link.list, link.update, macro.run, group, ungroup, outline.level")]
    public string Execute(
        [Description("Sub-command name (calc, calc.mode, wait.calc, wait.cell, time.calc, time.udf, selection.info, error.check, link.list, link.update, macro.run, group, ungroup, outline.level)")] string command,
        [Description("Command arguments as JSON object, e.g. {\"mode\":\"manual\"} or {\"macro\":\"MyMacro\"}")] string? args = null)
    {
        return Dispatch(command, args);
    }
}

// ═══════════════════════════════════════════════════════════════════════
// Tool 19: XRaiTest
// ═══════════════════════════════════════════════════════════════════════
[McpServerToolType]
public sealed class XRaiTest : XRaiMetaToolBase
{
    public XRaiTest(CommandRouter router) : base(router) { }

    [McpServerTool(Name = "xrai_test")]
    [Description("Testing, assertions, OCR, screenshot diff, intelligent waits, and test reporting (HTML/JUnit XML). Commands: test.start, test.step, test.assert, test.end, test.report, assert.cell, assert.pane, assert.model, screenshot.baseline, screenshot.compare, ocr.screen, ocr.element, wait.element, wait.window, wait.property, wait.gone")]
    public string Execute(
        [Description("Sub-command name (test.start, test.step, test.assert, test.end, test.report, assert.cell, assert.pane, assert.model, screenshot.baseline, screenshot.compare, ocr.screen, ocr.element, wait.element, wait.window, wait.property, wait.gone)")] string command,
        [Description("Command arguments as JSON object, e.g. {\"name\":\"My Test\"} or {\"ref\":\"A1\",\"value\":\"100\"}")] string? args = null)
    {
        return Dispatch(command, args);
    }
}
