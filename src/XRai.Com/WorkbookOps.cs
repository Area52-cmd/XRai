// Leak-audited: 2026-04-10 — Fixed two leaks in HandleList: the Sheets COM
// proxy used for sheet_count was never released (one leak per workbook per
// call), and the ActiveWorkbook getter result was discarded without release
// (one leak per call). Both now release the proxy explicitly. The workbook
// loop variable in HandleList and HandleClose is released on every iteration
// — including the early-return path inside HandleClose.

using System.IO;
using System.Runtime.InteropServices;
using System.Text.Json.Nodes;
using XRai.Core;
using Excel = Microsoft.Office.Interop.Excel;

namespace XRai.Com;

public class WorkbookOps
{
    private readonly ExcelSession _session;
    private readonly IDialogWatchdog? _watchdog;

    public WorkbookOps(ExcelSession session, IDialogWatchdog? watchdog = null)
    {
        _session = session;
        _watchdog = watchdog;
    }

    public void Register(CommandRouter router)
    {
        router.Register("workbooks", HandleList);
        router.Register("workbook.new", HandleNew);
        router.Register("workbook.open", HandleOpen);
        router.Register("workbook.save", HandleSave);
        router.Register("workbook.saveas", HandleSaveAs);
        router.Register("workbook.close", HandleClose);
        router.Register("workbook.properties", HandleProperties);
    }

    // Leak-audited: 2026-04-10. The previous implementation called
    // wb.Sheets.Count inline, which materialized a Sheets COM proxy that was
    // never released — once per workbook, leaked on every workbook.list call.
    // Now we track the Sheets reference explicitly and release it before
    // dropping the workbook proxy.
    private string HandleList(JsonObject args)
    {
        using var guard = new ComGuard();
        var workbooks = guard.Track(_session.App.Workbooks);

        var result = new JsonArray();
        foreach (Excel.Workbook wb in workbooks)
        {
            Excel.Sheets? sheets = null;
            int sheetCount = 0;
            try
            {
                sheets = wb.Sheets;
                sheetCount = sheets.Count;
            }
            catch { /* leave count at 0 if Sheets accessor throws */ }

            result.Add(new JsonObject
            {
                ["name"] = wb.Name,
                ["path"] = wb.FullName,
                ["saved"] = wb.Saved,
                ["read_only"] = wb.ReadOnly,
                ["sheet_count"] = sheetCount,
            });

            if (sheets != null) Marshal.ReleaseComObject(sheets);
            Marshal.ReleaseComObject(wb);
        }

        // The ActiveWorkbook getter returns a fresh COM proxy that must be
        // released — previously this was a one-shot leak per workbook.list call.
        Excel.Workbook? activeWb = null;
        string? active = null;
        try
        {
            activeWb = _session.App.ActiveWorkbook;
            active = activeWb?.Name;
        }
        catch { }
        finally
        {
            if (activeWb != null) Marshal.ReleaseComObject(activeWb);
        }

        return Response.Ok(new { workbooks = result, active });
    }

    private string HandleNew(JsonObject args)
    {
        using var guard = new ComGuard();
        var workbooks = guard.Track(_session.App.Workbooks);
        var wb = guard.Track(workbooks.Add());

        // Reset zoom and view to sane defaults — fixes bug where a machine-template
        // stuck at 10% zoom leaves new workbooks unusably tiny.
        try
        {
            var window = guard.Track(_session.App.ActiveWindow);
            try { window.Zoom = 100; } catch { }
            try { window.View = Excel.XlWindowView.xlNormalView; } catch { }
        }
        catch { /* not fatal if we can't set view state */ }

        return Response.Ok(new { name = wb.Name, created = true, zoom_reset = true });
    }

    private string HandleOpen(JsonObject args)
    {
        var path = args["path"]?.GetValue<string>()
            ?? throw new ArgumentException("workbook.open requires 'path'");
        var readOnly = args["read_only"]?.GetValue<bool>() ?? false;
        var password = args["password"]?.GetValue<string>();

        // Resolve and verify the file up front — Excel's own "file not found"
        // dialog is a NUIDialog that strands the STA thread just like the
        // others, and the error message is far better from C# land.
        string fullPath;
        try { fullPath = Path.GetFullPath(path); }
        catch (Exception ex) { return Response.Error($"Invalid path '{path}': {ex.Message}"); }
        if (!File.Exists(fullPath))
            return Response.Error($"File not found: {fullPath}");

        using var guard = new ComGuard();
        var app = _session.App;
        var workbooks = guard.Track(app.Workbooks);

        // ── Step 1: Suppress every dialog Excel can pop during Open ──
        // These three flags kill the overwhelming majority of NUIDialogs
        // that block Workbooks.Open behind a modal:
        //   DisplayAlerts=false       — suppresses "file format doesn't match", recovery prompts
        //   AskToUpdateLinks=false    — suppresses "This workbook contains links" prompt
        //   AlertBeforeOverwriting=false — suppresses paste-overwrite prompts that can fire on open
        // We snapshot the previous values and restore them in the finally block
        // so we don't change global Excel behavior beyond this single call.
        bool prevDisplayAlerts = true;
        bool prevAskToUpdate = true;
        bool prevAlertBeforeOverwriting = true;
        try { prevDisplayAlerts = app.DisplayAlerts; } catch { }
        try { prevAskToUpdate = app.AskToUpdateLinks; } catch { }
        try { prevAlertBeforeOverwriting = app.AlertBeforeOverwriting; } catch { }

        try { app.DisplayAlerts = false; } catch { }
        try { app.AskToUpdateLinks = false; } catch { }
        try { app.AlertBeforeOverwriting = false; } catch { }

        // ── Step 2: Start the dialog watchdog for the duration of Open ──
        // Even with every alert flag off, Excel can still throw:
        //   - Protected View "Enable Editing" bar promoted to modal on some
        //     files from the Internet zone
        //   - Security Warning for macro-enabled files from untrusted locations
        //   - "File in use" / "locked for editing" NUIDialog when another
        //     process holds the file open
        //   - Document Recovery pane that sometimes latches as a blocking modal
        // The watchdog polls the Excel process's top-level windows on a
        // background thread (Win32 APIs, apartment-agnostic) and clicks the
        // safest button on any NUIDialog it finds, unblocking the STA call.
        bool startedWatchdog = false;
        if (_watchdog != null)
        {
            _watchdog.EnableWatchdog(150);  // tight interval during open
            startedWatchdog = true;
        }

        Excel.Workbook wb;
        try
        {
            // ── Step 3: Call Workbooks.Open with every safety parameter set ──
            // UpdateLinks=0                    — do not update external links on open (0 = no prompt, no update)
            // ReadOnly                          — caller-provided
            // Format=5 (Nothing)                — use extension-based auto-detect
            // Password                          — caller-provided or null
            // WriteResPassword=null             — no write-res password
            // IgnoreReadOnlyRecommended=true    — skip the "open as read-only?" prompt
            // Origin=null                       — no origin override
            // Delimiter=null                    — no delimiter (CSVs)
            // Editable=null                     — default
            // Notify=false                      — if file is locked, FAIL instead of notifying (no modal)
            // Converter=null                    — no converter override
            // AddToMru=false                    — don't pollute MRU with automation opens
            // Local=false                       — use Excel's UI language for parsing
            // CorruptLoad=0 (xlNormalLoad)      — don't enter recovery mode on corrupt files
            object passwordArg = (object?)password ?? Type.Missing;
            wb = guard.Track(workbooks.Open(
                Filename: fullPath,
                UpdateLinks: 0,
                ReadOnly: readOnly,
                Format: Type.Missing,
                Password: passwordArg,
                WriteResPassword: Type.Missing,
                IgnoreReadOnlyRecommended: true,
                Origin: Type.Missing,
                Delimiter: Type.Missing,
                Editable: Type.Missing,
                Notify: false,
                Converter: Type.Missing,
                AddToMru: false,
                Local: false,
                CorruptLoad: 0  // xlNormalLoad
            ));
        }
        finally
        {
            // ── Step 4: Tear down watchdog + restore Excel state ──
            if (startedWatchdog)
            {
                try { _watchdog!.DisableWatchdog(); } catch { }
            }
            try { app.DisplayAlerts = prevDisplayAlerts; } catch { }
            try { app.AskToUpdateLinks = prevAskToUpdate; } catch { }
            try { app.AlertBeforeOverwriting = prevAlertBeforeOverwriting; } catch { }
        }

        return Response.Ok(new { name = wb.Name, path = wb.FullName, opened = true });
    }

    private string HandleSave(JsonObject args)
    {
        using var guard = new ComGuard();
        var wb = guard.Track(_session.GetActiveWorkbook());
        wb.Save();
        return Response.Ok(new { name = wb.Name, saved = true });
    }

    private string HandleSaveAs(JsonObject args)
    {
        var path = args["path"]?.GetValue<string>()
            ?? throw new ArgumentException("workbook.saveas requires 'path'");
        var format = args["format"]?.GetValue<string>() ?? "xlsx";

        var xlFormat = format.ToLowerInvariant() switch
        {
            "xlsx" => Excel.XlFileFormat.xlOpenXMLWorkbook,
            "xlsm" => Excel.XlFileFormat.xlOpenXMLWorkbookMacroEnabled,
            "xls" => Excel.XlFileFormat.xlWorkbookNormal,
            "csv" => Excel.XlFileFormat.xlCSV,
            "pdf" => (Excel.XlFileFormat)57, // xlTypePDF
            "txt" => Excel.XlFileFormat.xlTextWindows,
            _ => Excel.XlFileFormat.xlOpenXMLWorkbook,
        };

        using var guard = new ComGuard();
        var wb = guard.Track(_session.GetActiveWorkbook());

        _session.App.DisplayAlerts = false;
        try
        {
            wb.SaveAs(path, xlFormat);
        }
        finally
        {
            _session.App.DisplayAlerts = true;
        }

        return Response.Ok(new { name = wb.Name, path = wb.FullName, format, saved = true });
    }

    private string HandleClose(JsonObject args)
    {
        var saveChanges = args["save"]?.GetValue<bool>() ?? false;
        var name = args["name"]?.GetValue<string>();

        using var guard = new ComGuard();

        if (name != null)
        {
            var workbooks = guard.Track(_session.App.Workbooks);
            foreach (Excel.Workbook wb in workbooks)
            {
                if (string.Equals(wb.Name, name, StringComparison.OrdinalIgnoreCase))
                {
                    wb.Close(saveChanges);
                    Marshal.ReleaseComObject(wb);
                    return Response.Ok(new { name, closed = true });
                }
                Marshal.ReleaseComObject(wb);
            }
            return Response.Error($"Workbook not found: {name}");
        }

        var active = guard.Track(_session.GetActiveWorkbook());
        var activeName = active.Name;
        active.Close(saveChanges);
        return Response.Ok(new { name = activeName, closed = true });
    }

    private string HandleProperties(JsonObject args)
    {
        using var guard = new ComGuard();
        var wb = guard.Track(_session.GetActiveWorkbook());

        // Read built-in properties
        dynamic props = wb.BuiltinDocumentProperties;
        string? title = null, author = null, subject = null, keywords = null, comments = null;
        try { title = props["Title"].Value?.ToString(); } catch { }
        try { author = props["Author"].Value?.ToString(); } catch { }
        try { subject = props["Subject"].Value?.ToString(); } catch { }
        try { keywords = props["Keywords"].Value?.ToString(); } catch { }
        try { comments = props["Comments"].Value?.ToString(); } catch { }

        // Set properties if provided
        if (args["title"] != null) try { props["Title"].Value = args["title"]!.GetValue<string>(); } catch { }
        if (args["author"] != null) try { props["Author"].Value = args["author"]!.GetValue<string>(); } catch { }
        if (args["subject"] != null) try { props["Subject"].Value = args["subject"]!.GetValue<string>(); } catch { }

        return Response.Ok(new { name = wb.Name, title, author, subject, keywords, comments });
    }
}
