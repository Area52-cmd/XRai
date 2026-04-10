using System.Runtime.InteropServices;
using System.Text.Json.Nodes;
using XRai.Core;
using Excel = Microsoft.Office.Interop.Excel;

namespace XRai.Com;

public class LayoutOps
{
    private readonly ExcelSession _session;

    public LayoutOps(ExcelSession session) { _session = session; }

    public void Register(CommandRouter router)
    {
        router.Register("column.width", HandleColumnWidth);
        router.Register("row.height", HandleRowHeight);
        router.Register("autofit", HandleAutofit);
        router.Register("merge", HandleMerge);
        router.Register("unmerge", HandleUnmerge);
        router.Register("freeze", HandleFreeze);
        router.Register("unfreeze", HandleUnfreeze);
        router.Register("hide", HandleHide);
        router.Register("unhide", HandleUnhide);
        router.Register("insert.row", HandleInsertRow);
        router.Register("insert.col", HandleInsertCol);
        router.Register("delete.row", HandleDeleteRow);
        router.Register("delete.col", HandleDeleteCol);
        router.Register("column.count", HandleColumnCount);
        router.Register("row.count", HandleRowCount);
        router.Register("used.range", HandleUsedRange);
    }

    private string HandleColumnWidth(JsonObject args)
    {
        var refStr = args["ref"]?.GetValue<string>()
            ?? throw new ArgumentException("column.width requires 'ref'");

        using var guard = new ComGuard();
        var sheet = guard.Track(_session.GetActiveSheet());
        var range = guard.Track(sheet.Range[refStr]);
        var cols = guard.Track(range.Columns);

        if (args["width"] != null)
        {
            var width = args["width"]!.GetValue<double>();
            cols.ColumnWidth = width;
            return Response.Ok(new { @ref = refStr, width });
        }

        return Response.Ok(new { @ref = refStr, width = Convert.ToDouble(cols.ColumnWidth) });
    }

    private string HandleRowHeight(JsonObject args)
    {
        var refStr = args["ref"]?.GetValue<string>()
            ?? throw new ArgumentException("row.height requires 'ref'");

        using var guard = new ComGuard();
        var sheet = guard.Track(_session.GetActiveSheet());
        var range = guard.Track(sheet.Range[refStr]);
        var rows = guard.Track(range.Rows);

        if (args["height"] != null)
        {
            var height = args["height"]!.GetValue<double>();
            rows.RowHeight = height;
            return Response.Ok(new { @ref = refStr, height });
        }

        return Response.Ok(new { @ref = refStr, height = Convert.ToDouble(rows.RowHeight) });
    }

    private string HandleAutofit(JsonObject args)
    {
        var refStr = args["ref"]?.GetValue<string>()
            ?? throw new ArgumentException("autofit requires 'ref'");
        var target = args["target"]?.GetValue<string>() ?? "columns";

        using var guard = new ComGuard();
        var sheet = guard.Track(_session.GetActiveSheet());
        var range = guard.Track(sheet.Range[refStr]);

        if (target == "rows" || target == "both")
        {
            var rows = guard.Track(range.Rows);
            rows.AutoFit();
        }
        if (target == "columns" || target == "both")
        {
            var cols = guard.Track(range.Columns);
            cols.AutoFit();
        }

        return Response.Ok(new { @ref = refStr, autofit = target });
    }

    private string HandleMerge(JsonObject args)
    {
        var refStr = args["ref"]?.GetValue<string>()
            ?? throw new ArgumentException("merge requires 'ref'");

        using var guard = new ComGuard();
        var sheet = guard.Track(_session.GetActiveSheet());
        var range = guard.Track(sheet.Range[refStr]);
        range.Merge();

        return Response.Ok(new { @ref = refStr, merged = true });
    }

    private string HandleUnmerge(JsonObject args)
    {
        var refStr = args["ref"]?.GetValue<string>()
            ?? throw new ArgumentException("unmerge requires 'ref'");

        using var guard = new ComGuard();
        var sheet = guard.Track(_session.GetActiveSheet());
        var range = guard.Track(sheet.Range[refStr]);
        range.UnMerge();

        return Response.Ok(new { @ref = refStr, unmerged = true });
    }

    private string HandleFreeze(JsonObject args)
    {
        var refStr = args["ref"]?.GetValue<string>()
            ?? throw new ArgumentException("freeze requires 'ref'");

        using var guard = new ComGuard();
        var sheet = guard.Track(_session.GetActiveSheet());
        var range = guard.Track(sheet.Range[refStr]);
        range.Select();

        var window = guard.Track(_session.App.ActiveWindow);
        window.FreezePanes = true;

        return Response.Ok(new { @ref = refStr, frozen = true });
    }

    private string HandleUnfreeze(JsonObject args)
    {
        using var guard = new ComGuard();
        var window = guard.Track(_session.App.ActiveWindow);
        window.FreezePanes = false;

        return Response.Ok(new { unfrozen = true });
    }

    private string HandleHide(JsonObject args)
    {
        var target = args["target"]?.GetValue<string>()
            ?? throw new ArgumentException("hide requires 'target' (row/col)");
        var refStr = args["ref"]?.GetValue<string>()
            ?? throw new ArgumentException("hide requires 'ref'");

        using var guard = new ComGuard();
        var sheet = guard.Track(_session.GetActiveSheet());
        var range = guard.Track(sheet.Range[refStr]);

        if (target == "row" || target == "rows")
        {
            var rows = guard.Track(range.EntireRow);
            rows.Hidden = true;
        }
        else
        {
            var cols = guard.Track(range.EntireColumn);
            cols.Hidden = true;
        }

        return Response.Ok(new { @ref = refStr, target, hidden = true });
    }

    private string HandleUnhide(JsonObject args)
    {
        var target = args["target"]?.GetValue<string>()
            ?? throw new ArgumentException("unhide requires 'target' (row/col)");
        var refStr = args["ref"]?.GetValue<string>()
            ?? throw new ArgumentException("unhide requires 'ref'");

        using var guard = new ComGuard();
        var sheet = guard.Track(_session.GetActiveSheet());
        var range = guard.Track(sheet.Range[refStr]);

        if (target == "row" || target == "rows")
        {
            var rows = guard.Track(range.EntireRow);
            rows.Hidden = false;
        }
        else
        {
            var cols = guard.Track(range.EntireColumn);
            cols.Hidden = false;
        }

        return Response.Ok(new { @ref = refStr, target, unhidden = true });
    }

    private string HandleInsertRow(JsonObject args)
    {
        var refStr = args["ref"]?.GetValue<string>()
            ?? throw new ArgumentException("insert.row requires 'ref'");

        using var guard = new ComGuard();
        var sheet = guard.Track(_session.GetActiveSheet());
        var range = guard.Track(sheet.Range[refStr]);
        var row = guard.Track(range.EntireRow);
        row.Insert(Excel.XlInsertShiftDirection.xlShiftDown);

        return Response.Ok(new { @ref = refStr, inserted = "row" });
    }

    private string HandleInsertCol(JsonObject args)
    {
        var refStr = args["ref"]?.GetValue<string>()
            ?? throw new ArgumentException("insert.col requires 'ref'");

        using var guard = new ComGuard();
        var sheet = guard.Track(_session.GetActiveSheet());
        var range = guard.Track(sheet.Range[refStr]);
        var col = guard.Track(range.EntireColumn);
        col.Insert(Excel.XlInsertShiftDirection.xlShiftToRight);

        return Response.Ok(new { @ref = refStr, inserted = "column" });
    }

    private string HandleDeleteRow(JsonObject args)
    {
        var refStr = args["ref"]?.GetValue<string>()
            ?? throw new ArgumentException("delete.row requires 'ref'");

        using var guard = new ComGuard();
        var sheet = guard.Track(_session.GetActiveSheet());
        var range = guard.Track(sheet.Range[refStr]);
        var row = guard.Track(range.EntireRow);
        row.Delete(Excel.XlDeleteShiftDirection.xlShiftUp);

        return Response.Ok(new { @ref = refStr, deleted = "row" });
    }

    private string HandleDeleteCol(JsonObject args)
    {
        var refStr = args["ref"]?.GetValue<string>()
            ?? throw new ArgumentException("delete.col requires 'ref'");

        using var guard = new ComGuard();
        var sheet = guard.Track(_session.GetActiveSheet());
        var range = guard.Track(sheet.Range[refStr]);
        var col = guard.Track(range.EntireColumn);
        col.Delete(Excel.XlDeleteShiftDirection.xlShiftToLeft);

        return Response.Ok(new { @ref = refStr, deleted = "column" });
    }

    private string HandleColumnCount(JsonObject args)
    {
        using var guard = new ComGuard();
        var sheet = guard.Track(_session.GetActiveSheet());
        var used = guard.Track(sheet.UsedRange);
        return Response.Ok(new { columns = used.Columns.Count });
    }

    private string HandleRowCount(JsonObject args)
    {
        using var guard = new ComGuard();
        var sheet = guard.Track(_session.GetActiveSheet());
        var used = guard.Track(sheet.UsedRange);
        return Response.Ok(new { rows = used.Rows.Count });
    }

    private string HandleUsedRange(JsonObject args)
    {
        using var guard = new ComGuard();
        var sheet = guard.Track(_session.GetActiveSheet());
        var used = guard.Track(sheet.UsedRange);
        return Response.Ok(new
        {
            @ref = used.Address[false, false],
            rows = used.Rows.Count,
            columns = used.Columns.Count,
        });
    }
}
