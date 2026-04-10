using System.Runtime.InteropServices;
using System.Text.Json.Nodes;
using XRai.Core;
using Excel = Microsoft.Office.Interop.Excel;

namespace XRai.Com;

public class AdvancedOps
{
    private readonly ExcelSession _session;
    public AdvancedOps(ExcelSession session) { _session = session; }

    public void Register(CommandRouter router)
    {
        router.Register("undo", HandleUndo);
        router.Register("redo", HandleRedo);
        router.Register("find.all", HandleFindAll);
        router.Register("paste.special", HandlePasteSpecial);
        router.Register("fill.series", HandleFillSeries);
        router.Register("group", HandleGroup);
        router.Register("ungroup", HandleUngroup);
        router.Register("outline.level", HandleOutlineLevel);
        router.Register("name.delete", HandleNameDelete);
        router.Register("macro.run", HandleMacroRun);
        router.Register("selection.info", HandleSelectionInfo);
        router.Register("error.check", HandleErrorCheck);
        router.Register("link.list", HandleLinkList);
        router.Register("link.update", HandleLinkUpdate);
        router.Register("time.udf", HandleTimeUdf);
    }

    private string HandleUndo(JsonObject args)
    {
        _session.App.Undo();
        return Response.Ok(new { undone = true });
    }

    private string HandleRedo(JsonObject args)
    {
        _session.App.Repeat();
        return Response.Ok(new { redone = true });
    }

    private string HandleFindAll(JsonObject args)
    {
        var what = args["what"]?.GetValue<string>() ?? throw new ArgumentException("find.all requires 'what'");
        var refStr = args["ref"]?.GetValue<string>();

        using var guard = new ComGuard();
        var sheet = guard.Track(_session.GetActiveSheet());
        var searchRange = refStr != null ? guard.Track(sheet.Range[refStr]) : guard.Track(sheet.UsedRange);

        var results = new JsonArray();
        var found = searchRange.Find(what);
        if (found != null)
        {
            var firstAddr = found.Address[false, false];
            do
            {
                results.Add(new JsonObject
                {
                    ["ref"] = found.Address[false, false],
                    ["value"] = found.Value2?.ToString(),
                });
                var next = searchRange.FindNext(found);
                Marshal.ReleaseComObject(found);
                found = next;
            } while (found != null && found.Address[false, false] != firstAddr);

            if (found != null) Marshal.ReleaseComObject(found);
        }

        return Response.Ok(new { what, matches = results, count = results.Count });
    }

    private string HandlePasteSpecial(JsonObject args)
    {
        var refStr = args["ref"]?.GetValue<string>() ?? throw new ArgumentException("paste.special requires 'ref'");
        var type = args["type"]?.GetValue<string>() ?? "all";
        var operation = args["operation"]?.GetValue<string>() ?? "none";
        var skipBlanks = args["skip_blanks"]?.GetValue<bool>() ?? false;
        var transpose = args["transpose"]?.GetValue<bool>() ?? false;

        var xlPaste = type.ToLowerInvariant() switch
        {
            "values" => Excel.XlPasteType.xlPasteValues,
            "formats" => Excel.XlPasteType.xlPasteFormats,
            "formulas" => Excel.XlPasteType.xlPasteFormulas,
            "comments" => Excel.XlPasteType.xlPasteComments,
            "validation" => Excel.XlPasteType.xlPasteValidation,
            "column_widths" => Excel.XlPasteType.xlPasteColumnWidths,
            _ => Excel.XlPasteType.xlPasteAll,
        };

        var xlOp = operation.ToLowerInvariant() switch
        {
            "add" => Excel.XlPasteSpecialOperation.xlPasteSpecialOperationAdd,
            "subtract" => Excel.XlPasteSpecialOperation.xlPasteSpecialOperationSubtract,
            "multiply" => Excel.XlPasteSpecialOperation.xlPasteSpecialOperationMultiply,
            "divide" => Excel.XlPasteSpecialOperation.xlPasteSpecialOperationDivide,
            _ => Excel.XlPasteSpecialOperation.xlPasteSpecialOperationNone,
        };

        using var guard = new ComGuard();
        var sheet = guard.Track(_session.GetActiveSheet());
        var range = guard.Track(sheet.Range[refStr]);
        range.PasteSpecial(xlPaste, xlOp, skipBlanks, transpose);

        return Response.Ok(new { @ref = refStr, type, operation, pasted = true });
    }

    private string HandleFillSeries(JsonObject args)
    {
        var refStr = args["ref"]?.GetValue<string>() ?? throw new ArgumentException("fill.series requires 'ref'");
        var type = args["type"]?.GetValue<string>() ?? "linear";
        var step = args["step"]?.GetValue<double>() ?? 1;
        var stop = args["stop"]?.GetValue<double>();

        var xlType = type.ToLowerInvariant() switch
        {
            "growth" => Excel.XlAutoFillType.xlFillValues,
            "date" => Excel.XlAutoFillType.xlFillDays,
            _ => Excel.XlAutoFillType.xlFillSeries,
        };

        using var guard = new ComGuard();
        var sheet = guard.Track(_session.GetActiveSheet());
        var range = guard.Track(sheet.Range[refStr]);

        // Use DataSeries for more control
        range.DataSeries(
            Type: type.ToLowerInvariant() switch
            {
                "growth" => Excel.XlDataSeriesType.xlDataSeriesLinear,
                "date" => Excel.XlDataSeriesType.xlChronological,
                _ => Excel.XlDataSeriesType.xlDataSeriesLinear,
            },
            Step: step,
            Stop: stop ?? Type.Missing
        );

        return Response.Ok(new { @ref = refStr, type, step, filled = true });
    }

    private string HandleGroup(JsonObject args)
    {
        var refStr = args["ref"]?.GetValue<string>() ?? throw new ArgumentException("group requires 'ref'");
        var target = args["target"]?.GetValue<string>() ?? "rows";

        using var guard = new ComGuard();
        var sheet = guard.Track(_session.GetActiveSheet());
        var range = guard.Track(sheet.Range[refStr]);

        if (target == "rows")
            range.EntireRow.Group();
        else
            range.EntireColumn.Group();

        return Response.Ok(new { @ref = refStr, target, grouped = true });
    }

    private string HandleUngroup(JsonObject args)
    {
        var refStr = args["ref"]?.GetValue<string>() ?? throw new ArgumentException("ungroup requires 'ref'");
        var target = args["target"]?.GetValue<string>() ?? "rows";

        using var guard = new ComGuard();
        var sheet = guard.Track(_session.GetActiveSheet());
        var range = guard.Track(sheet.Range[refStr]);

        if (target == "rows")
            range.EntireRow.Ungroup();
        else
            range.EntireColumn.Ungroup();

        return Response.Ok(new { @ref = refStr, target, ungrouped = true });
    }

    private string HandleOutlineLevel(JsonObject args)
    {
        var level = args["level"]?.GetValue<int>() ?? 1;
        var target = args["target"]?.GetValue<string>() ?? "rows";

        using var guard = new ComGuard();
        var sheet = guard.Track(_session.GetActiveSheet());
        var outline = sheet.Outline;

        if (target == "rows")
            outline.ShowLevels(level);
        else
            outline.ShowLevels(ColumnLevels: level);

        return Response.Ok(new { level, target });
    }

    private string HandleNameDelete(JsonObject args)
    {
        var name = args["name"]?.GetValue<string>() ?? throw new ArgumentException("name.delete requires 'name'");

        using var guard = new ComGuard();
        var wb = guard.Track(_session.GetActiveWorkbook());
        var names = guard.Track(wb.Names);
        var named = guard.Track(names.Item(name));
        named.Delete();

        return Response.Ok(new { name, deleted = true });
    }

    private string HandleMacroRun(JsonObject args)
    {
        var name = args["name"]?.GetValue<string>() ?? throw new ArgumentException("macro.run requires 'name'");
        var macroArgs = args["args"]?.AsArray();

        object? result;
        if (macroArgs != null && macroArgs.Count > 0)
        {
            var argValues = macroArgs.Select(a => (object)(a?.ToString() ?? "")).ToArray();
            result = argValues.Length switch
            {
                1 => _session.App.Run(name, argValues[0]),
                2 => _session.App.Run(name, argValues[0], argValues[1]),
                3 => _session.App.Run(name, argValues[0], argValues[1], argValues[2]),
                _ => _session.App.Run(name),
            };
        }
        else
        {
            result = _session.App.Run(name);
        }

        return Response.Ok(new { name, result = result?.ToString(), executed = true });
    }

    private string HandleSelectionInfo(JsonObject args)
    {
        using var guard = new ComGuard();
        var activeCell = guard.Track(_session.App.ActiveCell);
        var selection = _session.App.Selection;

        string selAddr = "";
        try
        {
            if (selection is Excel.Range selRange)
            {
                selAddr = selRange.Address[false, false];
                Marshal.ReleaseComObject(selRange);
            }
        }
        catch { }

        return Response.Ok(new
        {
            active_cell = activeCell.Address[false, false],
            selection = selAddr,
            sheet = _session.GetActiveSheet().Name,
        });
    }

    private string HandleErrorCheck(JsonObject args)
    {
        var refStr = args["ref"]?.GetValue<string>() ?? throw new ArgumentException("error.check requires 'ref'");

        using var guard = new ComGuard();
        var sheet = guard.Track(_session.GetActiveSheet());
        var range = guard.Track(sheet.Range[refStr]);

        var val = range.Value2;
        var text = range.Text?.ToString() ?? "";
        bool isError = text.StartsWith("#");

        return Response.Ok(new
        {
            @ref = refStr,
            is_error = isError,
            error_text = isError ? text : null,
            value = val?.ToString(),
        });
    }

    private string HandleLinkList(JsonObject args)
    {
        using var guard = new ComGuard();
        var wb = guard.Track(_session.GetActiveWorkbook());
        var links = wb.LinkSources(Excel.XlLink.xlExcelLinks);

        var result = new JsonArray();
        if (links is Array arr)
        {
            foreach (var link in arr)
                result.Add(JsonValue.Create(link?.ToString()));
        }

        return Response.Ok(new { links = result, count = result.Count });
    }

    private string HandleLinkUpdate(JsonObject args)
    {
        using var guard = new ComGuard();
        var wb = guard.Track(_session.GetActiveWorkbook());
        wb.UpdateLink(Type: Excel.XlLinkType.xlLinkTypeExcelLinks);
        return Response.Ok(new { updated = true });
    }

    private string HandleTimeUdf(JsonObject args)
    {
        var function = args["function"]?.GetValue<string>() ?? throw new ArgumentException("time.udf requires 'function'");
        var udfArgs = args["args"]?.GetValue<string>() ?? "";
        var iterations = args["n"]?.GetValue<int>() ?? 100;

        using var guard = new ComGuard();
        var sheet = guard.Track(_session.GetActiveSheet());

        // Write UDF formula to a temp cell, time recalc
        var tempCell = guard.Track(sheet.Range["XFD1"]);
        var formula = string.IsNullOrEmpty(udfArgs)
            ? $"={function}()"
            : $"={function}({udfArgs})";

        tempCell.Formula = formula;

        var sw = System.Diagnostics.Stopwatch.StartNew();
        for (int i = 0; i < iterations; i++)
        {
            tempCell.Calculate();
        }
        sw.Stop();

        var result = tempCell.Value2;
        tempCell.Clear();

        return Response.Ok(new
        {
            function,
            iterations,
            total_ms = sw.ElapsedMilliseconds,
            avg_ms = (double)sw.ElapsedMilliseconds / iterations,
            result = result?.ToString(),
        });
    }
}
