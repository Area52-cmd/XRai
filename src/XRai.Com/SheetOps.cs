using System.Runtime.InteropServices;
using System.Text.Json.Nodes;
using XRai.Core;
using Excel = Microsoft.Office.Interop.Excel;

namespace XRai.Com;

public class SheetOps
{
    private readonly ExcelSession _session;

    public SheetOps(ExcelSession session)
    {
        _session = session;
    }

    public void Register(CommandRouter router)
    {
        router.Register("sheets", HandleSheets);
        router.Register("sheet.add", HandleSheetAdd);
        router.Register("sheet.rename", HandleSheetRename);
        router.Register("sheet.delete", HandleSheetDelete);
        router.Register("goto", HandleGoto);
        router.Register("names", HandleNames);
        router.Register("name.read", HandleNameRead);
        router.Register("name.set", HandleNameSet);
    }

    private string HandleSheets(JsonObject args)
    {
        using var guard = new ComGuard();
        var workbook = guard.Track(_session.GetActiveWorkbook());
        var sheets = guard.Track(workbook.Sheets);

        var result = new JsonArray();
        foreach (Excel.Worksheet sheet in sheets)
        {
            result.Add(new JsonObject
            {
                ["name"] = sheet.Name,
                ["index"] = sheet.Index,
                ["visible"] = sheet.Visible == Excel.XlSheetVisibility.xlSheetVisible,
            });
            Marshal.ReleaseComObject(sheet);
        }

        var active = guard.Track(_session.GetActiveSheet());
        return Response.Ok(new { sheets = result, active = active.Name });
    }

    private string HandleSheetAdd(JsonObject args)
    {
        var name = args["name"]?.GetValue<string>();
        var after = args["after"]?.GetValue<string>();

        using var guard = new ComGuard();
        var workbook = guard.Track(_session.GetActiveWorkbook());
        var sheets = guard.Track(workbook.Sheets);

        Excel.Worksheet newSheet;
        if (after != null)
        {
            // Find the sheet to add after
            Excel.Worksheet afterSheet = null!;
            foreach (Excel.Worksheet s in sheets)
            {
                if (string.Equals(s.Name, after, StringComparison.OrdinalIgnoreCase))
                {
                    afterSheet = s;
                    break;
                }
                Marshal.ReleaseComObject(s);
            }
            newSheet = (Excel.Worksheet)sheets.Add(After: afterSheet);
            Marshal.ReleaseComObject(afterSheet);
        }
        else
        {
            newSheet = (Excel.Worksheet)sheets.Add();
        }

        if (name != null)
            newSheet.Name = name;

        var resultName = newSheet.Name;
        Marshal.ReleaseComObject(newSheet);

        return Response.Ok(new { name = resultName, added = true });
    }

    private string HandleSheetRename(JsonObject args)
    {
        var from = args["from"]?.GetValue<string>()
            ?? throw new ArgumentException("sheet.rename requires 'from'");
        var to = args["to"]?.GetValue<string>()
            ?? throw new ArgumentException("sheet.rename requires 'to'");

        using var guard = new ComGuard();
        var workbook = guard.Track(_session.GetActiveWorkbook());
        var sheets = guard.Track(workbook.Sheets);

        foreach (Excel.Worksheet sheet in sheets)
        {
            if (string.Equals(sheet.Name, from, StringComparison.OrdinalIgnoreCase))
            {
                sheet.Name = to;
                Marshal.ReleaseComObject(sheet);
                return Response.Ok(new { from, to, renamed = true });
            }
            Marshal.ReleaseComObject(sheet);
        }

        return Response.Error($"Sheet not found: {from}", code: ErrorCodes.SheetNotFound);
    }

    private string HandleSheetDelete(JsonObject args)
    {
        var name = args["name"]?.GetValue<string>()
            ?? throw new ArgumentException("sheet.delete requires 'name'");

        using var guard = new ComGuard();
        var workbook = guard.Track(_session.GetActiveWorkbook());
        var sheets = guard.Track(workbook.Sheets);

        // Suppress the confirmation dialog
        _session.App.DisplayAlerts = false;
        try
        {
            foreach (Excel.Worksheet sheet in sheets)
            {
                if (string.Equals(sheet.Name, name, StringComparison.OrdinalIgnoreCase))
                {
                    sheet.Delete();
                    Marshal.ReleaseComObject(sheet);
                    return Response.Ok(new { name, deleted = true });
                }
                Marshal.ReleaseComObject(sheet);
            }
        }
        finally
        {
            _session.App.DisplayAlerts = true;
        }

        return Response.Error($"Sheet not found: {name}", code: ErrorCodes.SheetNotFound);
    }

    private string HandleGoto(JsonObject args)
    {
        var target = args["target"]?.GetValue<string>()
            ?? args["sheet"]?.GetValue<string>()
            ?? throw new ArgumentException("goto requires 'target' or 'sheet'");

        using var guard = new ComGuard();

        // Try as a range reference first (handles "A1", "A1:D10", "Sheet1!A1:D10")
        // This must come BEFORE the sheet-name lookup so sheet-qualified refs
        // like "Sheet1!A1" don't get misinterpreted as a sheet named "Sheet1!A1".
        if (target.Contains('!') || target.Contains(':') ||
            (target.Length >= 2 && char.IsLetter(target[0]) && char.IsDigit(target[^1])))
        {
            try
            {
                var range = guard.Track(_session.ResolveRange(target));
                range.Select();
                return Response.Ok(new { @goto = target, @ref = range.Address[false, false], selected = true });
            }
            catch { /* fall through to sheet/named-range lookup */ }
        }

        var workbook = guard.Track(_session.GetActiveWorkbook());

        // Try as sheet name
        var sheets = guard.Track(workbook.Sheets);
        foreach (Excel.Worksheet sheet in sheets)
        {
            if (string.Equals(sheet.Name, target, StringComparison.OrdinalIgnoreCase))
            {
                sheet.Activate();
                var name = sheet.Name;
                Marshal.ReleaseComObject(sheet);
                return Response.Ok(new { activated = name });
            }
            Marshal.ReleaseComObject(sheet);
        }

        // Try as named range
        try
        {
            var names = guard.Track(workbook.Names);
            var named = guard.Track(names.Item(target));
            var refRange = guard.Track(named.RefersToRange);
            refRange.Select();
            return Response.Ok(new { @goto = target, @ref = refRange.Address[false, false] });
        }
        catch
        {
            return Response.Error(
                $"Sheet, named range, or cell reference not found: {target}. " +
                "Valid forms: \"Sheet1\" (sheet name), \"MyRange\" (named range), " +
                "\"A1:D10\" (cells on active sheet), \"Sheet1!A1:D10\" (cells on specific sheet).",
                code: ErrorCodes.InvalidRange);
        }
    }

    private string HandleNames(JsonObject args)
    {
        using var guard = new ComGuard();
        var workbook = guard.Track(_session.GetActiveWorkbook());
        var names = guard.Track(workbook.Names);

        var result = new JsonArray();
        foreach (Excel.Name name in names)
        {
            result.Add(new JsonObject
            {
                ["name"] = name.Name,
                ["refers_to"] = name.RefersTo?.ToString(),
                ["visible"] = name.Visible,
            });
            Marshal.ReleaseComObject(name);
        }

        return Response.Ok(new { names = result });
    }

    private string HandleNameRead(JsonObject args)
    {
        var name = args["name"]?.GetValue<string>()
            ?? throw new ArgumentException("name.read requires 'name'");

        using var guard = new ComGuard();
        var workbook = guard.Track(_session.GetActiveWorkbook());
        var names = guard.Track(workbook.Names);
        var named = guard.Track(names.Item(name));

        return Response.Ok(new
        {
            name,
            refers_to = named.RefersTo?.ToString(),
            @ref = named.RefersToRange?.Address[false, false],
        });
    }

    private string HandleNameSet(JsonObject args)
    {
        var name = args["name"]?.GetValue<string>()
            ?? throw new ArgumentException("name.set requires 'name'");
        var refStr = args["ref"]?.GetValue<string>()
            ?? throw new ArgumentException("name.set requires 'ref'");

        using var guard = new ComGuard();
        var workbook = guard.Track(_session.GetActiveWorkbook());
        var names = guard.Track(workbook.Names);

        var sheet = guard.Track(_session.GetActiveSheet());
        var range = guard.Track(sheet.Range[refStr]);
        names.Add(name, range);

        return Response.Ok(new { name, @ref = refStr, defined = true });
    }
}
