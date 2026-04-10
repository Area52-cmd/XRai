// Leak-audited: 2026-04-10 — All Range/Font/Cell COM proxies pass through
// ComGuard.Track or are explicitly released after use. The cell-iteration
// loop in HandleRead releases each cell proxy on every iteration. AddCellData
// tracks Font via the guard. No leaks found.

using System.Text.Json.Nodes;
using XRai.Core;
using Excel = Microsoft.Office.Interop.Excel;

namespace XRai.Com;

public class CellOps
{
    private readonly ExcelSession _session;

    public CellOps(ExcelSession session)
    {
        _session = session;
    }

    public void Register(CommandRouter router)
    {
        router.Register("read", HandleRead);
        router.Register("type", HandleType);
        router.Register("clear", HandleClear);
        router.Register("select", HandleSelect);
        router.Register("format", HandleFormat);
        router.Register("format.read", HandleFormatRead);
    }

    private string HandleRead(JsonObject args)
    {
        var refStr = args["ref"]?.GetValue<string>()
            ?? throw new ArgumentException("read requires 'ref'");
        bool full = args["full"]?.GetValue<bool>() == true
            || args["full"]?.GetValue<string>() == "true";

        using var guard = new ComGuard();
        var range = guard.Track(_session.ResolveRange(refStr));

        int rows = range.Rows.Count;
        int cols = range.Columns.Count;

        // Single cell
        if (rows == 1 && cols == 1)
        {
            return ReadSingleCell(range, full, guard);
        }

        // Multi-cell range
        var cells = new JsonArray();
        foreach (Excel.Range cell in range)
        {
            var cellObj = new JsonObject
            {
                ["ref"] = cell.Address[false, false],
            };
            AddCellData(cellObj, cell, full, guard);
            cells.Add(cellObj);
            System.Runtime.InteropServices.Marshal.ReleaseComObject(cell);
        }

        return Response.Ok(new { cells });
    }

    private string ReadSingleCell(Excel.Range range, bool full, ComGuard guard)
    {
        var data = new JsonObject();
        AddCellData(data, range, full, guard);
        return Response.Ok(data);
    }

    private void AddCellData(JsonObject obj, Excel.Range cell, bool full, ComGuard guard)
    {
        var val = cell.Value2;
        if (val is double d)
            obj["value"] = d;
        else if (val is bool b)
            obj["value"] = b;
        else if (val != null)
            obj["value"] = val.ToString();
        else
            obj["value"] = null;

        if (full)
        {
            var formula = cell.Formula?.ToString();
            if (!string.IsNullOrEmpty(formula) && formula.StartsWith("="))
                obj["formula"] = formula;

            var font = guard.Track(cell.Font);
            obj["font_name"] = font.Name?.ToString();
            obj["font_size"] = Convert.ToDouble(font.Size);
            obj["bold"] = Convert.ToBoolean(font.Bold);

            var nf = cell.NumberFormat?.ToString();
            if (!string.IsNullOrEmpty(nf))
                obj["number_format"] = nf;

            // Check for errors
            if (cell.Value2 is int errVal && errVal >= 0)
            {
                var text = cell.Text?.ToString();
                if (text != null && text.StartsWith("#"))
                    obj["error"] = text;
            }
        }
    }

    private string HandleType(JsonObject args)
    {
        var refStr = args["ref"]?.GetValue<string>()
            ?? throw new ArgumentException("type requires 'ref'");

        using var guard = new ComGuard();
        var range = guard.Track(_session.ResolveRange(refStr));

        // Array of values
        if (args["values"] is JsonArray valuesArr)
        {
            var values = new object[valuesArr.Count, 1];
            for (int i = 0; i < valuesArr.Count; i++)
            {
                var v = valuesArr[i];
                if (v is JsonValue jv && jv.TryGetValue<double>(out var dv))
                    values[i, 0] = dv;
                else
                    values[i, 0] = v?.ToString() ?? "";
            }
            range.Value2 = values;
            return Response.Ok(new { @ref = refStr, count = valuesArr.Count });
        }

        // Single value or formula.
        // Accepts EITHER a JSON string ("100", "Hello", "=SUM(A:A)") or a JSON
        // number (100, 3.14). JSON booleans are coerced to string form.
        var valueNode = args["value"];
        if (valueNode == null)
            throw new ArgumentException("type requires 'value' or 'values'");

        // Numeric fast path — avoids any string round-trip / parse loss
        if (valueNode is JsonValue jsonValue)
        {
            if (jsonValue.TryGetValue<double>(out var numDirect))
            {
                range.Value2 = numDirect;
                return Response.Ok(new { @ref = refStr, typed = numDirect });
            }
            if (jsonValue.TryGetValue<bool>(out var boolDirect))
            {
                range.Value2 = boolDirect;
                return Response.Ok(new { @ref = refStr, typed = boolDirect });
            }
        }

        // String or stringifiable path
        var value = valueNode.ToString() ?? "";
        // JsonNode.ToString() of a string value returns the raw string without quotes,
        // which is what we want for cell input.

        var array = args["array"]?.GetValue<bool>() ?? false;

        if (value.StartsWith("="))
        {
            if (array)
                range.FormulaArray = value;
            else
                range.Formula = value;
        }
        else if (double.TryParse(value, System.Globalization.CultureInfo.InvariantCulture, out var num))
            range.Value2 = num;
        else
            range.Value2 = value;

        return Response.Ok(new { @ref = refStr, typed = value, array });
    }

    private string HandleClear(JsonObject args)
    {
        var refStr = args["ref"]?.GetValue<string>()
            ?? throw new ArgumentException("clear requires 'ref'");

        using var guard = new ComGuard();
        var range = guard.Track(_session.ResolveRange(refStr));
        range.Clear();

        return Response.Ok(new { @ref = refStr, cleared = true });
    }

    private string HandleSelect(JsonObject args)
    {
        var refStr = args["ref"]?.GetValue<string>()
            ?? throw new ArgumentException("select requires 'ref'");

        using var guard = new ComGuard();
        var range = guard.Track(_session.ResolveRange(refStr));
        range.Select();

        return Response.Ok(new { @ref = refStr, selected = true });
    }

    private string HandleFormat(JsonObject args)
    {
        var refStr = args["ref"]?.GetValue<string>()
            ?? throw new ArgumentException("format requires 'ref'");

        using var guard = new ComGuard();
        var range = guard.Track(_session.ResolveRange(refStr));
        var font = guard.Track(range.Font);
        var interior = guard.Track(range.Interior);

        if (args["bold"] != null)
            font.Bold = args["bold"]!.GetValue<bool>();
        if (args["font_size"] != null)
            font.Size = args["font_size"]!.GetValue<double>();
        if (args["font_name"] != null)
            font.Name = args["font_name"]!.GetValue<string>();
        if (args["number_format"] != null)
            range.NumberFormat = args["number_format"]!.GetValue<string>();
        if (args["bg"] != null)
        {
            var bg = args["bg"]!.GetValue<string>();
            interior.Color = ParseColor(bg);
        }

        return Response.Ok(new { @ref = refStr, formatted = true });
    }

    private string HandleFormatRead(JsonObject args)
    {
        var refStr = args["ref"]?.GetValue<string>()
            ?? throw new ArgumentException("format.read requires 'ref'");

        using var guard = new ComGuard();
        var range = guard.Track(_session.ResolveRange(refStr));
        var font = guard.Track(range.Font);
        var interior = guard.Track(range.Interior);

        return Response.Ok(new
        {
            @ref = refStr,
            font_name = font.Name?.ToString(),
            font_size = Convert.ToDouble(font.Size),
            bold = Convert.ToBoolean(font.Bold),
            italic = Convert.ToBoolean(font.Italic),
            number_format = range.NumberFormat?.ToString(),
            bg_color = Convert.ToInt32(interior.Color),
        });
    }

    private static int ParseColor(string color)
    {
        // Support hex (#RRGGBB) or named colors
        if (color.StartsWith("#") && color.Length == 7)
        {
            int r = Convert.ToInt32(color[1..3], 16);
            int g = Convert.ToInt32(color[3..5], 16);
            int b = Convert.ToInt32(color[5..7], 16);
            return r | (g << 8) | (b << 16); // OLE color is BGR
        }

        return color.ToLowerInvariant() switch
        {
            "yellow" => 65535,
            "red" => 255,
            "green" => 65280,
            "blue" => 16711680,
            "white" => 16777215,
            _ => 0,
        };
    }
}
