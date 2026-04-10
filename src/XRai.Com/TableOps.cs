using System.Runtime.InteropServices;
using System.Text.Json.Nodes;
using XRai.Core;
using Excel = Microsoft.Office.Interop.Excel;

namespace XRai.Com;

public class TableOps
{
    private readonly ExcelSession _session;

    public TableOps(ExcelSession session) { _session = session; }

    public void Register(CommandRouter router)
    {
        router.Register("table.list", HandleList);
        router.Register("table.create", HandleCreate);
        router.Register("table.delete", HandleDelete);
        router.Register("table.style", HandleStyle);
        router.Register("table.resize", HandleResize);
        router.Register("table.totals", HandleTotals);
        router.Register("table.filter", HandleFilter);
        router.Register("table.filter.clear", HandleFilterClear);
        router.Register("table.sort", HandleSort);
        router.Register("table.row.add", HandleRowAdd);
        router.Register("table.column.add", HandleColumnAdd);
        router.Register("table.data", HandleData);
    }

    private string HandleList(JsonObject args)
    {
        try
        {
            using var guard = new ComGuard();
            var sheet = guard.Track(_session.GetActiveSheet());
            var tables = guard.Track(sheet.ListObjects);

            var result = new JsonArray();
            foreach (Excel.ListObject tbl in tables)
            {
                var range = guard.Track(tbl.Range);
                result.Add(new JsonObject
                {
                    ["name"] = tbl.Name,
                    ["range"] = range.Address[false, false],
                    ["style"] = tbl.TableStyle?.ToString() ?? "",
                    ["rows"] = tbl.ListRows.Count,
                    ["columns"] = tbl.ListColumns.Count,
                });
                Marshal.ReleaseComObject(tbl);
            }

            return Response.Ok(new { tables = result, count = result.Count });
        }
        catch (Exception ex)
        {
            return Response.Error($"table.list: {ex.Message}");
        }
    }

    private string HandleCreate(JsonObject args)
    {
        try
        {
            var refStr = args["ref"]?.GetValue<string>()
                ?? throw new ArgumentException("table.create requires 'ref'");
            var name = args["name"]?.GetValue<string>();
            var style = args["style"]?.GetValue<string>();

            using var guard = new ComGuard();
            var sheet = guard.Track(_session.GetActiveSheet());
            var range = guard.Track(sheet.Range[refStr]);
            var tables = guard.Track(sheet.ListObjects);

            var table = guard.Track(tables.Add(
                Excel.XlListObjectSourceType.xlSrcRange,
                range,
                XlListObjectHasHeaders: Excel.XlYesNoGuess.xlYes));

            if (name != null)
                table.Name = name;

            if (style != null)
                table.TableStyle = style;

            return Response.Ok(new
            {
                name = table.Name,
                @ref = range.Address[false, false],
                created = true,
            });
        }
        catch (Exception ex)
        {
            return Response.Error($"table.create: {ex.Message}");
        }
    }

    private string HandleDelete(JsonObject args)
    {
        try
        {
            var name = args["name"]?.GetValue<string>()
                ?? throw new ArgumentException("table.delete requires 'name'");

            using var guard = new ComGuard();
            var sheet = guard.Track(_session.GetActiveSheet());
            var table = FindTable(sheet, name, guard);
            table.Unlist(); // Convert back to range without deleting data

            return Response.Ok(new { name, deleted = true });
        }
        catch (Exception ex)
        {
            return Response.Error($"table.delete: {ex.Message}");
        }
    }

    private string HandleStyle(JsonObject args)
    {
        try
        {
            var name = args["name"]?.GetValue<string>()
                ?? throw new ArgumentException("table.style requires 'name'");
            var style = args["style"]?.GetValue<string>()
                ?? throw new ArgumentException("table.style requires 'style'");

            using var guard = new ComGuard();
            var sheet = guard.Track(_session.GetActiveSheet());
            var table = FindTable(sheet, name, guard);
            table.TableStyle = style;

            return Response.Ok(new { name, style, applied = true });
        }
        catch (Exception ex)
        {
            return Response.Error($"table.style: {ex.Message}");
        }
    }

    private string HandleResize(JsonObject args)
    {
        try
        {
            var name = args["name"]?.GetValue<string>()
                ?? throw new ArgumentException("table.resize requires 'name'");
            var refStr = args["ref"]?.GetValue<string>()
                ?? throw new ArgumentException("table.resize requires 'ref'");

            using var guard = new ComGuard();
            var sheet = guard.Track(_session.GetActiveSheet());
            var table = FindTable(sheet, name, guard);
            var newRange = guard.Track(sheet.Range[refStr]);
            table.Resize(newRange);

            return Response.Ok(new { name, @ref = refStr, resized = true });
        }
        catch (Exception ex)
        {
            return Response.Error($"table.resize: {ex.Message}");
        }
    }

    private string HandleTotals(JsonObject args)
    {
        try
        {
            var name = args["name"]?.GetValue<string>()
                ?? throw new ArgumentException("table.totals requires 'name'");

            using var guard = new ComGuard();
            var sheet = guard.Track(_session.GetActiveSheet());
            var table = FindTable(sheet, name, guard);

            if (args["show"] is JsonNode showNode)
                table.ShowTotals = showNode.GetValue<bool>();

            if (args["column"] is JsonNode colNode && args["function"] is JsonNode funcNode)
            {
                table.ShowTotals = true;
                var colName = colNode.GetValue<string>();
                var columns = guard.Track(table.ListColumns);
                var column = guard.Track(columns[colName]);

                column.TotalsCalculation = funcNode.GetValue<string>().ToLowerInvariant() switch
                {
                    "sum" => Excel.XlTotalsCalculation.xlTotalsCalculationSum,
                    "average" or "avg" => Excel.XlTotalsCalculation.xlTotalsCalculationAverage,
                    "count" => Excel.XlTotalsCalculation.xlTotalsCalculationCount,
                    "min" => Excel.XlTotalsCalculation.xlTotalsCalculationMin,
                    "max" => Excel.XlTotalsCalculation.xlTotalsCalculationMax,
                    "none" => Excel.XlTotalsCalculation.xlTotalsCalculationNone,
                    _ => Excel.XlTotalsCalculation.xlTotalsCalculationSum,
                };
            }

            return Response.Ok(new { name, totals = table.ShowTotals });
        }
        catch (Exception ex)
        {
            return Response.Error($"table.totals: {ex.Message}");
        }
    }

    private string HandleFilter(JsonObject args)
    {
        try
        {
            var name = args["name"]?.GetValue<string>()
                ?? throw new ArgumentException("table.filter requires 'name'");
            var columnName = args["column"]?.GetValue<string>()
                ?? throw new ArgumentException("table.filter requires 'column'");
            var value = args["value"]?.ToString()
                ?? throw new ArgumentException("table.filter requires 'value'");

            using var guard = new ComGuard();
            var sheet = guard.Track(_session.GetActiveSheet());
            var table = FindTable(sheet, name, guard);

            // Find column index
            var columns = guard.Track(table.ListColumns);
            int colIndex = -1;
            foreach (Excel.ListColumn col in columns)
            {
                if (string.Equals(col.Name, columnName, StringComparison.OrdinalIgnoreCase))
                {
                    colIndex = col.Index;
                    Marshal.ReleaseComObject(col);
                    break;
                }
                Marshal.ReleaseComObject(col);
            }

            if (colIndex < 0)
                return Response.Error($"Column not found: {columnName}");

            var range = guard.Track(table.Range);
            range.AutoFilter(colIndex, value);

            return Response.Ok(new { name, column = columnName, value, filtered = true });
        }
        catch (Exception ex)
        {
            return Response.Error($"table.filter: {ex.Message}");
        }
    }

    private string HandleFilterClear(JsonObject args)
    {
        try
        {
            var name = args["name"]?.GetValue<string>()
                ?? throw new ArgumentException("table.filter.clear requires 'name'");

            using var guard = new ComGuard();
            var sheet = guard.Track(_session.GetActiveSheet());
            var table = FindTable(sheet, name, guard);

            if (table.AutoFilter != null)
            {
                var af = guard.Track(table.AutoFilter);
                af.ShowAllData();
            }

            return Response.Ok(new { name, filters_cleared = true });
        }
        catch (Exception ex)
        {
            return Response.Error($"table.filter.clear: {ex.Message}");
        }
    }

    private string HandleSort(JsonObject args)
    {
        try
        {
            var name = args["name"]?.GetValue<string>()
                ?? throw new ArgumentException("table.sort requires 'name'");
            var columnName = args["column"]?.GetValue<string>()
                ?? throw new ArgumentException("table.sort requires 'column'");
            var order = args["order"]?.GetValue<string>() ?? "asc";

            using var guard = new ComGuard();
            var sheet = guard.Track(_session.GetActiveSheet());
            var table = FindTable(sheet, name, guard);

            // Find column range for sort key
            var columns = guard.Track(table.ListColumns);
            Excel.ListColumn? targetCol = null;
            foreach (Excel.ListColumn col in columns)
            {
                if (string.Equals(col.Name, columnName, StringComparison.OrdinalIgnoreCase))
                {
                    targetCol = col;
                    break;
                }
                Marshal.ReleaseComObject(col);
            }

            if (targetCol == null)
                return Response.Error($"Column not found: {columnName}");

            var keyRange = guard.Track(targetCol.DataBodyRange);
            var sort = guard.Track(table.Sort);
            var sortFields = guard.Track(sort.SortFields);
            sortFields.Clear();

            var xlOrder = order.ToLowerInvariant() switch
            {
                "desc" or "descending" => Excel.XlSortOrder.xlDescending,
                _ => Excel.XlSortOrder.xlAscending,
            };

            sortFields.Add(keyRange, Order: xlOrder);
            sort.Header = Excel.XlYesNoGuess.xlYes;
            sort.Apply();

            Marshal.ReleaseComObject(targetCol);

            return Response.Ok(new { name, column = columnName, order, sorted = true });
        }
        catch (Exception ex)
        {
            return Response.Error($"table.sort: {ex.Message}");
        }
    }

    private string HandleRowAdd(JsonObject args)
    {
        try
        {
            var name = args["name"]?.GetValue<string>()
                ?? throw new ArgumentException("table.row.add requires 'name'");
            var values = args["values"] as JsonArray
                ?? throw new ArgumentException("table.row.add requires 'values' array");

            using var guard = new ComGuard();
            var sheet = guard.Track(_session.GetActiveSheet());
            var table = FindTable(sheet, name, guard);
            var rows = guard.Track(table.ListRows);
            var newRow = guard.Track(rows.Add());
            var rowRange = guard.Track(newRow.Range);

            for (int i = 0; i < values.Count; i++)
            {
                var cell = guard.Track(rowRange.Cells[1, i + 1] as Excel.Range
                    ?? throw new InvalidOperationException("Cannot access cell"));
                var v = values[i];
                if (v is JsonValue jv && jv.TryGetValue<double>(out var dv))
                    cell.Value2 = dv;
                else if (v is JsonValue bv && bv.TryGetValue<bool>(out var bval))
                    cell.Value2 = bval;
                else
                    cell.Value2 = v?.ToString() ?? "";
            }

            return Response.Ok(new { name, row_index = newRow.Index, added = true });
        }
        catch (Exception ex)
        {
            return Response.Error($"table.row.add: {ex.Message}");
        }
    }

    private string HandleColumnAdd(JsonObject args)
    {
        try
        {
            var name = args["name"]?.GetValue<string>()
                ?? throw new ArgumentException("table.column.add requires 'name'");
            var header = args["header"]?.GetValue<string>()
                ?? throw new ArgumentException("table.column.add requires 'header'");
            var formula = args["formula"]?.GetValue<string>();

            using var guard = new ComGuard();
            var sheet = guard.Track(_session.GetActiveSheet());
            var table = FindTable(sheet, name, guard);
            var columns = guard.Track(table.ListColumns);
            var newCol = guard.Track(columns.Add());
            newCol.Name = header;

            if (formula != null)
            {
                var dataRange = guard.Track(newCol.DataBodyRange);
                dataRange.Formula = formula;
            }

            return Response.Ok(new { name, header, column_index = newCol.Index, added = true });
        }
        catch (Exception ex)
        {
            return Response.Error($"table.column.add: {ex.Message}");
        }
    }

    private string HandleData(JsonObject args)
    {
        try
        {
            var name = args["name"]?.GetValue<string>()
                ?? throw new ArgumentException("table.data requires 'name'");

            using var guard = new ComGuard();
            var sheet = guard.Track(_session.GetActiveSheet());
            var table = FindTable(sheet, name, guard);

            // Collect headers
            var columns = guard.Track(table.ListColumns);
            var headers = new List<string>();
            foreach (Excel.ListColumn col in columns)
            {
                headers.Add(col.Name);
                Marshal.ReleaseComObject(col);
            }

            // Collect row data
            var rows = guard.Track(table.ListRows);
            var result = new JsonArray();
            foreach (Excel.ListRow row in rows)
            {
                var rowRange = guard.Track(row.Range);
                var obj = new JsonObject();
                for (int c = 0; c < headers.Count; c++)
                {
                    var cell = rowRange.Cells[1, c + 1] as Excel.Range;
                    if (cell != null)
                    {
                        var val = cell.Value2;
                        if (val is double d)
                            obj[headers[c]] = d;
                        else if (val is bool b)
                            obj[headers[c]] = b;
                        else if (val != null)
                            obj[headers[c]] = val.ToString();
                        else
                            obj[headers[c]] = null;
                        Marshal.ReleaseComObject(cell);
                    }
                }
                result.Add(obj);
                Marshal.ReleaseComObject(row);
            }

            return Response.Ok(new { name, row_count = result.Count, data = result });
        }
        catch (Exception ex)
        {
            return Response.Error($"table.data: {ex.Message}");
        }
    }

    // ── Helpers ──────────────────────────────────────────────────────────

    private static Excel.ListObject FindTable(Excel.Worksheet sheet, string name, ComGuard guard)
    {
        var tables = guard.Track(sheet.ListObjects);
        foreach (Excel.ListObject tbl in tables)
        {
            if (string.Equals(tbl.Name, name, StringComparison.OrdinalIgnoreCase))
                return tbl;
            Marshal.ReleaseComObject(tbl);
        }
        throw new ArgumentException($"Table not found: {name}");
    }
}
