using System.Runtime.InteropServices;
using System.Text.Json.Nodes;
using XRai.Core;
using Excel = Microsoft.Office.Interop.Excel;

namespace XRai.Com;

public class PivotOps
{
    private readonly ExcelSession _session;

    public PivotOps(ExcelSession session) { _session = session; }

    public void Register(CommandRouter router)
    {
        router.Register("pivot.list", HandlePivotList);
        router.Register("pivot.create", HandlePivotCreate);
        router.Register("pivot.refresh", HandlePivotRefresh);
        router.Register("pivot.field.add", HandlePivotFieldAdd);
        router.Register("pivot.field.remove", HandlePivotFieldRemove);
        router.Register("pivot.style", HandlePivotStyle);
        router.Register("pivot.data", HandlePivotData);
        router.Register("pivot.field.format", HandlePivotFieldFormat);
        router.Register("pivot.field.filter", HandlePivotFieldFilter);
        router.Register("pivot.field.sort", HandlePivotFieldSort);
        router.Register("pivot.field.function", HandlePivotFieldFunction);
        router.Register("pivot.layout", HandlePivotLayout);
        router.Register("pivot.subtotals", HandlePivotSubtotals);
        router.Register("pivot.grandtotal", HandlePivotGrandTotal);
        router.Register("pivot.calculated", HandlePivotCalculated);
    }

    private string HandlePivotList(JsonObject args)
    {
        using var guard = new ComGuard();
        var sheet = guard.Track(_session.GetActiveSheet());
        var pivots = guard.Track(sheet.PivotTables());

        var list = new JsonArray();
        for (int i = 1; i <= pivots.Count; i++)
        {
            var pt = guard.Track((Excel.PivotTable)pivots.Item(i));
            var tableRange = guard.Track(pt.TableRange1);
            var sourceData = string.Empty;
            try { sourceData = pt.SourceData?.ToString(); } catch { }

            list.Add(new JsonObject
            {
                ["name"] = pt.Name,
                ["source"] = sourceData,
                ["location"] = tableRange.Address[false, false],
            });
        }

        return Response.Ok(new { count = pivots.Count, pivots = list });
    }

    private string HandlePivotCreate(JsonObject args)
    {
        var source = args["source"]?.GetValue<string>()
            ?? throw new ArgumentException("pivot.create requires 'source'");
        var destination = args["destination"]?.GetValue<string>()
            ?? throw new ArgumentException("pivot.create requires 'destination'");
        var name = args["name"]?.GetValue<string>()
            ?? throw new ArgumentException("pivot.create requires 'name'");

        using var guard = new ComGuard();
        var wb = guard.Track(_session.GetActiveWorkbook());
        var sheet = guard.Track(_session.GetActiveSheet());
        var sourceRange = guard.Track(sheet.Range[source]);
        var destRange = guard.Track(sheet.Range[destination]);

        var cache = guard.Track(wb.PivotCaches().Create(
            Excel.XlPivotTableSourceType.xlDatabase,
            sourceRange));

        var pt = guard.Track(cache.CreatePivotTable(destRange, name));

        return Response.Ok(new { name, source, destination, created = true });
    }

    private string HandlePivotRefresh(JsonObject args)
    {
        var name = args["name"]?.GetValue<string>()
            ?? throw new ArgumentException("pivot.refresh requires 'name'");

        using var guard = new ComGuard();
        var sheet = guard.Track(_session.GetActiveSheet());
        var pt = guard.Track((Excel.PivotTable)sheet.PivotTables(name));
        pt.RefreshTable();

        return Response.Ok(new { name, refreshed = true });
    }

    private string HandlePivotFieldAdd(JsonObject args)
    {
        var name = args["name"]?.GetValue<string>()
            ?? throw new ArgumentException("pivot.field.add requires 'name'");
        var field = args["field"]?.GetValue<string>()
            ?? throw new ArgumentException("pivot.field.add requires 'field'");
        var area = args["area"]?.GetValue<string>()
            ?? throw new ArgumentException("pivot.field.add requires 'area'");
        var function = args["function"]?.GetValue<string>();

        using var guard = new ComGuard();
        var sheet = guard.Track(_session.GetActiveSheet());
        var pt = guard.Track((Excel.PivotTable)sheet.PivotTables(name));
        var pfObj = pt.PivotFields(field) as Excel.PivotField
            ?? throw new ArgumentException($"Field '{field}' not found in pivot table '{name}'");
        var pf = guard.Track(pfObj);

        var xlArea = area.ToLowerInvariant() switch
        {
            "row" => Excel.XlPivotFieldOrientation.xlRowField,
            "column" or "col" => Excel.XlPivotFieldOrientation.xlColumnField,
            "value" or "data" => Excel.XlPivotFieldOrientation.xlDataField,
            "filter" or "page" => Excel.XlPivotFieldOrientation.xlPageField,
            _ => throw new ArgumentException($"Unknown area: {area}. Use row/column/value/filter"),
        };

        pf.Orientation = xlArea;

        if (xlArea == Excel.XlPivotFieldOrientation.xlDataField && function != null)
        {
            pf.Function = function.ToLowerInvariant() switch
            {
                "sum" => Excel.XlConsolidationFunction.xlSum,
                "count" => Excel.XlConsolidationFunction.xlCount,
                "average" or "avg" => Excel.XlConsolidationFunction.xlAverage,
                "min" => Excel.XlConsolidationFunction.xlMin,
                "max" => Excel.XlConsolidationFunction.xlMax,
                _ => Excel.XlConsolidationFunction.xlSum,
            };
        }

        return Response.Ok(new { name, field, area, added = true });
    }

    private string HandlePivotFieldRemove(JsonObject args)
    {
        var name = args["name"]?.GetValue<string>()
            ?? throw new ArgumentException("pivot.field.remove requires 'name'");
        var field = args["field"]?.GetValue<string>()
            ?? throw new ArgumentException("pivot.field.remove requires 'field'");

        using var guard = new ComGuard();
        var sheet = guard.Track(_session.GetActiveSheet());
        var pt = guard.Track((Excel.PivotTable)sheet.PivotTables(name));
        var pfObj = pt.PivotFields(field) as Excel.PivotField
            ?? throw new ArgumentException($"Field '{field}' not found in pivot table '{name}'");
        var pf = guard.Track(pfObj);

        pf.Orientation = Excel.XlPivotFieldOrientation.xlHidden;

        return Response.Ok(new { name, field, removed = true });
    }

    private string HandlePivotStyle(JsonObject args)
    {
        var name = args["name"]?.GetValue<string>()
            ?? throw new ArgumentException("pivot.style requires 'name'");
        var style = args["style"]?.GetValue<string>()
            ?? throw new ArgumentException("pivot.style requires 'style'");

        using var guard = new ComGuard();
        var sheet = guard.Track(_session.GetActiveSheet());
        var pt = guard.Track((Excel.PivotTable)sheet.PivotTables(name));
        pt.TableStyle2 = style;

        return Response.Ok(new { name, style, applied = true });
    }

    private string HandlePivotData(JsonObject args)
    {
        var name = args["name"]?.GetValue<string>()
            ?? throw new ArgumentException("pivot.data requires 'name'");

        using var guard = new ComGuard();
        var sheet = guard.Track(_session.GetActiveSheet());
        var pt = guard.Track((Excel.PivotTable)sheet.PivotTables(name));
        var tableRange = guard.Track(pt.TableRange1);

        int rows = tableRange.Rows.Count;
        int cols = tableRange.Columns.Count;
        var values = tableRange.Value2;

        var data = new JsonArray();
        for (int r = 1; r <= rows; r++)
        {
            var row = new JsonArray();
            for (int c = 1; c <= cols; c++)
            {
                var val = values is object[,] arr ? arr[r, c] : null;
                if (val is double d)
                    row.Add(d);
                else if (val != null)
                    row.Add(val.ToString());
                else
                    row.Add(null);
            }
            data.Add(row);
        }

        return Response.Ok(new
        {
            name,
            range = tableRange.Address[false, false],
            rows,
            columns = cols,
            data,
        });
    }

    private string HandlePivotFieldFormat(JsonObject args)
    {
        var name = args["name"]?.GetValue<string>()
            ?? throw new ArgumentException("pivot.field.format requires 'name'");
        var field = args["field"]?.GetValue<string>()
            ?? throw new ArgumentException("pivot.field.format requires 'field'");
        var numberFormat = args["number_format"]?.GetValue<string>();
        var caption = args["caption"]?.GetValue<string>();

        using var guard = new ComGuard();
        var sheet = guard.Track(_session.GetActiveSheet());
        var pt = guard.Track((Excel.PivotTable)sheet.PivotTables(name));
        var pf = guard.Track((Excel.PivotField)pt.PivotFields(field));

        if (numberFormat != null)
            pf.NumberFormat = numberFormat;
        if (caption != null)
            pf.Caption = caption;

        return Response.Ok(new { name, field, number_format = numberFormat, caption, formatted = true });
    }

    private string HandlePivotFieldFilter(JsonObject args)
    {
        var name = args["name"]?.GetValue<string>()
            ?? throw new ArgumentException("pivot.field.filter requires 'name'");
        var field = args["field"]?.GetValue<string>()
            ?? throw new ArgumentException("pivot.field.filter requires 'field'");
        var itemsArr = args["items"]?.AsArray()
            ?? throw new ArgumentException("pivot.field.filter requires 'items' (array of visible item names)");

        var visibleNames = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        foreach (var item in itemsArr)
            if (item != null) visibleNames.Add(item.GetValue<string>());

        using var guard = new ComGuard();
        var sheet = guard.Track(_session.GetActiveSheet());
        var pt = guard.Track((Excel.PivotTable)sheet.PivotTables(name));
        var pf = guard.Track((Excel.PivotField)pt.PivotFields(field));

        // Set visible items
        var pivotItems = pf.PivotItems() as object[];
        if (pivotItems != null)
        {
            foreach (Excel.PivotItem pi in pivotItems)
            {
                pi.Visible = visibleNames.Contains(pi.Name);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(pi);
            }
        }
        else
        {
            // Fallback: use dynamic iteration
            dynamic dynField = pf;
            dynamic items = dynField.PivotItems();
            int count = items.Count;
            for (int i = 1; i <= count; i++)
            {
                dynamic pi = items[i];
                string piName = pi.Name;
                pi.Visible = visibleNames.Contains(piName);
            }
        }

        return Response.Ok(new { name, field, visible_count = visibleNames.Count, filtered = true });
    }

    private string HandlePivotFieldSort(JsonObject args)
    {
        var name = args["name"]?.GetValue<string>()
            ?? throw new ArgumentException("pivot.field.sort requires 'name'");
        var field = args["field"]?.GetValue<string>()
            ?? throw new ArgumentException("pivot.field.sort requires 'field'");
        var order = args["order"]?.GetValue<string>() ?? "ascending";

        using var guard = new ComGuard();
        var sheet = guard.Track(_session.GetActiveSheet());
        var pt = guard.Track((Excel.PivotTable)sheet.PivotTables(name));
        var pf = guard.Track((Excel.PivotField)pt.PivotFields(field));

        var xlOrder = order.ToLowerInvariant() switch
        {
            "descending" or "desc" => Excel.XlSortOrder.xlDescending,
            _ => Excel.XlSortOrder.xlAscending,
        };

        pf.AutoSort((int)xlOrder, field);

        return Response.Ok(new { name, field, order, sorted = true });
    }

    private string HandlePivotFieldFunction(JsonObject args)
    {
        var name = args["name"]?.GetValue<string>()
            ?? throw new ArgumentException("pivot.field.function requires 'name'");
        var field = args["field"]?.GetValue<string>()
            ?? throw new ArgumentException("pivot.field.function requires 'field'");
        var function_ = args["function"]?.GetValue<string>()
            ?? throw new ArgumentException("pivot.field.function requires 'function'");

        using var guard = new ComGuard();
        var sheet = guard.Track(_session.GetActiveSheet());
        var pt = guard.Track((Excel.PivotTable)sheet.PivotTables(name));
        var pf = guard.Track((Excel.PivotField)pt.PivotFields(field));

        pf.Function = function_.ToLowerInvariant() switch
        {
            "sum" => Excel.XlConsolidationFunction.xlSum,
            "count" => Excel.XlConsolidationFunction.xlCount,
            "average" or "avg" => Excel.XlConsolidationFunction.xlAverage,
            "min" => Excel.XlConsolidationFunction.xlMin,
            "max" => Excel.XlConsolidationFunction.xlMax,
            "product" => Excel.XlConsolidationFunction.xlProduct,
            "stdev" => Excel.XlConsolidationFunction.xlStDev,
            "var" => Excel.XlConsolidationFunction.xlVar,
            "countnums" => Excel.XlConsolidationFunction.xlCountNums,
            _ => throw new ArgumentException($"Unknown function: {function_}. Use sum/count/average/min/max/product/stdev/var/countnums"),
        };

        return Response.Ok(new { name, field, function = function_, updated = true });
    }

    private string HandlePivotLayout(JsonObject args)
    {
        var name = args["name"]?.GetValue<string>()
            ?? throw new ArgumentException("pivot.layout requires 'name'");
        var mode = args["mode"]?.GetValue<string>()
            ?? throw new ArgumentException("pivot.layout requires 'mode' (tabular/outline/compact)");

        using var guard = new ComGuard();
        var sheet = guard.Track(_session.GetActiveSheet());
        var pt = guard.Track((Excel.PivotTable)sheet.PivotTables(name));

        switch (mode.ToLowerInvariant())
        {
            case "tabular":
                pt.RowAxisLayout(Excel.XlLayoutRowType.xlTabularRow);
                break;
            case "outline":
                pt.RowAxisLayout(Excel.XlLayoutRowType.xlOutlineRow);
                break;
            case "compact":
                pt.RowAxisLayout(Excel.XlLayoutRowType.xlCompactRow);
                break;
            default:
                return Response.Error($"Unknown layout mode: {mode}. Use tabular/outline/compact");
        }

        return Response.Ok(new { name, mode, applied = true });
    }

    private string HandlePivotSubtotals(JsonObject args)
    {
        var name = args["name"]?.GetValue<string>()
            ?? throw new ArgumentException("pivot.subtotals requires 'name'");
        var field = args["field"]?.GetValue<string>()
            ?? throw new ArgumentException("pivot.subtotals requires 'field'");
        var show = args["show"]?.GetValue<bool>() ?? true;

        using var guard = new ComGuard();
        var sheet = guard.Track(_session.GetActiveSheet());
        var pt = guard.Track((Excel.PivotTable)sheet.PivotTables(name));
        var pf = guard.Track((Excel.PivotField)pt.PivotFields(field));

        // Subtotals is a 12-element boolean array
        // Index 1 = Automatic, rest are specific functions
        var subtotals = new bool[13]; // 1-based indexing in COM
        if (show)
            subtotals[1] = true; // Automatic subtotals
        pf.Subtotals = subtotals;

        return Response.Ok(new { name, field, subtotals_visible = show });
    }

    private string HandlePivotGrandTotal(JsonObject args)
    {
        var name = args["name"]?.GetValue<string>()
            ?? throw new ArgumentException("pivot.grandtotal requires 'name'");
        var rows = args["rows"]?.GetValue<bool>();
        var columns = args["columns"]?.GetValue<bool>();

        using var guard = new ComGuard();
        var sheet = guard.Track(_session.GetActiveSheet());
        var pt = guard.Track((Excel.PivotTable)sheet.PivotTables(name));

        if (rows.HasValue)
            pt.RowGrand = rows.Value;
        if (columns.HasValue)
            pt.ColumnGrand = columns.Value;

        return Response.Ok(new { name, row_grand = pt.RowGrand, column_grand = pt.ColumnGrand });
    }

    private string HandlePivotCalculated(JsonObject args)
    {
        var name = args["name"]?.GetValue<string>()
            ?? throw new ArgumentException("pivot.calculated requires 'name'");
        var fieldName = args["field"]?.GetValue<string>()
            ?? throw new ArgumentException("pivot.calculated requires 'field'");
        var formula = args["formula"]?.GetValue<string>()
            ?? throw new ArgumentException("pivot.calculated requires 'formula'");

        using var guard = new ComGuard();
        var sheet = guard.Track(_session.GetActiveSheet());
        var pt = guard.Track((Excel.PivotTable)sheet.PivotTables(name));

        var calcFields = guard.Track(pt.CalculatedFields());
        calcFields.Add(fieldName, formula);

        return Response.Ok(new { name, field = fieldName, formula, created = true });
    }
}
