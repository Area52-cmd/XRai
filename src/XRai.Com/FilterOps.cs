using System.Runtime.InteropServices;
using System.Text.Json.Nodes;
using XRai.Core;
using Excel = Microsoft.Office.Interop.Excel;

namespace XRai.Com;

public class FilterOps
{
    private readonly ExcelSession _session;

    public FilterOps(ExcelSession session) { _session = session; }

    public void Register(CommandRouter router)
    {
        router.Register("filter.on", HandleFilterOn);
        router.Register("filter.off", HandleFilterOff);
        router.Register("filter.set", HandleFilterSet);
        router.Register("filter.clear", HandleFilterClear);
        router.Register("filter.read", HandleFilterRead);
    }

    private string HandleFilterOn(JsonObject args)
    {
        var refStr = args["ref"]?.GetValue<string>()
            ?? throw new ArgumentException("filter.on requires 'ref'");

        using var guard = new ComGuard();
        var sheet = guard.Track(_session.GetActiveSheet());
        var range = guard.Track(sheet.Range[refStr]);

        if (sheet.AutoFilterMode)
            return Response.Ok(new { @ref = refStr, auto_filter = "already_enabled" });

        range.AutoFilter(1);
        return Response.Ok(new { @ref = refStr, auto_filter = "enabled" });
    }

    private string HandleFilterOff(JsonObject args)
    {
        using var guard = new ComGuard();
        var sheet = guard.Track(_session.GetActiveSheet());

        if (!sheet.AutoFilterMode)
            return Response.Ok(new { auto_filter = "already_disabled" });

        sheet.AutoFilterMode = false;
        return Response.Ok(new { auto_filter = "disabled" });
    }

    private string HandleFilterSet(JsonObject args)
    {
        var column = args["column"]?.GetValue<int>()
            ?? throw new ArgumentException("filter.set requires 'column' (1-based)");
        var op = args["operator"]?.GetValue<string>();

        using var guard = new ComGuard();
        var sheet = guard.Track(_session.GetActiveSheet());

        if (!sheet.AutoFilterMode)
            throw new InvalidOperationException("AutoFilter is not enabled. Use filter.on first.");

        var autoFilter = guard.Track(sheet.AutoFilter);
        var filterRange = guard.Track(autoFilter.Range);

        // Determine criteria and operator
        if (args["criteria"] is JsonArray criteriaArr)
        {
            var items = new string[criteriaArr.Count];
            for (int i = 0; i < criteriaArr.Count; i++)
                items[i] = criteriaArr[i]?.GetValue<string>() ?? "";

            if (op != null && op.Equals("and", StringComparison.OrdinalIgnoreCase) && items.Length >= 2)
            {
                filterRange.AutoFilter(column,
                    items[0],
                    Excel.XlAutoFilterOperator.xlAnd,
                    items[1]);
            }
            else if (op != null && op.Equals("or", StringComparison.OrdinalIgnoreCase) && items.Length >= 2)
            {
                filterRange.AutoFilter(column,
                    items[0],
                    Excel.XlAutoFilterOperator.xlOr,
                    items[1]);
            }
            else
            {
                // Multi-select filter
                filterRange.AutoFilter(column,
                    items,
                    Excel.XlAutoFilterOperator.xlFilterValues);
            }
        }
        else
        {
            var criteria = args["criteria"]?.GetValue<string>()
                ?? throw new ArgumentException("filter.set requires 'criteria'");

            if (op != null && op.Equals("top10", StringComparison.OrdinalIgnoreCase))
            {
                filterRange.AutoFilter(column,
                    criteria,
                    Excel.XlAutoFilterOperator.xlTop10Items);
            }
            else
            {
                filterRange.AutoFilter(column, criteria);
            }
        }

        return Response.Ok(new { column, filter_set = true });
    }

    private string HandleFilterClear(JsonObject args)
    {
        using var guard = new ComGuard();
        var sheet = guard.Track(_session.GetActiveSheet());

        if (!sheet.AutoFilterMode)
            return Response.Ok(new { filter_clear = "no_autofilter" });

        var autoFilter = guard.Track(sheet.AutoFilter);
        var filters = guard.Track(autoFilter.Filters);

        if (args["column"] != null)
        {
            var column = args["column"]!.GetValue<int>();
            if (column < 1 || column > filters.Count)
                throw new ArgumentException($"column must be between 1 and {filters.Count}");

            var filterRange = guard.Track(autoFilter.Range);
            filterRange.AutoFilter(column);
            return Response.Ok(new { column, filter_cleared = true });
        }

        // Clear all column filters
        for (int i = 1; i <= filters.Count; i++)
        {
            var filter = filters[i];
            if (filter.On)
            {
                var filterRange = guard.Track(autoFilter.Range);
                filterRange.AutoFilter(i);
            }
            Marshal.ReleaseComObject(filter);
        }

        return Response.Ok(new { filter_cleared = "all" });
    }

    private string HandleFilterRead(JsonObject args)
    {
        using var guard = new ComGuard();
        var sheet = guard.Track(_session.GetActiveSheet());

        if (!sheet.AutoFilterMode)
            return Response.Ok(new { auto_filter = false });

        var autoFilter = guard.Track(sheet.AutoFilter);
        var filters = guard.Track(autoFilter.Filters);
        var filterRange = guard.Track(autoFilter.Range);
        var rangeAddr = filterRange.Address[false, false];

        var columns = new JsonArray();
        for (int i = 1; i <= filters.Count; i++)
        {
            var filter = filters[i];
            if (filter.On)
            {
                var entry = new JsonObject
                {
                    ["column"] = i,
                    ["operator"] = filter.Operator.ToString(),
                };

                try { entry["criteria1"] = filter.Criteria1?.ToString(); } catch { }
                try { entry["criteria2"] = filter.Criteria2?.ToString(); } catch { }

                columns.Add(entry);
            }
            Marshal.ReleaseComObject(filter);
        }

        return Response.Ok(new { auto_filter = true, range = rangeAddr, filters = columns });
    }
}
