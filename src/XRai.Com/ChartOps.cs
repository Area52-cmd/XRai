using System.Runtime.InteropServices;
using System.Text.Json.Nodes;
using XRai.Core;
using Excel = Microsoft.Office.Interop.Excel;

namespace XRai.Com;

public class ChartOps
{
    private readonly ExcelSession _session;

    public ChartOps(ExcelSession session) { _session = session; }

    public void Register(CommandRouter router)
    {
        router.Register("chart.list", HandleList);
        router.Register("chart.create", HandleCreate);
        router.Register("chart.delete", HandleDelete);
        router.Register("chart.type", HandleType);
        router.Register("chart.title", HandleTitle);
        router.Register("chart.data", HandleData);
        router.Register("chart.legend", HandleLegend);
        router.Register("chart.axis", HandleAxis);
        router.Register("chart.series", HandleSeries);
        router.Register("chart.export", HandleExport);
        router.Register("chart.trendline", HandleTrendline);
        router.Register("chart.datalabel", HandleDataLabel);
        router.Register("chart.gridlines", HandleGridlines);
        router.Register("chart.axis.scale", HandleAxisScale);
    }

    private string HandleList(JsonObject args)
    {
        try
        {
            using var guard = new ComGuard();
            var sheet = guard.Track(_session.GetActiveSheet());
            var chartObjects = guard.Track(sheet.ChartObjects() as Excel.ChartObjects)
                ?? throw new InvalidOperationException("Cannot access chart objects");

            var result = new JsonArray();
            foreach (Excel.ChartObject co in chartObjects)
            {
                var chart = guard.Track(co.Chart);
                result.Add(new JsonObject
                {
                    ["name"] = co.Name,
                    ["type"] = chart.ChartType.ToString(),
                    ["left"] = co.Left,
                    ["top"] = co.Top,
                    ["width"] = co.Width,
                    ["height"] = co.Height,
                });
                Marshal.ReleaseComObject(chart);
                Marshal.ReleaseComObject(co);
            }

            return Response.Ok(new { charts = result, count = result.Count });
        }
        catch (Exception ex)
        {
            return Response.Error($"chart.list: {ex.Message}");
        }
    }

    private string HandleCreate(JsonObject args)
    {
        try
        {
            var typeStr = args["type"]?.GetValue<string>() ?? "column";
            var dataRef = args["data"]?.GetValue<string>()
                ?? throw new ArgumentException("chart.create requires 'data'");
            var title = args["title"]?.GetValue<string>();
            var position = args["position"]?.GetValue<string>();

            using var guard = new ComGuard();
            var sheet = guard.Track(_session.GetActiveSheet());
            var dataRange = guard.Track(sheet.Range[dataRef]);

            // Determine position from a cell reference or default
            double left = 100, top = 100;
            if (position != null)
            {
                var posRange = guard.Track(sheet.Range[position]);
                left = (double)posRange.Left;
                top = (double)posRange.Top;
            }

            var chartObjects = guard.Track(sheet.ChartObjects() as Excel.ChartObjects)
                ?? throw new InvalidOperationException("Cannot access chart objects");
            var chartObj = guard.Track(chartObjects.Add(left, top, 480, 300));
            var chart = guard.Track(chartObj.Chart);

            chart.ChartType = ParseChartType(typeStr);
            chart.SetSourceData(dataRange);

            if (title != null)
            {
                chart.HasTitle = true;
                var chartTitle = guard.Track(chart.ChartTitle);
                chartTitle.Text = title;
            }

            return Response.Ok(new
            {
                name = chartObj.Name,
                type = typeStr,
                data = dataRef,
                title,
                created = true,
            });
        }
        catch (Exception ex)
        {
            return Response.Error($"chart.create: {ex.Message}");
        }
    }

    private string HandleDelete(JsonObject args)
    {
        try
        {
            var name = args["name"]?.GetValue<string>()
                ?? throw new ArgumentException("chart.delete requires 'name'");

            using var guard = new ComGuard();
            var sheet = guard.Track(_session.GetActiveSheet());
            var chartObj = guard.Track(
                (sheet.ChartObjects(name) as Excel.ChartObject)
                ?? throw new ArgumentException($"Chart not found: {name}"));
            chartObj.Delete();

            return Response.Ok(new { name, deleted = true });
        }
        catch (Exception ex)
        {
            return Response.Error($"chart.delete: {ex.Message}");
        }
    }

    private string HandleType(JsonObject args)
    {
        try
        {
            var name = args["name"]?.GetValue<string>()
                ?? throw new ArgumentException("chart.type requires 'name'");

            using var guard = new ComGuard();
            var sheet = guard.Track(_session.GetActiveSheet());
            var chartObj = guard.Track(
                (sheet.ChartObjects(name) as Excel.ChartObject)
                ?? throw new ArgumentException($"Chart not found: {name}"));
            var chart = guard.Track(chartObj.Chart);

            if (args["type"] is JsonNode typeNode)
            {
                var newType = ParseChartType(typeNode.GetValue<string>());
                chart.ChartType = newType;
                return Response.Ok(new { name, type = typeNode.GetValue<string>(), updated = true });
            }

            return Response.Ok(new { name, type = chart.ChartType.ToString() });
        }
        catch (Exception ex)
        {
            return Response.Error($"chart.type: {ex.Message}");
        }
    }

    private string HandleTitle(JsonObject args)
    {
        try
        {
            var name = args["name"]?.GetValue<string>()
                ?? throw new ArgumentException("chart.title requires 'name'");

            using var guard = new ComGuard();
            var sheet = guard.Track(_session.GetActiveSheet());
            var chartObj = guard.Track(
                (sheet.ChartObjects(name) as Excel.ChartObject)
                ?? throw new ArgumentException($"Chart not found: {name}"));
            var chart = guard.Track(chartObj.Chart);

            if (args["title"] is JsonNode titleNode)
            {
                chart.HasTitle = true;
                var chartTitle = guard.Track(chart.ChartTitle);
                chartTitle.Text = titleNode.GetValue<string>();
                return Response.Ok(new { name, title = titleNode.GetValue<string>(), updated = true });
            }

            if (chart.HasTitle)
            {
                var chartTitle = guard.Track(chart.ChartTitle);
                return Response.Ok(new { name, title = chartTitle.Text });
            }

            return Response.Ok(new { name, title = (string?)null });
        }
        catch (Exception ex)
        {
            return Response.Error($"chart.title: {ex.Message}");
        }
    }

    private string HandleData(JsonObject args)
    {
        try
        {
            var name = args["name"]?.GetValue<string>()
                ?? throw new ArgumentException("chart.data requires 'name'");

            using var guard = new ComGuard();
            var sheet = guard.Track(_session.GetActiveSheet());
            var chartObj = guard.Track(
                (sheet.ChartObjects(name) as Excel.ChartObject)
                ?? throw new ArgumentException($"Chart not found: {name}"));
            var chart = guard.Track(chartObj.Chart);

            if (args["data"] is JsonNode dataNode)
            {
                var dataRange = guard.Track(sheet.Range[dataNode.GetValue<string>()]);
                chart.SetSourceData(dataRange);
                return Response.Ok(new { name, data = dataNode.GetValue<string>(), updated = true });
            }

            // Read current data source by inspecting series
            var sc = guard.Track(chart.SeriesCollection() as Excel.SeriesCollection)
                ?? throw new InvalidOperationException("Cannot read series");
            var seriesInfo = new JsonArray();
            foreach (Excel.Series s in sc)
            {
                seriesInfo.Add(new JsonObject
                {
                    ["name"] = s.Name,
                    ["formula"] = s.Formula,
                });
                Marshal.ReleaseComObject(s);
            }

            return Response.Ok(new { name, series = seriesInfo });
        }
        catch (Exception ex)
        {
            return Response.Error($"chart.data: {ex.Message}");
        }
    }

    private string HandleLegend(JsonObject args)
    {
        try
        {
            var name = args["name"]?.GetValue<string>()
                ?? throw new ArgumentException("chart.legend requires 'name'");

            using var guard = new ComGuard();
            var sheet = guard.Track(_session.GetActiveSheet());
            var chartObj = guard.Track(
                (sheet.ChartObjects(name) as Excel.ChartObject)
                ?? throw new ArgumentException($"Chart not found: {name}"));
            var chart = guard.Track(chartObj.Chart);

            if (args["show"] is JsonNode showNode)
            {
                var show = showNode.GetValue<bool>();
                chart.HasLegend = show;

                if (show && args["position"] is JsonNode posNode)
                {
                    var legend = guard.Track(chart.Legend);
                    legend.Position = posNode.GetValue<string>().ToLowerInvariant() switch
                    {
                        "top" => Excel.XlLegendPosition.xlLegendPositionTop,
                        "bottom" => Excel.XlLegendPosition.xlLegendPositionBottom,
                        "left" => Excel.XlLegendPosition.xlLegendPositionLeft,
                        "right" => Excel.XlLegendPosition.xlLegendPositionRight,
                        _ => Excel.XlLegendPosition.xlLegendPositionBottom,
                    };
                }

                return Response.Ok(new { name, legend_visible = show });
            }

            return Response.Ok(new { name, legend_visible = chart.HasLegend });
        }
        catch (Exception ex)
        {
            return Response.Error($"chart.legend: {ex.Message}");
        }
    }

    private string HandleAxis(JsonObject args)
    {
        try
        {
            var name = args["name"]?.GetValue<string>()
                ?? throw new ArgumentException("chart.axis requires 'name'");
            var which = args["which"]?.GetValue<string>() ?? "y";

            using var guard = new ComGuard();
            var sheet = guard.Track(_session.GetActiveSheet());
            var chartObj = guard.Track(
                (sheet.ChartObjects(name) as Excel.ChartObject)
                ?? throw new ArgumentException($"Chart not found: {name}"));
            var chart = guard.Track(chartObj.Chart);

            Excel.XlAxisType axisType;
            Excel.XlAxisGroup axisGroup = Excel.XlAxisGroup.xlPrimary;
            switch (which.ToLowerInvariant())
            {
                case "x":
                    axisType = Excel.XlAxisType.xlCategory;
                    break;
                case "secondary_y":
                    axisType = Excel.XlAxisType.xlValue;
                    axisGroup = Excel.XlAxisGroup.xlSecondary;
                    break;
                default: // "y"
                    axisType = Excel.XlAxisType.xlValue;
                    break;
            }

            var axes = guard.Track(chart.Axes(axisType, axisGroup) as Excel.Axis)
                ?? throw new InvalidOperationException($"Cannot access {which} axis");

            if (args["title"] is JsonNode titleNode)
            {
                axes.HasTitle = true;
                var axisTitle = guard.Track(axes.AxisTitle);
                axisTitle.Text = titleNode.GetValue<string>();
            }

            if (args["min"] is JsonNode minNode)
                axes.MinimumScale = minNode.GetValue<double>();

            if (args["max"] is JsonNode maxNode)
                axes.MaximumScale = maxNode.GetValue<double>();

            if (args["format"] is JsonNode fmtNode)
                axes.TickLabels.NumberFormat = fmtNode.GetValue<string>();

            return Response.Ok(new { name, axis = which, configured = true });
        }
        catch (Exception ex)
        {
            return Response.Error($"chart.axis: {ex.Message}");
        }
    }

    private string HandleSeries(JsonObject args)
    {
        try
        {
            var name = args["name"]?.GetValue<string>()
                ?? throw new ArgumentException("chart.series requires 'name'");
            var action = args["action"]?.GetValue<string>() ?? "list";

            using var guard = new ComGuard();
            var sheet = guard.Track(_session.GetActiveSheet());
            var chartObj = guard.Track(
                (sheet.ChartObjects(name) as Excel.ChartObject)
                ?? throw new ArgumentException($"Chart not found: {name}"));
            var chart = guard.Track(chartObj.Chart);
            var sc = guard.Track(chart.SeriesCollection() as Excel.SeriesCollection)
                ?? throw new InvalidOperationException("Cannot access series collection");

            switch (action.ToLowerInvariant())
            {
                case "list":
                {
                    var list = new JsonArray();
                    foreach (Excel.Series s in sc)
                    {
                        list.Add(new JsonObject
                        {
                            ["name"] = s.Name,
                            ["formula"] = s.Formula,
                        });
                        Marshal.ReleaseComObject(s);
                    }
                    return Response.Ok(new { name, series = list });
                }
                case "add":
                {
                    var valuesRef = args["values_ref"]?.GetValue<string>()
                        ?? throw new ArgumentException("chart.series add requires 'values_ref'");
                    var seriesName = args["series_name"]?.GetValue<string>();
                    var valuesRange = guard.Track(sheet.Range[valuesRef]);

                    var newSeries = guard.Track(sc.NewSeries());
                    newSeries.Values = valuesRange;
                    if (seriesName != null)
                        newSeries.Name = seriesName;

                    if (args["color"] is JsonNode colorNode)
                    {
                        var format = guard.Track(newSeries.Format);
                        var fill = guard.Track(format.Fill);
                        fill.ForeColor.RGB = ParseColor(colorNode.GetValue<string>());
                    }

                    return Response.Ok(new { name, series_name = seriesName, added = true });
                }
                case "remove":
                {
                    var seriesName = args["series_name"]?.GetValue<string>()
                        ?? throw new ArgumentException("chart.series remove requires 'series_name'");

                    foreach (Excel.Series s in sc)
                    {
                        if (string.Equals(s.Name, seriesName, StringComparison.OrdinalIgnoreCase))
                        {
                            s.Delete();
                            Marshal.ReleaseComObject(s);
                            return Response.Ok(new { name, series_name = seriesName, removed = true });
                        }
                        Marshal.ReleaseComObject(s);
                    }
                    return Response.Error($"Series not found: {seriesName}");
                }
                case "format":
                {
                    var seriesName = args["series_name"]?.GetValue<string>()
                        ?? throw new ArgumentException("chart.series format requires 'series_name'");

                    foreach (Excel.Series s in sc)
                    {
                        if (string.Equals(s.Name, seriesName, StringComparison.OrdinalIgnoreCase))
                        {
                            if (args["color"] is JsonNode colorNode)
                            {
                                var format = guard.Track(s.Format);
                                var fill = guard.Track(format.Fill);
                                fill.ForeColor.RGB = ParseColor(colorNode.GetValue<string>());
                            }
                            Marshal.ReleaseComObject(s);
                            return Response.Ok(new { name, series_name = seriesName, formatted = true });
                        }
                        Marshal.ReleaseComObject(s);
                    }
                    return Response.Error($"Series not found: {seriesName}");
                }
                default:
                    return Response.Error($"chart.series: unknown action '{action}'");
            }
        }
        catch (Exception ex)
        {
            return Response.Error($"chart.series: {ex.Message}");
        }
    }

    private string HandleExport(JsonObject args)
    {
        try
        {
            var name = args["name"]?.GetValue<string>()
                ?? throw new ArgumentException("chart.export requires 'name'");
            var path = args["path"]?.GetValue<string>()
                ?? throw new ArgumentException("chart.export requires 'path'");

            using var guard = new ComGuard();
            var sheet = guard.Track(_session.GetActiveSheet());
            var chartObj = guard.Track(
                (sheet.ChartObjects(name) as Excel.ChartObject)
                ?? throw new ArgumentException($"Chart not found: {name}"));
            var chart = guard.Track(chartObj.Chart);

            chart.Export(path, "PNG");

            return Response.Ok(new { name, path, exported = true });
        }
        catch (Exception ex)
        {
            return Response.Error($"chart.export: {ex.Message}");
        }
    }

    private string HandleTrendline(JsonObject args)
    {
        try
        {
            var name = args["name"]?.GetValue<string>()
                ?? throw new ArgumentException("chart.trendline requires 'name'");
            var seriesName = args["series_name"]?.GetValue<string>();
            var seriesIndex = args["series_index"]?.GetValue<int>() ?? 1;
            var type = args["type"]?.GetValue<string>() ?? "linear";
            var forward = args["forward"]?.GetValue<int>();
            var backward = args["backward"]?.GetValue<int>();
            var displayEquation = args["display_equation"]?.GetValue<bool>() ?? false;
            var displayRSquared = args["display_r_squared"]?.GetValue<bool>() ?? false;

            using var guard = new ComGuard();
            var sheet = guard.Track(_session.GetActiveSheet());
            var chartObj = guard.Track(
                (sheet.ChartObjects(name) as Excel.ChartObject)
                ?? throw new ArgumentException($"Chart not found: {name}"));
            var chart = guard.Track(chartObj.Chart);
            var sc = guard.Track(chart.SeriesCollection() as Excel.SeriesCollection)
                ?? throw new InvalidOperationException("Cannot access series collection");

            Excel.Series targetSeries;
            if (seriesName != null)
            {
                targetSeries = FindSeriesByName(sc, seriesName);
            }
            else
            {
                targetSeries = (Excel.Series)sc.Item(seriesIndex);
            }
            guard.Track(targetSeries);

            var xlType = type.ToLowerInvariant() switch
            {
                "linear" => Excel.XlTrendlineType.xlLinear,
                "exponential" => Excel.XlTrendlineType.xlExponential,
                "logarithmic" => Excel.XlTrendlineType.xlLogarithmic,
                "polynomial" => Excel.XlTrendlineType.xlPolynomial,
                "power" => Excel.XlTrendlineType.xlPower,
                "moving_average" or "movingaverage" => Excel.XlTrendlineType.xlMovingAvg,
                _ => Excel.XlTrendlineType.xlLinear,
            };

            var trendlines = guard.Track(targetSeries.Trendlines());
            var trendline = guard.Track(trendlines.Add(xlType));

            if (forward.HasValue) trendline.Forward = forward.Value;
            if (backward.HasValue) trendline.Backward = backward.Value;
            trendline.DisplayEquation = displayEquation;
            trendline.DisplayRSquared = displayRSquared;

            if (xlType == Excel.XlTrendlineType.xlPolynomial && args["order"] is JsonNode orderNode)
                trendline.Order = orderNode.GetValue<int>();

            if (xlType == Excel.XlTrendlineType.xlMovingAvg && args["period"] is JsonNode periodNode)
                trendline.Period = periodNode.GetValue<int>();

            return Response.Ok(new { name, type, added = true });
        }
        catch (Exception ex)
        {
            return Response.Error($"chart.trendline: {ex.Message}");
        }
    }

    private string HandleDataLabel(JsonObject args)
    {
        try
        {
            var name = args["name"]?.GetValue<string>()
                ?? throw new ArgumentException("chart.datalabel requires 'name'");
            var seriesName = args["series_name"]?.GetValue<string>();
            var seriesIndex = args["series_index"]?.GetValue<int>() ?? 1;
            var showValues = args["show_values"]?.GetValue<bool>();
            var showPercentage = args["show_percentage"]?.GetValue<bool>();
            var showCategoryName = args["show_category_name"]?.GetValue<bool>();
            var showSeriesName = args["show_series_name"]?.GetValue<bool>();

            using var guard = new ComGuard();
            var sheet = guard.Track(_session.GetActiveSheet());
            var chartObj = guard.Track(
                (sheet.ChartObjects(name) as Excel.ChartObject)
                ?? throw new ArgumentException($"Chart not found: {name}"));
            var chart = guard.Track(chartObj.Chart);
            var sc = guard.Track(chart.SeriesCollection() as Excel.SeriesCollection)
                ?? throw new InvalidOperationException("Cannot access series collection");

            Excel.Series targetSeries;
            if (seriesName != null)
                targetSeries = FindSeriesByName(sc, seriesName);
            else
                targetSeries = (Excel.Series)sc.Item(seriesIndex);
            guard.Track(targetSeries);

            targetSeries.HasDataLabels = true;
            var labels = guard.Track(targetSeries.DataLabels() as Excel.DataLabels)
                ?? throw new InvalidOperationException("Cannot access data labels");

            if (showValues.HasValue) labels.ShowValue = showValues.Value;
            if (showPercentage.HasValue) labels.ShowPercentage = showPercentage.Value;
            if (showCategoryName.HasValue) labels.ShowCategoryName = showCategoryName.Value;
            if (showSeriesName.HasValue) labels.ShowSeriesName = showSeriesName.Value;

            return Response.Ok(new { name, configured = true });
        }
        catch (Exception ex)
        {
            return Response.Error($"chart.datalabel: {ex.Message}");
        }
    }

    private string HandleGridlines(JsonObject args)
    {
        try
        {
            var name = args["name"]?.GetValue<string>()
                ?? throw new ArgumentException("chart.gridlines requires 'name'");
            var which = args["which"]?.GetValue<string>() ?? "y";
            var major = args["major"]?.GetValue<bool>();
            var minor = args["minor"]?.GetValue<bool>();

            using var guard = new ComGuard();
            var sheet = guard.Track(_session.GetActiveSheet());
            var chartObj = guard.Track(
                (sheet.ChartObjects(name) as Excel.ChartObject)
                ?? throw new ArgumentException($"Chart not found: {name}"));
            var chart = guard.Track(chartObj.Chart);

            Excel.XlAxisType axisType = which.ToLowerInvariant() switch
            {
                "x" => Excel.XlAxisType.xlCategory,
                _ => Excel.XlAxisType.xlValue,
            };

            var axis = guard.Track(chart.Axes(axisType, Excel.XlAxisGroup.xlPrimary) as Excel.Axis)
                ?? throw new InvalidOperationException($"Cannot access {which} axis");

            if (major.HasValue) axis.HasMajorGridlines = major.Value;
            if (minor.HasValue) axis.HasMinorGridlines = minor.Value;

            return Response.Ok(new
            {
                name,
                axis = which,
                major_gridlines = axis.HasMajorGridlines,
                minor_gridlines = axis.HasMinorGridlines,
            });
        }
        catch (Exception ex)
        {
            return Response.Error($"chart.gridlines: {ex.Message}");
        }
    }

    private string HandleAxisScale(JsonObject args)
    {
        try
        {
            var name = args["name"]?.GetValue<string>()
                ?? throw new ArgumentException("chart.axis.scale requires 'name'");
            var which = args["which"]?.GetValue<string>() ?? "y";
            var min = args["min"]?.GetValue<double>();
            var max = args["max"]?.GetValue<double>();
            var majorUnit = args["major_unit"]?.GetValue<double>();
            var minorUnit = args["minor_unit"]?.GetValue<double>();

            using var guard = new ComGuard();
            var sheet = guard.Track(_session.GetActiveSheet());
            var chartObj = guard.Track(
                (sheet.ChartObjects(name) as Excel.ChartObject)
                ?? throw new ArgumentException($"Chart not found: {name}"));
            var chart = guard.Track(chartObj.Chart);

            Excel.XlAxisType axisType;
            Excel.XlAxisGroup axisGroup = Excel.XlAxisGroup.xlPrimary;
            switch (which.ToLowerInvariant())
            {
                case "x":
                    axisType = Excel.XlAxisType.xlCategory;
                    break;
                case "secondary_y":
                    axisType = Excel.XlAxisType.xlValue;
                    axisGroup = Excel.XlAxisGroup.xlSecondary;
                    break;
                default:
                    axisType = Excel.XlAxisType.xlValue;
                    break;
            }

            var axis = guard.Track(chart.Axes(axisType, axisGroup) as Excel.Axis)
                ?? throw new InvalidOperationException($"Cannot access {which} axis");

            if (min.HasValue) axis.MinimumScale = min.Value;
            if (max.HasValue) axis.MaximumScale = max.Value;
            if (majorUnit.HasValue) axis.MajorUnit = majorUnit.Value;
            if (minorUnit.HasValue) axis.MinorUnit = minorUnit.Value;

            return Response.Ok(new
            {
                name,
                axis = which,
                min = axis.MinimumScale,
                max = axis.MaximumScale,
                major_unit = axis.MajorUnit,
                minor_unit = axis.MinorUnit,
                configured = true,
            });
        }
        catch (Exception ex)
        {
            return Response.Error($"chart.axis.scale: {ex.Message}");
        }
    }

    // ── Helpers ──────────────────────────────────────────────────────────

    private static Excel.Series FindSeriesByName(Excel.SeriesCollection sc, string name)
    {
        foreach (Excel.Series s in sc)
        {
            if (string.Equals(s.Name, name, StringComparison.OrdinalIgnoreCase))
                return s;
            System.Runtime.InteropServices.Marshal.ReleaseComObject(s);
        }
        throw new ArgumentException($"Series not found: {name}");
    }

    private static Excel.XlChartType ParseChartType(string type)
    {
        return type.ToLowerInvariant() switch
        {
            "column" => Excel.XlChartType.xlColumnClustered,
            "bar" => Excel.XlChartType.xlBarClustered,
            "line" => Excel.XlChartType.xlLine,
            "pie" => Excel.XlChartType.xlPie,
            "scatter" => Excel.XlChartType.xlXYScatter,
            "area" => Excel.XlChartType.xlArea,
            "doughnut" => Excel.XlChartType.xlDoughnut,
            "radar" => Excel.XlChartType.xlRadar,
            "combo" => Excel.XlChartType.xlColumnClustered, // combo handled by adding secondary axis series
            _ => Excel.XlChartType.xlColumnClustered,
        };
    }

    private static int ParseColor(string color)
    {
        if (color.StartsWith("#") && color.Length == 7)
        {
            int r = Convert.ToInt32(color[1..3], 16);
            int g = Convert.ToInt32(color[3..5], 16);
            int b = Convert.ToInt32(color[5..7], 16);
            return r | (g << 8) | (b << 16);
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
