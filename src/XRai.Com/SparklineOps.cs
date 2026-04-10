using System.Runtime.InteropServices;
using System.Text.Json.Nodes;
using XRai.Core;
using Excel = Microsoft.Office.Interop.Excel;

namespace XRai.Com;

public class SparklineOps
{
    private readonly ExcelSession _session;

    public SparklineOps(ExcelSession session) { _session = session; }

    public void Register(CommandRouter router)
    {
        router.Register("sparkline.list", HandleList);
        router.Register("sparkline.create", HandleCreate);
        router.Register("sparkline.delete", HandleDelete);
    }

    private string HandleList(JsonObject args)
    {
        try
        {
            using var guard = new ComGuard();
            var sheet = guard.Track(_session.GetActiveSheet());
            dynamic groups = ((dynamic)sheet).SparklineGroups;
            guard.Track((object)groups);

            int groupCount = (int)groups.Count;
            var result = new JsonArray();
            for (int i = 1; i <= groupCount; i++)
            {
                try
                {
                    dynamic grp = groups[i];
                    guard.Track((object)grp);
                    var entry = new JsonObject
                    {
                        ["index"] = i - 1,
                        ["count"] = (int)grp.Count,
                    };

                    try
                    {
                        int sparkType = (int)grp.Type;
                        entry["type"] = sparkType switch
                        {
                            // xlSparkLine = 1, xlSparkColumn = 2, xlSparkColumnStacked100 = 3
                            1 => "line",
                            2 => "column",
                            3 => "win_loss",
                            _ => "unknown",
                        };
                    }
                    catch { entry["type"] = "unknown"; }

                    try { entry["source_data"] = (string)grp.SourceData; } catch { }
                    try
                    {
                        Excel.Range loc = grp.Location;
                        entry["location"] = loc.Address[false, false];
                        Marshal.ReleaseComObject(loc);
                    }
                    catch { }

                    result.Add(entry);
                }
                catch
                {
                    // Skip groups that cannot be read
                }
            }

            return Response.Ok(new { sparkline_groups = result, count = result.Count });
        }
        catch (Exception ex)
        {
            return Response.Error($"sparkline.list: {ex.Message}");
        }
    }

    private string HandleCreate(JsonObject args)
    {
        try
        {
            var dataRef = args["data"]?.GetValue<string>()
                ?? throw new ArgumentException("sparkline.create requires 'data'");
            var locationRef = args["location"]?.GetValue<string>()
                ?? throw new ArgumentException("sparkline.create requires 'location'");
            var typeStr = args["type"]?.GetValue<string>() ?? "line";

            // xlSparkLine = 1, xlSparkColumn = 2, xlSparkColumnStacked100 = 3
            int xlType = typeStr.ToLowerInvariant() switch
            {
                "column" => 2,
                "win_loss" => 3,
                _ => 1,
            };

            using var guard = new ComGuard();
            var sheet = guard.Track(_session.GetActiveSheet());
            var location = guard.Track(sheet.Range[locationRef]);
            dynamic groups = ((dynamic)sheet).SparklineGroups;
            guard.Track((object)groups);

            groups.Add((Excel.XlSparkType)xlType, dataRef, location);

            return Response.Ok(new { data = dataRef, location = locationRef, type = typeStr, created = true });
        }
        catch (Exception ex)
        {
            return Response.Error($"sparkline.create: {ex.Message}");
        }
    }

    private string HandleDelete(JsonObject args)
    {
        try
        {
            var all = args["all"]?.GetValue<bool>() ?? false;

            using var guard = new ComGuard();
            var sheet = guard.Track(_session.GetActiveSheet());
            dynamic groups = ((dynamic)sheet).SparklineGroups;
            guard.Track((object)groups);

            if (all)
            {
                groups.ClearGroups();
                return Response.Ok(new { deleted_all = true });
            }

            var index = args["index"]?.GetValue<int>()
                ?? throw new ArgumentException("sparkline.delete requires 'index' or 'all':true");

            // SparklineGroups is 1-based
            dynamic grp = groups[index + 1];
            guard.Track((object)grp);
            grp.Delete();

            return Response.Ok(new { index, deleted = true });
        }
        catch (Exception ex)
        {
            return Response.Error($"sparkline.delete: {ex.Message}");
        }
    }
}
