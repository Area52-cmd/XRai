using System.Text.Json.Nodes;
using XRai.Core;

namespace XRai.Com;

public class SlicerOps
{
    private readonly ExcelSession _session;

    public SlicerOps(ExcelSession session) { _session = session; }

    public void Register(CommandRouter router)
    {
        router.Register("slicer.list", HandleList);
        router.Register("slicer.create", HandleCreate);
        router.Register("slicer.delete", HandleDelete);
        router.Register("slicer.set", HandleSet);
        router.Register("slicer.clear", HandleClear);
        router.Register("slicer.read", HandleRead);
    }

    private string HandleList(JsonObject args)
    {
        try
        {
            using var guard = new ComGuard();
            var wb = guard.Track(_session.GetActiveWorkbook());
            dynamic caches;
            try { caches = ((dynamic)wb).SlicerCaches; }
            catch { return Response.Error("Slicers not available in this Excel version"); }

            var result = new JsonArray();
            int count = caches.Count;
            for (int i = 1; i <= count; i++)
            {
                dynamic cache = caches[i];
                string cacheName = cache.Name;
                string? sourceName = null;
                try { sourceName = cache.SourceName?.ToString(); } catch { }

                var slicers = new JsonArray();
                try
                {
                    int slicerCount = cache.Slicers.Count;
                    for (int j = 1; j <= slicerCount; j++)
                    {
                        dynamic slicer = cache.Slicers[j];
                        slicers.Add(new JsonObject
                        {
                            ["name"] = (string)slicer.Name,
                            ["caption"] = (string)slicer.Caption,
                        });
                    }
                }
                catch { }

                var items = new JsonArray();
                try
                {
                    int itemCount = cache.SlicerItems.Count;
                    for (int j = 1; j <= itemCount; j++)
                    {
                        dynamic item = cache.SlicerItems[j];
                        items.Add(new JsonObject
                        {
                            ["name"] = (string)item.Name,
                            ["selected"] = (bool)item.Selected,
                        });
                    }
                }
                catch { }

                result.Add(new JsonObject
                {
                    ["cache_name"] = cacheName,
                    ["source"] = sourceName,
                    ["slicers"] = slicers,
                    ["items"] = items,
                });
            }

            return Response.Ok(new { count, slicer_caches = result });
        }
        catch (Exception ex)
        {
            return Response.Error($"slicer.list: {ex.Message}");
        }
    }

    private string HandleCreate(JsonObject args)
    {
        try
        {
            var source = args["source"]?.GetValue<string>()
                ?? throw new ArgumentException("slicer.create requires 'source' (PivotTable or Table name)");
            var field = args["field"]?.GetValue<string>()
                ?? throw new ArgumentException("slicer.create requires 'field'");
            var name = args["name"]?.GetValue<string>();
            var position = args["position"]?.GetValue<string>();

            using var guard = new ComGuard();
            var wb = guard.Track(_session.GetActiveWorkbook());
            var sheet = guard.Track(_session.GetActiveSheet());

            dynamic caches;
            try { caches = ((dynamic)wb).SlicerCaches; }
            catch { return Response.Error("Slicers not available in this Excel version"); }

            // Try PivotTable source first, then Table
            dynamic cache;
            try
            {
                // Try as PivotTable
                var pt = sheet.PivotTables(source);
                cache = caches.Add2(pt, field);
            }
            catch
            {
                try
                {
                    // Try as ListObject (Table)
                    dynamic listObj = sheet.ListObjects[source];
                    cache = caches.Add2(listObj, field);
                }
                catch
                {
                    return Response.Error($"Source '{source}' not found as PivotTable or Table on active sheet");
                }
            }

            // Add the slicer visual
            dynamic slicer;
            double left = 10, top = 10;
            if (position != null)
            {
                try
                {
                    var posRange = guard.Track(sheet.Range[position]);
                    left = (double)posRange.Left;
                    top = (double)posRange.Top;
                }
                catch { }
            }

            slicer = cache.Slicers.Add(sheet, Type.Missing, Type.Missing, name ?? field + "Slicer",
                name ?? field, left, top, 200, 300);

            string slicerName = slicer.Name;
            return Response.Ok(new { name = slicerName, field, source, created = true });
        }
        catch (Exception ex)
        {
            return Response.Error($"slicer.create: {ex.Message}");
        }
    }

    private string HandleDelete(JsonObject args)
    {
        try
        {
            var name = args["name"]?.GetValue<string>()
                ?? throw new ArgumentException("slicer.delete requires 'name'");

            using var guard = new ComGuard();
            var wb = guard.Track(_session.GetActiveWorkbook());
            dynamic caches;
            try { caches = ((dynamic)wb).SlicerCaches; }
            catch { return Response.Error("Slicers not available in this Excel version"); }

            int count = caches.Count;
            for (int i = 1; i <= count; i++)
            {
                dynamic cache = caches[i];
                int slicerCount = cache.Slicers.Count;
                for (int j = 1; j <= slicerCount; j++)
                {
                    dynamic slicer = cache.Slicers[j];
                    if (string.Equals((string)slicer.Name, name, StringComparison.OrdinalIgnoreCase))
                    {
                        slicer.Delete();
                        return Response.Ok(new { name, deleted = true });
                    }
                }
            }

            return Response.Error($"Slicer not found: {name}");
        }
        catch (Exception ex)
        {
            return Response.Error($"slicer.delete: {ex.Message}");
        }
    }

    private string HandleSet(JsonObject args)
    {
        try
        {
            var name = args["name"]?.GetValue<string>()
                ?? throw new ArgumentException("slicer.set requires 'name' (slicer cache name)");
            var itemsArr = args["items"]?.AsArray()
                ?? throw new ArgumentException("slicer.set requires 'items' (array of item names to select)");

            var selectedNames = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            foreach (var item in itemsArr)
                if (item != null) selectedNames.Add(item.GetValue<string>());

            using var guard = new ComGuard();
            var wb = guard.Track(_session.GetActiveWorkbook());
            dynamic caches;
            try { caches = ((dynamic)wb).SlicerCaches; }
            catch { return Response.Error("Slicers not available in this Excel version"); }

            dynamic cache = FindCache(caches, name);
            int itemCount = cache.SlicerItems.Count;
            for (int j = 1; j <= itemCount; j++)
            {
                dynamic slicerItem = cache.SlicerItems[j];
                string itemName = slicerItem.Name;
                slicerItem.Selected = selectedNames.Contains(itemName);
            }

            return Response.Ok(new { name, selected = selectedNames.Count, set = true });
        }
        catch (Exception ex)
        {
            return Response.Error($"slicer.set: {ex.Message}");
        }
    }

    private string HandleClear(JsonObject args)
    {
        try
        {
            var name = args["name"]?.GetValue<string>()
                ?? throw new ArgumentException("slicer.clear requires 'name' (slicer cache name)");

            using var guard = new ComGuard();
            var wb = guard.Track(_session.GetActiveWorkbook());
            dynamic caches;
            try { caches = ((dynamic)wb).SlicerCaches; }
            catch { return Response.Error("Slicers not available in this Excel version"); }

            dynamic cache = FindCache(caches, name);
            cache.ClearManualFilter();

            return Response.Ok(new { name, cleared = true });
        }
        catch (Exception ex)
        {
            return Response.Error($"slicer.clear: {ex.Message}");
        }
    }

    private string HandleRead(JsonObject args)
    {
        try
        {
            var name = args["name"]?.GetValue<string>()
                ?? throw new ArgumentException("slicer.read requires 'name' (slicer cache name)");

            using var guard = new ComGuard();
            var wb = guard.Track(_session.GetActiveWorkbook());
            dynamic caches;
            try { caches = ((dynamic)wb).SlicerCaches; }
            catch { return Response.Error("Slicers not available in this Excel version"); }

            dynamic cache = FindCache(caches, name);
            var selected = new JsonArray();
            var unselected = new JsonArray();
            int itemCount = cache.SlicerItems.Count;
            for (int j = 1; j <= itemCount; j++)
            {
                dynamic item = cache.SlicerItems[j];
                string itemName = item.Name;
                if ((bool)item.Selected)
                    selected.Add(itemName);
                else
                    unselected.Add(itemName);
            }

            return Response.Ok(new
            {
                name,
                selected,
                unselected,
                selected_count = selected.Count,
                total_count = itemCount,
            });
        }
        catch (Exception ex)
        {
            return Response.Error($"slicer.read: {ex.Message}");
        }
    }

    // ── Helpers ──────────────────────────────────────────────────────────

    private static dynamic FindCache(dynamic caches, string name)
    {
        int count = caches.Count;
        for (int i = 1; i <= count; i++)
        {
            dynamic cache = caches[i];
            if (string.Equals((string)cache.Name, name, StringComparison.OrdinalIgnoreCase))
                return cache;
        }
        throw new ArgumentException($"Slicer cache not found: {name}");
    }
}
