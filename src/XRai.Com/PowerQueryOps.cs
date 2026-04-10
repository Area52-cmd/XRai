using System.Text.Json.Nodes;
using XRai.Core;

namespace XRai.Com;

public class PowerQueryOps
{
    private readonly ExcelSession _session;

    public PowerQueryOps(ExcelSession session) { _session = session; }

    public void Register(CommandRouter router)
    {
        router.Register("powerquery.list", HandleList);
        router.Register("powerquery.view", HandleView);
        router.Register("powerquery.create", HandleCreate);
        router.Register("powerquery.edit", HandleEdit);
        router.Register("powerquery.refresh", HandleRefresh);
        router.Register("powerquery.delete", HandleDelete);
    }

    private string HandleList(JsonObject args)
    {
        try
        {
            using var guard = new ComGuard();
            var wb = guard.Track(_session.GetActiveWorkbook());
            dynamic queries;
            try { queries = ((dynamic)wb).Queries; }
            catch { return Response.Error("Power Query not available in this Excel version", code: ErrorCodes.PowerQueryNotAvailable); }

            var result = new JsonArray();
            int count = queries.Count;
            for (int i = 1; i <= count; i++)
            {
                dynamic q = queries.Item(i);
                string name = q.Name;
                string formula = q.Formula ?? "";
                var preview = formula.Length > 200 ? formula.Substring(0, 200) + "..." : formula;

                string? lastRefresh = null;
                try { lastRefresh = q.RefreshDate?.ToString(); } catch { }

                result.Add(new JsonObject
                {
                    ["name"] = name,
                    ["formula_preview"] = preview,
                    ["last_refresh"] = lastRefresh,
                });
            }

            return Response.Ok(new { count, queries = result });
        }
        catch (Exception ex)
        {
            return Response.Error($"powerquery.list: {ex.Message}");
        }
    }

    private string HandleView(JsonObject args)
    {
        try
        {
            var name = args["name"]?.GetValue<string>()
                ?? throw new ArgumentException("powerquery.view requires 'name'");

            using var guard = new ComGuard();
            var wb = guard.Track(_session.GetActiveWorkbook());
            dynamic queries;
            try { queries = ((dynamic)wb).Queries; }
            catch { return Response.Error("Power Query not available in this Excel version", code: ErrorCodes.PowerQueryNotAvailable); }

            dynamic query = FindQuery(queries, name);
            string formula = query.Formula ?? "";

            return Response.Ok(new { name, formula });
        }
        catch (Exception ex)
        {
            return Response.Error($"powerquery.view: {ex.Message}");
        }
    }

    private string HandleCreate(JsonObject args)
    {
        try
        {
            var name = args["name"]?.GetValue<string>()
                ?? throw new ArgumentException("powerquery.create requires 'name'");
            var formula = args["formula"]?.GetValue<string>()
                ?? throw new ArgumentException("powerquery.create requires 'formula'");

            using var guard = new ComGuard();
            var wb = guard.Track(_session.GetActiveWorkbook());
            dynamic queries;
            try { queries = ((dynamic)wb).Queries; }
            catch { return Response.Error("Power Query not available in this Excel version", code: ErrorCodes.PowerQueryNotAvailable); }

            dynamic newQuery = queries.Add(name, formula);
            string createdName = newQuery.Name;

            return Response.Ok(new { name = createdName, created = true });
        }
        catch (Exception ex)
        {
            return Response.Error($"powerquery.create: {ex.Message}");
        }
    }

    private string HandleEdit(JsonObject args)
    {
        try
        {
            var name = args["name"]?.GetValue<string>()
                ?? throw new ArgumentException("powerquery.edit requires 'name'");
            var formula = args["formula"]?.GetValue<string>()
                ?? throw new ArgumentException("powerquery.edit requires 'formula'");

            using var guard = new ComGuard();
            var wb = guard.Track(_session.GetActiveWorkbook());
            dynamic queries;
            try { queries = ((dynamic)wb).Queries; }
            catch { return Response.Error("Power Query not available in this Excel version", code: ErrorCodes.PowerQueryNotAvailable); }

            dynamic query = FindQuery(queries, name);
            query.Formula = formula;

            return Response.Ok(new { name, updated = true });
        }
        catch (Exception ex)
        {
            return Response.Error($"powerquery.edit: {ex.Message}");
        }
    }

    private string HandleRefresh(JsonObject args)
    {
        try
        {
            var name = args["name"]?.GetValue<string>();

            using var guard = new ComGuard();
            var wb = guard.Track(_session.GetActiveWorkbook());

            if (name == null)
            {
                // Refresh all
                wb.RefreshAll();
                return Response.Ok(new { refreshed_all = true });
            }

            dynamic queries;
            try { queries = ((dynamic)wb).Queries; }
            catch { return Response.Error("Power Query not available in this Excel version", code: ErrorCodes.PowerQueryNotAvailable); }

            // Find the connection that backs this query and refresh it
            // Power Query connections are named "Query - <name>"
            dynamic connections = wb.Connections;
            int connCount = connections.Count;
            for (int i = 1; i <= connCount; i++)
            {
                dynamic conn = connections[i];
                string connName = conn.Name;
                if (connName == $"Query - {name}" || connName == name)
                {
                    conn.Refresh();
                    return Response.Ok(new { name, refreshed = true });
                }
            }

            // Fallback: refresh all
            wb.RefreshAll();
            return Response.Ok(new { name, refreshed = true, note = "Refreshed all connections (specific connection not found)" });
        }
        catch (Exception ex)
        {
            return Response.Error($"powerquery.refresh: {ex.Message}");
        }
    }

    private string HandleDelete(JsonObject args)
    {
        try
        {
            var name = args["name"]?.GetValue<string>()
                ?? throw new ArgumentException("powerquery.delete requires 'name'");

            using var guard = new ComGuard();
            var wb = guard.Track(_session.GetActiveWorkbook());
            dynamic queries;
            try { queries = ((dynamic)wb).Queries; }
            catch { return Response.Error("Power Query not available in this Excel version", code: ErrorCodes.PowerQueryNotAvailable); }

            dynamic query = FindQuery(queries, name);
            query.Delete();

            return Response.Ok(new { name, deleted = true });
        }
        catch (Exception ex)
        {
            return Response.Error($"powerquery.delete: {ex.Message}");
        }
    }

    // ── Helpers ──────────────────────────────────────────────────────────

    private static dynamic FindQuery(dynamic queries, string name)
    {
        int count = queries.Count;
        for (int i = 1; i <= count; i++)
        {
            dynamic q = queries.Item(i);
            if (string.Equals((string)q.Name, name, StringComparison.OrdinalIgnoreCase))
                return q;
        }
        throw new ArgumentException($"Query not found: {name}");
    }
}
