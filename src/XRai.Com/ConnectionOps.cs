using System.Text.Json.Nodes;
using XRai.Core;

namespace XRai.Com;

public class ConnectionOps
{
    private readonly ExcelSession _session;

    public ConnectionOps(ExcelSession session) { _session = session; }

    public void Register(CommandRouter router)
    {
        router.Register("connection.list", HandleList);
        router.Register("connection.refresh", HandleRefresh);
        router.Register("connection.delete", HandleDelete);
    }

    private string HandleList(JsonObject args)
    {
        try
        {
            using var guard = new ComGuard();
            var wb = guard.Track(_session.GetActiveWorkbook());
            dynamic connections = ((dynamic)wb).Connections;
            int count = connections.Count;

            var result = new JsonArray();
            for (int i = 1; i <= count; i++)
            {
                dynamic conn = connections[i];
                string connName = conn.Name;
                string? description = null;
                string? connType = null;
                string? connStringPreview = null;

                try { description = conn.Description?.ToString(); } catch { }
                try { connType = ((int)conn.Type) switch
                {
                    1 => "OLEDB",
                    2 => "ODBC",
                    4 => "ADO",
                    5 => "DataFeed",
                    6 => "Model",
                    7 => "Worksheet",
                    _ => $"Unknown({(int)conn.Type})",
                }; } catch { }
                try
                {
                    string full = conn.OLEDBConnection?.Connection?.ToString()
                        ?? conn.ODBCConnection?.Connection?.ToString()
                        ?? "";
                    connStringPreview = full.Length > 200 ? full.Substring(0, 200) + "..." : full;
                }
                catch { }

                result.Add(new JsonObject
                {
                    ["name"] = connName,
                    ["type"] = connType,
                    ["description"] = description,
                    ["connection_string_preview"] = connStringPreview,
                });
            }

            return Response.Ok(new { count, connections = result });
        }
        catch (Exception ex)
        {
            return Response.Error($"connection.list: {ex.Message}");
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
                wb.RefreshAll();
                return Response.Ok(new { refreshed_all = true });
            }

            dynamic connections = ((dynamic)wb).Connections;
            int count = connections.Count;
            for (int i = 1; i <= count; i++)
            {
                dynamic conn = connections[i];
                if (string.Equals((string)conn.Name, name, StringComparison.OrdinalIgnoreCase))
                {
                    conn.Refresh();
                    return Response.Ok(new { name, refreshed = true });
                }
            }

            return Response.Error($"Connection not found: {name}");
        }
        catch (Exception ex)
        {
            return Response.Error($"connection.refresh: {ex.Message}");
        }
    }

    private string HandleDelete(JsonObject args)
    {
        try
        {
            var name = args["name"]?.GetValue<string>()
                ?? throw new ArgumentException("connection.delete requires 'name'");

            using var guard = new ComGuard();
            var wb = guard.Track(_session.GetActiveWorkbook());
            dynamic connections = ((dynamic)wb).Connections;
            int count = connections.Count;
            for (int i = 1; i <= count; i++)
            {
                dynamic conn = connections[i];
                if (string.Equals((string)conn.Name, name, StringComparison.OrdinalIgnoreCase))
                {
                    conn.Delete();
                    return Response.Ok(new { name, deleted = true });
                }
            }

            return Response.Error($"Connection not found: {name}");
        }
        catch (Exception ex)
        {
            return Response.Error($"connection.delete: {ex.Message}");
        }
    }
}
