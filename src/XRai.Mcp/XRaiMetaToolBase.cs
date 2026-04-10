using System.Text.Json.Nodes;
using XRai.Core;

namespace XRai.Mcp;

/// <summary>
/// Base class for all XRai meta-tools. Provides a Dispatch helper that builds
/// the JSON command envelope and routes it through the CommandRouter.
/// </summary>
public abstract class XRaiMetaToolBase
{
    protected readonly CommandRouter Router;

    protected XRaiMetaToolBase(CommandRouter router) => Router = router;

    protected string Dispatch(string command, string? argsJson = null)
    {
        var obj = new JsonObject { ["cmd"] = command };
        if (!string.IsNullOrEmpty(argsJson))
        {
            try
            {
                var parsed = JsonNode.Parse(argsJson);
                if (parsed is JsonObject extra)
                    foreach (var kvp in extra)
                        obj[kvp.Key] = kvp.Value?.DeepClone();
            }
            catch
            {
                // If args is not valid JSON, ignore and send command without extra args
            }
        }
        return Router.Dispatch(obj.ToJsonString());
    }
}
