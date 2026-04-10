using System.Text.Json.Nodes;
using XRai.Core;

namespace XRai.HooksClient;

public class ModelClient
{
    private readonly HookConnection _connection;

    public ModelClient(HookConnection connection)
    {
        _connection = connection;
    }

    public void Register(CommandRouter router)
    {
        router.RegisterNoSta("model", HandleModel);
        router.RegisterNoSta("model.set", HandleModelSet);
        router.RegisterNoSta("functions", HandleFunctions);
    }

    private string HandleModel(JsonObject args)
    {
        // Optional 'name' targets a specific keyed ViewModel exposed via
        // Pilot.ExposeModel(vm, "SomeName"). Without it, the unkeyed default
        // model (last-exposed) is returned.
        var name = args["name"]?.GetValue<string>();
        var payload = name != null
            ? (object)new { name }
            : new { };
        var resp = _connection.SendCommand("model", payload);
        if (resp?["ok"]?.GetValue<bool>() != true)
            return Response.Error(resp?["error"]?.GetValue<string>() ?? "Failed to get model");

        return resp!.ToJsonString();
    }

    private string HandleModelSet(JsonObject args)
    {
        var property = args["property"]?.GetValue<string>()
            ?? throw new ArgumentException("model.set requires 'property'");

        // Pass the value as-is (could be string, number, bool)
        var value = args["value"];
        var name = args["name"]?.GetValue<string>();

        var resp = name != null
            ? _connection.SendCommand("model_set", new { property, value, name })
            : _connection.SendCommand("model_set", new { property, value });
        if (resp?["ok"]?.GetValue<bool>() != true)
            return Response.Error(resp?["error"]?.GetValue<string>() ?? "Failed to set property");

        return resp!.ToJsonString();
    }

    private string HandleFunctions(JsonObject args)
    {
        var resp = _connection.SendCommand("functions");
        if (resp?["ok"]?.GetValue<bool>() != true)
            return Response.Error(resp?["error"]?.GetValue<string>() ?? "Failed to get functions");

        return resp!.ToJsonString();
    }
}
