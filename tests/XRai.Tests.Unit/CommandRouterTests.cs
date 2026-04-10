using Xunit;
using System.Text.Json.Nodes;
using XRai.Core;

namespace XRai.Tests.Unit;

public class CommandRouterTests
{
    private CommandRouter CreateRouter()
    {
        var events = new EventStream(TextWriter.Null);
        return new CommandRouter(events);
    }

    [Fact]
    public void Dispatch_UnknownCommand_ReturnsError()
    {
        var router = CreateRouter();
        var result = router.Dispatch("{\"cmd\": \"nonexistent\"}");
        var node = JsonNode.Parse(result)!;
        Assert.False(node["ok"]!.GetValue<bool>());
        Assert.Contains("Unknown command", node["error"]!.GetValue<string>());
    }

    [Fact]
    public void Dispatch_RegisteredCommand_CallsHandler()
    {
        var router = CreateRouter();
        router.Register("test", args => Response.Ok(new { handled = true }));

        var result = router.Dispatch("{\"cmd\": \"test\"}");
        var node = JsonNode.Parse(result)!;
        Assert.True(node["ok"]!.GetValue<bool>());
        Assert.True(node["handled"]!.GetValue<bool>());
    }

    [Fact]
    public void Dispatch_MissingCmd_ReturnsError()
    {
        var router = CreateRouter();
        var result = router.Dispatch("{\"foo\": \"bar\"}");
        var node = JsonNode.Parse(result)!;
        Assert.False(node["ok"]!.GetValue<bool>());
    }

    [Fact]
    public void Dispatch_InvalidJson_ReturnsError()
    {
        var router = CreateRouter();
        var result = router.Dispatch("not json at all");
        var node = JsonNode.Parse(result)!;
        Assert.False(node["ok"]!.GetValue<bool>());
    }

    [Fact]
    public void Dispatch_Batch_ExecutesAll()
    {
        var router = CreateRouter();
        int callCount = 0;
        router.Register("inc", args =>
        {
            callCount++;
            return Response.Ok(new { count = callCount });
        });

        var result = router.Dispatch("{\"cmd\": \"batch\", \"commands\": [{\"cmd\": \"inc\"}, {\"cmd\": \"inc\"}, {\"cmd\": \"inc\"}]}");
        var node = JsonNode.Parse(result)!;
        Assert.True(node["ok"]!.GetValue<bool>());
        Assert.Equal(3, callCount);
        Assert.Equal(3, node["results"]!.AsArray().Count);
    }

    [Fact]
    public void Dispatch_HandlerException_ReturnsError()
    {
        var router = CreateRouter();
        router.Register("boom", _ => throw new InvalidOperationException("kaboom"));

        var result = router.Dispatch("{\"cmd\": \"boom\"}");
        var node = JsonNode.Parse(result)!;
        Assert.False(node["ok"]!.GetValue<bool>());
        Assert.Contains("kaboom", node["error"]!.GetValue<string>());
    }
}
