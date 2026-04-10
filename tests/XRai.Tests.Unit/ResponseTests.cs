using Xunit;
using System.Text.Json.Nodes;
using XRai.Core;

namespace XRai.Tests.Unit;

public class ResponseTests
{
    [Fact]
    public void Ok_ReturnsOkTrue()
    {
        var json = Response.Ok();
        var node = JsonNode.Parse(json)!;
        Assert.True(node["ok"]!.GetValue<bool>());
    }

    [Fact]
    public void Ok_WithData_MergesFields()
    {
        var json = Response.Ok(new { value = 42, name = "test" });
        var node = JsonNode.Parse(json)!;
        Assert.True(node["ok"]!.GetValue<bool>());
        Assert.Equal(42, node["value"]!.GetValue<int>());
        Assert.Equal("test", node["name"]!.GetValue<string>());
    }

    [Fact]
    public void Error_ReturnsOkFalse()
    {
        var json = Response.Error("something broke");
        var node = JsonNode.Parse(json)!;
        Assert.False(node["ok"]!.GetValue<bool>());
        Assert.Equal("something broke", node["error"]!.GetValue<string>());
    }

    [Fact]
    public void Event_HasEventField()
    {
        var json = Response.Event("log", new { message = "hello" });
        var node = JsonNode.Parse(json)!;
        Assert.Equal("log", node["event"]!.GetValue<string>());
        Assert.Equal("hello", node["message"]!.GetValue<string>());
    }
}
