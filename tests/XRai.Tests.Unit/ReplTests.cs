using Xunit;
using System.Text.Json.Nodes;
using XRai.Core;

namespace XRai.Tests.Unit;

public class ReplTests
{
    [Fact]
    public void Repl_ProcessesCommandsFromReader()
    {
        var input = new StringReader("{\"cmd\": \"echo\"}\n{\"cmd\": \"echo\"}\n");
        var output = new StringWriter();
        var events = new EventStream(output);
        var router = new CommandRouter(events);
        router.Register("echo", _ => Response.Ok(new { echo = true }));

        var repl = new Repl(router, events, input);
        repl.Run();

        var lines = output.ToString().Split('\n', StringSplitOptions.RemoveEmptyEntries);
        Assert.Equal(2, lines.Length);
        foreach (var line in lines)
        {
            var node = JsonNode.Parse(line)!;
            Assert.True(node["ok"]!.GetValue<bool>());
        }
    }

    [Fact]
    public void Repl_SkipsEmptyLines()
    {
        var input = new StringReader("\n\n{\"cmd\": \"test\"}\n\n");
        var output = new StringWriter();
        var events = new EventStream(output);
        var router = new CommandRouter(events);
        router.Register("test", _ => Response.Ok());

        var repl = new Repl(router, events, input);
        repl.Run();

        var lines = output.ToString().Split('\n', StringSplitOptions.RemoveEmptyEntries);
        Assert.Single(lines);
    }

    [Fact]
    public void Repl_AcceptsMultiLinePrettyPrintedJson()
    {
        // A pretty-printed JSON object spanning multiple lines MUST be accepted
        // as a single command. This was issue 1g from CellVault testing — agents
        // write pretty JSON for legibility and the strict line-parser was rejecting it.
        var input = new StringReader(@"{
  ""cmd"": ""batch"",
  ""commands"": [
    {""cmd"": ""ping""},
    {""cmd"": ""ping""}
  ]
}");
        var output = new StringWriter();
        var events = new EventStream(output);
        var router = new CommandRouter(events);
        router.Register("ping", _ => Response.Ok(new { pong = true }));

        var repl = new Repl(router, events, input);
        repl.Run();

        var lines = output.ToString().Split('\n', StringSplitOptions.RemoveEmptyEntries);
        Assert.Single(lines); // one batch → one response
        var node = JsonNode.Parse(lines[0])!;
        Assert.True(node["ok"]!.GetValue<bool>());
        Assert.Equal(2, node["results"]!.AsArray().Count);
    }

    [Fact]
    public void Repl_AcceptsStringWithBracesAndQuotes()
    {
        // Braces and brackets INSIDE string values must not be counted as structural
        var input = new StringReader(@"{""cmd"":""echo"",""value"":""a {fake} [bracket] \""quoted\"" string""}");
        var output = new StringWriter();
        var events = new EventStream(output);
        var router = new CommandRouter(events);
        router.Register("echo", _ => Response.Ok());

        var repl = new Repl(router, events, input);
        repl.Run();

        var lines = output.ToString().Split('\n', StringSplitOptions.RemoveEmptyEntries);
        Assert.Single(lines);
    }

    [Fact]
    public void Repl_HandlesMultipleJsonDocumentsInStream()
    {
        var input = new StringReader(@"
{""cmd"": ""a""}

{
  ""cmd"": ""b""
}
{""cmd"": ""c""}
");
        var output = new StringWriter();
        var events = new EventStream(output);
        var router = new CommandRouter(events);
        router.Register("a", _ => Response.Ok(new { name = "a" }));
        router.Register("b", _ => Response.Ok(new { name = "b" }));
        router.Register("c", _ => Response.Ok(new { name = "c" }));

        var repl = new Repl(router, events, input);
        repl.Run();

        var lines = output.ToString().Split('\n', StringSplitOptions.RemoveEmptyEntries);
        Assert.Equal(3, lines.Length);
    }
}
