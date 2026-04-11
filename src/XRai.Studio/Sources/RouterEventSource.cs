using System.Diagnostics;
using System.Text.Json;
using System.Text.Json.Nodes;
using XRai.Core;

namespace XRai.Studio.Sources;

/// <summary>
/// Wraps a <see cref="CommandRouter"/> so every dispatched command emits
/// <c>command.start</c> and <c>command.end</c> events onto the Studio bus.
/// This gives the dashboard a live tail of every JSON command the agent
/// dispatches — the single most valuable signal for the "what is the AI
/// actually doing right now?" question.
///
/// The wrapper is installed by the daemon via a simple decorator pattern:
/// it does NOT modify <see cref="CommandRouter"/>. Instead, the daemon's
/// client handler calls <see cref="WrapDispatch"/> before forwarding.
///
/// Events carry:
///   command.start: { cmd, ts }
///   command.end:   { cmd, ok, elapsedMs, error?, ts }
///
/// This source is agnostic to what kind of target is being driven — the
/// router is just a command dispatcher and the same pattern works for Excel,
/// Word, any desktop app, or even pure filesystem operations.
/// </summary>
public sealed class RouterEventSource : IDisposable
{
    private readonly EventBus _bus;
    private bool _disposed;

    public RouterEventSource(EventBus bus)
    {
        _bus = bus;
    }

    /// <summary>
    /// Decorate a dispatch call: publishes command.start, calls through to
    /// the underlying dispatcher, publishes command.end with elapsed time
    /// and ok/error status.
    /// </summary>
    public string WrapDispatch(string jsonLine, Func<string, string> innerDispatch)
    {
        if (_disposed) return innerDispatch(jsonLine);

        var sw = Stopwatch.StartNew();
        string? cmdName = null;
        try
        {
            // Extract the cmd name without re-parsing the whole thing on the
            // hot path — we only need it for the event payload.
            var node = JsonNode.Parse(jsonLine);
            cmdName = node?["cmd"]?.GetValue<string>();
        }
        catch { }

        if (cmdName != null)
        {
            try
            {
                _bus.Publish(StudioEvent.Now("command.start", "router", new JsonObject
                {
                    ["cmd"] = cmdName,
                }));
            }
            catch { }
        }

        string response;
        bool ok = true;
        string? error = null;
        try
        {
            response = innerDispatch(jsonLine);

            // Parse the response so we can extract ok/error for the end event
            try
            {
                var rnode = JsonNode.Parse(response);
                ok = rnode?["ok"]?.GetValue<bool>() ?? true;
                if (!ok) error = rnode?["error"]?.GetValue<string>();
            }
            catch { /* not a JSON response — assume ok */ }
        }
        catch (Exception ex)
        {
            ok = false;
            error = ex.Message;
            response = $"{{\"ok\":false,\"error\":\"{JsonEncodedText.Encode(ex.Message)}\"}}";
        }

        sw.Stop();

        try
        {
            var endData = new JsonObject
            {
                ["cmd"] = cmdName ?? "?",
                ["ok"] = ok,
                ["elapsedMs"] = sw.ElapsedMilliseconds,
            };
            if (error != null) endData["error"] = error;

            _bus.Publish(StudioEvent.Now("command.end", "router", endData));
        }
        catch { }

        return response;
    }

    public void Dispose() => _disposed = true;
}
