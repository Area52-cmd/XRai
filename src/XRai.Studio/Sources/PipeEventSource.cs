using System.Text.Json;
using System.Text.Json.Nodes;

namespace XRai.Studio.Sources;

/// <summary>
/// Subscribes to the static events on XRai.Hooks types via reflection and
/// forwards them onto the Studio EventBus. Uses reflection so XRai.Studio
/// does NOT take a build-time dependency on XRai.Hooks (which is a WPF
/// assembly that needs a UI-thread context). The Hooks assembly loads
/// inside the Excel add-in process; Studio runs inside the daemon process.
///
/// Wait — the daemon is a separate process from the add-in. The static
/// events we added live in the add-in process. How does Studio see them?
///
/// Two paths:
///   1. In-process: if Studio were running INSIDE the add-in (future:
///      Pilot.StartStudio()), the static-event route would work directly.
///   2. Out-of-process (current): the daemon's HookConnection receives
///      events over the named pipe. We hook into THAT read loop instead.
///
/// This class implements path 2 — it's not actually reflection into
/// XRai.Hooks. It wires into HookConnection's pipe-line reader so the same
/// events that currently get logged via PushEvent show up on the Studio bus.
/// </summary>
public sealed class PipeEventSource : IDisposable
{
    private readonly EventBus _bus;
    private bool _disposed;

    public PipeEventSource(EventBus bus)
    {
        _bus = bus;
    }

    /// <summary>
    /// Called by HookConnection (or a wrapper) whenever an event JSON line
    /// arrives from the add-in's pipe server. The line is a JsonObject with
    /// "event" and optional other fields.
    /// </summary>
    public void OnPipeEvent(string rawLine)
    {
        if (_disposed) return;
        if (string.IsNullOrWhiteSpace(rawLine)) return;

        try
        {
            var node = JsonNode.Parse(rawLine);
            if (node is not JsonObject obj) return;
            var eventType = obj["event"]?.GetValue<string>();
            if (string.IsNullOrEmpty(eventType)) return;

            // Everything under the "event" envelope that isn't the type itself
            // becomes the data payload.
            var data = new JsonObject();
            foreach (var kvp in obj)
            {
                if (kvp.Key == "event") continue;
                data[kvp.Key] = kvp.Value?.DeepClone();
            }

            _bus.Publish(StudioEvent.Now(eventType, "hooks", data));
        }
        catch
        {
            // Malformed event lines are silently dropped.
        }
    }

    /// <summary>
    /// Convenience: directly publish a payload as if it came from the pipe.
    /// Used when Studio runs in the add-in process itself and wires into
    /// PipeServer.OnEventEmitted (no pipe round-trip needed).
    /// </summary>
    public void OnInProcessEvent(string eventType, object? data)
    {
        if (_disposed || string.IsNullOrEmpty(eventType)) return;
        try
        {
            JsonNode? node = null;
            if (data != null)
            {
                var json = JsonSerializer.Serialize(data);
                node = JsonNode.Parse(json);
            }
            _bus.Publish(StudioEvent.Now(eventType, "hooks", node));
        }
        catch { }
    }

    public void Dispose() => _disposed = true;
}
