using System.Text.Json.Nodes;

namespace XRai.Studio;

/// <summary>
/// A single event that flows through the Studio event bus to the dashboard.
/// Immutable record so it can be safely shared across threads without locking.
///
/// Kind is a dotted string identifying the event type. Examples:
///   "rebuild.step"     — one step of a rebuild completed or started
///   "model.change"     — a ViewModel property changed
///   "control.change"   — a WPF control state changed
///   "pane.exposed"     — Pilot.Expose just walked a new visual tree
///   "model.exposed"    — Pilot.ExposeModel just registered a new ViewModel
///   "file.changed"     — a file in the watched project directory changed
///   "frame"            — a new JPEG screenshot frame is available
///   "command.start"    — a CLI command began executing
///   "command.end"      — a CLI command completed (ok or error)
///   "log"              — an add-in log line was captured
///   "error"            — an add-in exception was captured
/// </summary>
public sealed record StudioEvent(
    long Ts,
    string Kind,
    string Source,
    JsonNode? Data)
{
    public static StudioEvent Now(string kind, string source, JsonNode? data = null)
        => new(DateTimeOffset.UtcNow.ToUnixTimeMilliseconds(), kind, source, data);

    public JsonObject ToJson()
    {
        var obj = new JsonObject
        {
            ["ts"] = Ts,
            ["kind"] = Kind,
            ["source"] = Source,
        };
        if (Data != null) obj["data"] = Data.DeepClone();
        return obj;
    }
}
