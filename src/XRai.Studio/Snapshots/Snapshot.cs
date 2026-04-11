using System.Text.Json.Nodes;

namespace XRai.Studio.Snapshots;

/// <summary>
/// A point-in-time capture of the observed target application state.
/// Used by the time-travel scrubber: the user can slide back through
/// snapshots and the dashboard panels all snap to the historic values.
///
/// Capture triggers (event → snapshot):
///   - command.end      after every dispatched command
///   - rebuild.step     after every build step (ok or error)
///   - pane.exposed     a new visual tree was walked
///   - model.exposed    a new ViewModel was registered
///   - file.changed     a watched source file was edited
///
/// Storage: a circular ring buffer holds the most recent N snapshots
/// (default 500). Optional disk spill when recording is enabled.
/// </summary>
public sealed class Snapshot
{
    /// <summary>Monotonic per-daemon id. Starts at 1.</summary>
    public required long Id { get; init; }

    /// <summary>Unix ms when the snapshot was captured.</summary>
    public required long Ts { get; init; }

    /// <summary>What triggered the capture — maps 1:1 to a StudioEvent kind.</summary>
    public required string Cause { get; init; }

    /// <summary>
    /// Latest JPEG frame bytes from CaptureLoop at the moment of capture,
    /// or null if no frame had arrived yet.
    /// </summary>
    public byte[]? Jpeg { get; init; }

    /// <summary>Control tree as of capture — null if no pane was exposed yet.</summary>
    public JsonNode? PaneTree { get; init; }

    /// <summary>Flat ViewModel property dictionary as of capture.</summary>
    public JsonNode? Model { get; init; }

    /// <summary>Tail of the pilot log at capture time, up to ~200 lines.</summary>
    public string[]? LogsTail { get; init; }

    /// <summary>
    /// Optional user-configured cell watchlist values. Future extension —
    /// Studio gains a {"cmd":"studio.watch.cells","range":"A1:D10"} command
    /// that adds cells to the watchlist; snapshots record their values.
    /// </summary>
    public JsonNode? WatchedCells { get; init; }

    /// <summary>
    /// Command stream tail — last N command.start/command.end pairs relative
    /// to this snapshot. Lets the dashboard show "last 5 commands" per frame.
    /// </summary>
    public JsonArray? CommandTail { get; init; }

    /// <summary>
    /// Return a lightweight metadata-only view suitable for the /snapshots
    /// listing endpoint. Excludes the JPEG bytes so the listing stays small.
    /// </summary>
    public JsonObject ToMetadata()
    {
        var obj = new JsonObject
        {
            ["id"] = Id,
            ["ts"] = Ts,
            ["cause"] = Cause,
            ["hasJpeg"] = Jpeg != null,
            ["jpegBytes"] = Jpeg?.Length ?? 0,
            ["hasPane"] = PaneTree != null,
            ["hasModel"] = Model != null,
            ["logLineCount"] = LogsTail?.Length ?? 0,
            ["commandCount"] = CommandTail?.Count ?? 0,
        };
        return obj;
    }

    /// <summary>
    /// Return the full snapshot as JSON, inlining the JPEG as base64. Used
    /// by the /snapshots/{id} endpoint when the dashboard scrubs to a
    /// specific point in time.
    /// </summary>
    public JsonObject ToJson()
    {
        var obj = ToMetadata();
        if (Jpeg != null)
        {
            obj["jpegB64"] = Convert.ToBase64String(Jpeg);
            obj["jpegMime"] = "image/jpeg";
        }
        if (PaneTree != null) obj["pane"] = PaneTree.DeepClone();
        if (Model != null) obj["model"] = Model.DeepClone();
        if (LogsTail != null)
        {
            var logs = new JsonArray();
            foreach (var line in LogsTail) logs.Add(line);
            obj["logs"] = logs;
        }
        if (WatchedCells != null) obj["cells"] = WatchedCells.DeepClone();
        if (CommandTail != null) obj["commands"] = CommandTail.DeepClone();
        return obj;
    }
}
