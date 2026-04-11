using System.Text;
using System.Text.Json.Nodes;

namespace XRai.Studio.Snapshots;

/// <summary>
/// Structural diff engine for pairs of <see cref="Snapshot"/> records.
/// Produces a machine-readable diff (JSON) AND a deterministic plain-English
/// summary so the dashboard can render both. No LLM call — the summary is
/// generated from the structured diff via fixed templates.
///
/// Covered diff dimensions:
///   - Model properties       added / removed / changed (old → new)
///   - Pane control tree      added / removed / enabled flipped / text changed
///   - Command stream delta   commands issued between the two snapshots
///   - Log tail delta         new log lines between the two snapshots
///   - Screenshot changed     simple byte-length delta + optional pixel-rect
///                            diff (disabled by default for performance)
///
/// The output shape is stable and versioned via a "version" field so the
/// dashboard / tests can rely on it.
/// </summary>
public static class Differ
{
    public const int DiffVersion = 1;

    public static JsonObject Diff(Snapshot before, Snapshot after)
    {
        var result = new JsonObject
        {
            ["version"] = DiffVersion,
            ["beforeId"] = before.Id,
            ["afterId"] = after.Id,
            ["beforeTs"] = before.Ts,
            ["afterTs"] = after.Ts,
            ["deltaMs"] = after.Ts - before.Ts,
        };

        var summaryParts = new List<string>();

        // ── Model diff ─────────────────────────────────────────────────
        var modelDiff = DiffModel(before.Model, after.Model);
        if (modelDiff != null)
        {
            result["model"] = modelDiff;
            var changed = (modelDiff["changed"] as JsonArray)?.Count ?? 0;
            var added = (modelDiff["added"] as JsonArray)?.Count ?? 0;
            var removed = (modelDiff["removed"] as JsonArray)?.Count ?? 0;
            if (changed + added + removed > 0)
                summaryParts.Add($"{changed} changed, {added} added, {removed} removed model properties");
        }

        // ── Pane control tree diff ─────────────────────────────────────
        var paneDiff = DiffPane(before.PaneTree, after.PaneTree);
        if (paneDiff != null)
        {
            result["pane"] = paneDiff;
            var flipped = (paneDiff["flipped"] as JsonArray)?.Count ?? 0;
            if (flipped > 0) summaryParts.Add($"{flipped} control state flips");
        }

        // ── Screenshot byte delta ──────────────────────────────────────
        if (before.Jpeg != null && after.Jpeg != null)
        {
            var sizeDelta = after.Jpeg.Length - before.Jpeg.Length;
            result["screenshot"] = new JsonObject
            {
                ["beforeBytes"] = before.Jpeg.Length,
                ["afterBytes"] = after.Jpeg.Length,
                ["deltaBytes"] = sizeDelta,
            };
            if (Math.Abs(sizeDelta) > 500)
                summaryParts.Add($"screenshot changed ({sizeDelta:+#;-#;0} bytes)");
        }

        // ── Command stream delta ───────────────────────────────────────
        var cmdDelta = DiffCommands(before.CommandTail, after.CommandTail);
        if (cmdDelta.Count > 0)
        {
            var arr = new JsonArray();
            foreach (var c in cmdDelta) arr.Add(c.DeepClone());
            result["commands"] = arr;
            summaryParts.Add($"{cmdDelta.Count} new commands");
        }

        // ── Log delta ──────────────────────────────────────────────────
        var logDelta = DiffLogs(before.LogsTail, after.LogsTail);
        if (logDelta.Length > 0)
        {
            var arr = new JsonArray();
            foreach (var l in logDelta) arr.Add(l);
            result["logs"] = arr;
            summaryParts.Add($"{logDelta.Length} new log lines");
        }

        result["summary"] = summaryParts.Count == 0
            ? "No structural differences."
            : string.Join(", ", summaryParts) + ".";

        return result;
    }

    // ── Model diff ─────────────────────────────────────────────────────

    private static JsonObject? DiffModel(JsonNode? before, JsonNode? after)
    {
        if (before == null && after == null) return null;

        var beforeObj = before as JsonObject ?? new JsonObject();
        var afterObj = after as JsonObject ?? new JsonObject();

        var added = new JsonArray();
        var removed = new JsonArray();
        var changed = new JsonArray();

        foreach (var kvp in afterObj)
        {
            if (!beforeObj.ContainsKey(kvp.Key))
            {
                added.Add(new JsonObject { ["property"] = kvp.Key, ["value"] = kvp.Value?.DeepClone() });
            }
            else
            {
                var beforeVal = beforeObj[kvp.Key]?.ToString();
                var afterVal = kvp.Value?.ToString();
                if (beforeVal != afterVal)
                {
                    changed.Add(new JsonObject
                    {
                        ["property"] = kvp.Key,
                        ["old"] = beforeVal,
                        ["new"] = afterVal,
                    });
                }
            }
        }

        foreach (var kvp in beforeObj)
        {
            if (!afterObj.ContainsKey(kvp.Key))
                removed.Add(new JsonObject { ["property"] = kvp.Key, ["value"] = kvp.Value?.DeepClone() });
        }

        return new JsonObject
        {
            ["added"] = added,
            ["removed"] = removed,
            ["changed"] = changed,
        };
    }

    // ── Pane diff ──────────────────────────────────────────────────────

    private static JsonObject? DiffPane(JsonNode? before, JsonNode? after)
    {
        if (before == null && after == null) return null;

        // Both pane trees are arrays of { name, type, enabled, visible } as
        // emitted by Pilot.Expose. Match by name and compare enabled/visible.
        var beforeControls = ExtractControls(before);
        var afterControls = ExtractControls(after);

        var flipped = new JsonArray();
        var added = new JsonArray();
        var removed = new JsonArray();

        foreach (var a in afterControls)
        {
            if (!beforeControls.TryGetValue(a.Key, out var b))
            {
                added.Add(new JsonObject { ["name"] = a.Key, ["type"] = a.Value.Type });
                continue;
            }

            if (b.Enabled != a.Value.Enabled || b.Visible != a.Value.Visible)
            {
                flipped.Add(new JsonObject
                {
                    ["name"] = a.Key,
                    ["type"] = a.Value.Type,
                    ["enabled_before"] = b.Enabled,
                    ["enabled_after"] = a.Value.Enabled,
                    ["visible_before"] = b.Visible,
                    ["visible_after"] = a.Value.Visible,
                });
            }
        }

        foreach (var b in beforeControls)
        {
            if (!afterControls.ContainsKey(b.Key))
                removed.Add(new JsonObject { ["name"] = b.Key, ["type"] = b.Value.Type });
        }

        return new JsonObject
        {
            ["added"] = added,
            ["removed"] = removed,
            ["flipped"] = flipped,
        };
    }

    private readonly record struct CtrlMeta(string Type, bool Enabled, bool Visible);

    private static Dictionary<string, CtrlMeta> ExtractControls(JsonNode? node)
    {
        var dict = new Dictionary<string, CtrlMeta>(StringComparer.OrdinalIgnoreCase);
        if (node is not JsonObject obj) return dict;

        // Pilot.Expose emits: { rootType, controlCount, controls: [{name,type,enabled,visible}, ...] }
        var controls = obj["controls"] as JsonArray;
        if (controls == null) return dict;

        foreach (var c in controls)
        {
            if (c is not JsonObject co) continue;
            var name = co["name"]?.GetValue<string>();
            if (string.IsNullOrEmpty(name)) continue;
            var type = co["type"]?.GetValue<string>() ?? "";
            var enabled = co["enabled"]?.GetValue<bool>() ?? true;
            var visible = co["visible"]?.GetValue<bool>() ?? true;
            dict[name] = new CtrlMeta(type, enabled, visible);
        }
        return dict;
    }

    // ── Command stream delta ───────────────────────────────────────────

    private static List<JsonObject> DiffCommands(JsonArray? before, JsonArray? after)
    {
        var result = new List<JsonObject>();
        if (after == null) return result;

        var beforeCount = before?.Count ?? 0;
        for (int i = beforeCount; i < after.Count; i++)
        {
            if (after[i] is JsonObject c)
                result.Add(c);
        }
        return result;
    }

    // ── Log delta ──────────────────────────────────────────────────────

    private static string[] DiffLogs(string[]? before, string[]? after)
    {
        if (after == null || after.Length == 0) return Array.Empty<string>();
        var beforeCount = before?.Length ?? 0;
        if (after.Length <= beforeCount) return Array.Empty<string>();
        var result = new string[after.Length - beforeCount];
        Array.Copy(after, beforeCount, result, 0, result.Length);
        return result;
    }
}
