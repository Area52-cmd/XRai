using System.Reflection;
using System.Text.Json.Nodes;
using XRai.Core;

namespace XRai.HooksClient;

public class PaneClient
{
    private readonly HookConnection _connection;

    // Build marker baked into whichever assembly owns Program.Main (XRai.Tool).
    // Surfaced in pane.click responses as "command_dispatcher_build" so the
    // consumer can verify end-to-end which CLI binary actually handled a
    // command — catches the "Hooks bumped but Tool is stale" failure mode
    // where a mismatched Tool/Hooks pair produces mysterious behavior.
    private static readonly string ToolBuildStamp = ResolveToolBuildStamp();

    private static string ResolveToolBuildStamp()
    {
        try
        {
            var entry = Assembly.GetEntryAssembly();
            var meta = entry?.GetCustomAttributes<AssemblyMetadataAttribute>()
                .FirstOrDefault(a => a.Key == "XRaiToolBuildTimestamp");
            if (meta?.Value != null) return meta.Value;

            var fv = entry?.GetCustomAttribute<AssemblyFileVersionAttribute>()?.Version;
            return fv ?? "unknown";
        }
        catch { return "unknown"; }
    }

    public PaneClient(HookConnection connection)
    {
        _connection = connection;
    }

    public void Register(CommandRouter router)
    {
        // Basic control interaction
        router.Register("pane", HandlePane);
        router.Register("pane.status", HandlePaneStatus);
        router.Register("pane.type", HandlePaneType);
        router.Register("pane.click", HandlePaneClick);
        router.Register("pane.select", HandlePaneSelect);
        router.Register("pane.toggle", HandlePaneToggle);
        router.Register("pane.read", HandlePaneRead);

        // Human simulation
        router.Register("pane.double_click", HandlePaneDoubleClick);
        router.Register("pane.right_click", HandlePaneRightClick);
        router.Register("pane.hover", HandlePaneHover);
        router.Register("pane.focus", HandlePaneFocus);
        router.Register("pane.key", HandlePaneKey);
        router.Register("pane.scroll", HandlePaneScroll);
        router.Register("pane.info", HandlePaneInfo);
        router.Register("pane.tree", HandlePaneTree);

        // DataGrid operations
        router.Register("pane.grid.read", HandleGridRead);
        router.Register("pane.grid.cell", HandleGridCell);
        router.Register("pane.grid.select", HandleGridSelectRow);

        // TreeView operations
        router.Register("pane.tree.expand", HandleTreeExpand);

        // TabControl operations
        router.Register("pane.tab", HandleTabSelect);

        // ListBox/ListView/ComboBox read items + selection by index or text
        router.Register("pane.list.read", HandleListRead);
        router.Register("pane.list.select", HandleListSelect);

        // Open/close ComboBox dropdown, Expander, TreeViewItem, MenuItem submenu
        router.Register("pane.expand", HandleExpand);

        // AI-agent interaction commands
        router.Register("pane.wait", HandlePaneWait);
        router.Register("pane.screenshot", HandlePaneScreenshot);
        router.Register("pane.drag", HandlePaneDrag);
        router.Register("pane.context_menu", HandlePaneContextMenu);

        // Diagnostics
        router.Register("log.read", HandleLogRead);
    }

    /// <summary>
    /// Read recent lines from an XRai log. Reads directly from disk — does NOT
    /// require an active hooks pipe, because log.read is the primary diagnostic
    /// for "why are hooks down?" and must work when hooks are unreachable.
    ///
    /// File locations (under %LOCALAPPDATA%\XRai\logs):
    ///   - pilot-{pid}.log   — written by Pilot.Log in the add-in process
    ///   - daemon.log        — written by the XRai daemon
    ///   - startup logs      — %TEMP%\*-startup.log written by scaffolded
    ///                         AutoOpen() before Pilot.Start() runs
    ///
    /// Args:
    ///   source ("pilot" | "daemon" | "startup", default "pilot")
    ///   lines  (int, default 100)
    ///   pid    (optional int; when source="pilot", picks pilot-{pid}.log
    ///           instead of the most-recent one)
    /// </summary>
    private string HandleLogRead(JsonObject args)
    {
        var source = args["source"]?.GetValue<string>() ?? "pilot";
        var lines = args["lines"]?.GetValue<int>() ?? 100;

        try
        {
            var logsDir = System.IO.Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                "XRai", "logs");

            string? path = null;
            string[]? candidates = null;

            if (source == "daemon")
            {
                path = System.IO.Path.Combine(logsDir, "daemon.log");
            }
            else if (source == "startup")
            {
                // Scaffolded AutoOpen() writes to %TEMP%\{name}-startup.log
                var tempDir = System.IO.Path.GetTempPath();
                candidates = System.IO.Directory.Exists(tempDir)
                    ? System.IO.Directory.GetFiles(tempDir, "*-startup.log")
                    : Array.Empty<string>();
                if (candidates.Length == 0)
                    return Response.Ok(new
                    {
                        source,
                        search_dir = tempDir,
                        lines = Array.Empty<string>(),
                        exists = false,
                        hint = "No *-startup.log files found in %TEMP%. Scaffolded add-ins write to this file only after AutoOpen() runs — make sure the .xll has been loaded at least once.",
                    });

                // Pick newest by last-write-time
                path = candidates.OrderByDescending(p => new System.IO.FileInfo(p).LastWriteTimeUtc).First();
            }
            else if (source == "pilot")
            {
                // Direct file read from disk — do NOT round-trip through hooks.
                // log.read MUST work when hooks are broken, because that's
                // exactly when the user is trying to diagnose why.
                var pid = args["pid"]?.GetValue<int?>();

                if (!System.IO.Directory.Exists(logsDir))
                    return Response.Ok(new
                    {
                        source,
                        search_dir = logsDir,
                        lines = Array.Empty<string>(),
                        exists = false,
                        hint = "No XRai logs directory. The add-in hasn't called Pilot.Log() yet, or Pilot.Start() never ran.",
                    });

                candidates = System.IO.Directory.GetFiles(logsDir, "pilot-*.log");
                if (candidates.Length == 0)
                    return Response.Ok(new
                    {
                        source,
                        search_dir = logsDir,
                        lines = Array.Empty<string>(),
                        exists = false,
                        hint = "No pilot-*.log files. Pilot.Start() hasn't run yet — check 'log.read' with source='startup' for scaffold-level startup failures.",
                    });

                if (pid.HasValue)
                {
                    path = System.IO.Path.Combine(logsDir, $"pilot-{pid}.log");
                    if (!System.IO.File.Exists(path))
                        return Response.ErrorWithData(
                            $"pilot-{pid}.log not found. Available pids: " +
                            string.Join(", ", candidates.Select(p => System.IO.Path.GetFileNameWithoutExtension(p).Replace("pilot-", ""))),
                            null,
                            "XRAI_INVALID_ARGUMENT");
                }
                else
                {
                    // Pick newest
                    path = candidates.OrderByDescending(p => new System.IO.FileInfo(p).LastWriteTimeUtc).First();
                }
            }
            else
            {
                return Response.ErrorWithData(
                    $"Unknown log source '{source}'",
                    new { valid = new[] { "pilot", "daemon", "startup" } },
                    "XRAI_INVALID_ARGUMENT");
            }

            if (path == null || !System.IO.File.Exists(path))
                return Response.Ok(new { source, path, lines = Array.Empty<string>(), exists = false });

            var allLines = System.IO.File.ReadAllLines(path);
            var tail = allLines.Length > lines ? allLines[^lines..] : allLines;
            return Response.Ok(new
            {
                source,
                path,
                lines = tail,
                total_lines = allLines.Length,
                exists = true,
                candidates = candidates?.Length ?? 1,
            });
        }
        catch (Exception ex)
        {
            return Response.ErrorFromException(ex, "log.read");
        }
    }

    private string Forward(string cmd, object? data = null)
    {
        JsonNode? resp;
        try
        {
            resp = _connection.SendCommand(cmd, data);
        }
        catch (Exception ex)
        {
            return Response.ErrorFromException(ex, $"pane.{cmd}");
        }

        if (resp == null)
            return Response.Error($"pane.{cmd} failed: no response from hooks pipe (connection may have died)");

        // Pass the hooks response through verbatim — it may contain detailed
        // diagnostic fields (exception_type, stack_frame, resolved_target_type,
        // command_can_execute, method_attempted, etc.) that the caller needs
        // for debugging. Previously this stripped everything except the error
        // message, leaving callers flying blind.
        return resp.ToJsonString();
    }

    private string HandlePane(JsonObject args) => Forward("pane_tree");

    private string HandlePaneStatus(JsonObject args)
    {
        if (!_connection.IsConnected)
        {
            return Response.Ok(new
            {
                pipe_connected = false,
                exposed_controls = 0,
                exposed_models = 0,
                hint = "Hooks pipe not connected. The add-in is not loaded, or Pilot.Start() was not called in AutoOpen()."
            });
        }
        return Forward("pane_status");
    }

    private string HandlePaneType(JsonObject args)
    {
        var control = args["control"]?.GetValue<string>() ?? throw new ArgumentException("requires 'control'");
        var value = args["value"]?.GetValue<string>() ?? throw new ArgumentException("requires 'value'");
        var timeout = args["timeout"]?.GetValue<int>() ?? 0;
        return Forward("set_control", new { name = control, value, timeout });
    }

    private string HandlePaneClick(JsonObject args)
    {
        var control = args["control"]?.GetValue<string>() ?? throw new ArgumentException("requires 'control'");
        // Optional auto-wait: poll for the named control to appear before
        // dispatching the click. Eliminates the ribbon.click → pane.click race
        // where the pane's WPF Loaded event fires asynchronously after the
        // ribbon button returns. Default 0 = no wait, preserving prior behavior.
        var timeout = args["timeout"]?.GetValue<int>() ?? 0;

        // Single dispatch to the Hooks pipe. No retry, no post-call verification,
        // no polling at the click step. The Hooks side invokes ButtonBase.OnClick
        // exactly once and returns synchronously.
        var forwarded = Forward("click", new { name = control, timeout });

        // Inject the Tool build stamp into the response so downstream can
        // confirm which CLI binary handled the call. Preserves the existing
        // Hooks response shape (method, command_executed, etc.) untouched.
        try
        {
            var node = JsonNode.Parse(forwarded);
            if (node is JsonObject obj)
            {
                obj["command_dispatcher_build"] = ToolBuildStamp;
                return obj.ToJsonString();
            }
        }
        catch { /* fall through with unmodified response */ }

        return forwarded;
    }

    private string HandlePaneSelect(JsonObject args)
    {
        var control = args["control"]?.GetValue<string>() ?? throw new ArgumentException("requires 'control'");
        var value = args["value"]?.GetValue<string>() ?? throw new ArgumentException("requires 'value'");
        var timeout = args["timeout"]?.GetValue<int>() ?? 0;
        return Forward("select_control", new { name = control, value, timeout });
    }

    private string HandlePaneToggle(JsonObject args)
    {
        var control = args["control"]?.GetValue<string>() ?? throw new ArgumentException("requires 'control'");
        var timeout = args["timeout"]?.GetValue<int>() ?? 0;
        return Forward("toggle_control", new { name = control, timeout });
    }

    private string HandlePaneRead(JsonObject args)
    {
        var control = args["control"]?.GetValue<string>() ?? throw new ArgumentException("requires 'control'");
        var timeout = args["timeout"]?.GetValue<int>() ?? 0;
        return Forward("read_control", new { name = control, timeout });
    }

    private string HandlePaneDoubleClick(JsonObject args)
    {
        var control = args["control"]?.GetValue<string>() ?? throw new ArgumentException("requires 'control'");
        return Forward("double_click", new { name = control });
    }

    private string HandlePaneRightClick(JsonObject args)
    {
        var control = args["control"]?.GetValue<string>() ?? throw new ArgumentException("requires 'control'");
        return Forward("right_click", new { name = control });
    }

    private string HandlePaneHover(JsonObject args)
    {
        var control = args["control"]?.GetValue<string>() ?? throw new ArgumentException("requires 'control'");
        return Forward("hover", new { name = control });
    }

    private string HandlePaneFocus(JsonObject args)
    {
        var control = args["control"]?.GetValue<string>() ?? throw new ArgumentException("requires 'control'");
        return Forward("focus", new { name = control });
    }

    private string HandlePaneKey(JsonObject args)
    {
        var control = args["control"]?.GetValue<string>() ?? throw new ArgumentException("requires 'control'");
        // Accept both 'keys' (canonical) and 'key' (common typo) as aliases
        var keys = args["keys"]?.GetValue<string>()
                   ?? args["key"]?.GetValue<string>()
                   ?? throw new ArgumentException("requires 'keys' (or 'key' as alias) — e.g. \"Enter\", \"Escape\", \"Control+A\"");
        return Forward("send_keys", new { name = control, keys });
    }

    private string HandlePaneScroll(JsonObject args)
    {
        var control = args["control"]?.GetValue<string>() ?? throw new ArgumentException("requires 'control'");
        var offset = args["offset"]?.GetValue<double>() ?? 0;
        return Forward("scroll", new { name = control, offset });
    }

    private string HandlePaneInfo(JsonObject args)
    {
        var control = args["control"]?.GetValue<string>() ?? throw new ArgumentException("requires 'control'");
        return Forward("control_info", new { name = control });
    }

    private string HandlePaneTree(JsonObject args) => Forward("pane_tree");

    private string HandleGridRead(JsonObject args)
    {
        var control = args["control"]?.GetValue<string>() ?? throw new ArgumentException("requires 'control'");
        var timeout = args["timeout"]?.GetValue<int>() ?? 0;
        return Forward("datagrid_read", new { name = control, timeout });
    }

    private string HandleGridCell(JsonObject args)
    {
        var control = args["control"]?.GetValue<string>() ?? throw new ArgumentException("requires 'control'");
        var row = args["row"]?.GetValue<int>() ?? throw new ArgumentException("requires 'row'");
        var col = args["col"]?.GetValue<int>() ?? throw new ArgumentException("requires 'col'");
        return Forward("datagrid_cell", new { name = control, row, col });
    }

    private string HandleGridSelectRow(JsonObject args)
    {
        var control = args["control"]?.GetValue<string>() ?? throw new ArgumentException("requires 'control'");
        var row = args["row"]?.GetValue<int>() ?? throw new ArgumentException("requires 'row'");
        return Forward("datagrid_select", new { name = control, row });
    }

    private string HandleTreeExpand(JsonObject args)
    {
        var control = args["control"]?.GetValue<string>() ?? throw new ArgumentException("requires 'control'");
        var path = args["path"]?.GetValue<string>() ?? throw new ArgumentException("requires 'path'");
        return Forward("tree_expand", new { name = control, path });
    }

    private string HandleTabSelect(JsonObject args)
    {
        var control = args["control"]?.GetValue<string>() ?? throw new ArgumentException("requires 'control'");
        var tab = args["tab"]?.GetValue<string>() ?? throw new ArgumentException("requires 'tab'");
        return Forward("tab_select", new { name = control, tab });
    }

    private string HandleListRead(JsonObject args)
    {
        var control = args["control"]?.GetValue<string>() ?? throw new ArgumentException("requires 'control'");
        var timeout = args["timeout"]?.GetValue<int>() ?? 0;
        return Forward("list_read", new { name = control, timeout });
    }

    private string HandleListSelect(JsonObject args)
    {
        var control = args["control"]?.GetValue<string>() ?? throw new ArgumentException("requires 'control'");
        var index = args["index"]?.GetValue<int>();
        var text = args["text"]?.GetValue<string>();
        var timeout = args["timeout"]?.GetValue<int>() ?? 0;
        if (index == null && text == null)
            throw new ArgumentException("pane.list.select requires 'index' (int) or 'text' (string)");
        return Forward("list_select", new { name = control, index, text, timeout });
    }

    private string HandleExpand(JsonObject args)
    {
        var control = args["control"]?.GetValue<string>() ?? throw new ArgumentException("requires 'control'");
        var open = args["open"]?.GetValue<bool>() ?? true;
        return Forward("expand", new { name = control, open });
    }

    private string HandlePaneWait(JsonObject args)
    {
        var control = args["control"]?.GetValue<string>() ?? throw new ArgumentException("requires 'control'");
        var value = args["value"]?.GetValue<string>();
        bool? enabled = args["enabled"] is JsonNode en ? en.GetValue<bool>() : null;
        bool? exists = args["exists"] is JsonNode ex ? ex.GetValue<bool>() : null;
        var timeout = args["timeout"]?.GetValue<int>();
        var pollMs = args["poll_ms"]?.GetValue<int>();

        var data = new Dictionary<string, object?> { ["name"] = control };
        if (value != null) data["value"] = value;
        if (enabled.HasValue) data["enabled"] = enabled.Value;
        if (exists.HasValue) data["exists"] = exists.Value;
        if (timeout.HasValue) data["timeout"] = timeout.Value;
        if (pollMs.HasValue) data["poll_ms"] = pollMs.Value;

        return Forward("wait_control", data);
    }

    private string HandlePaneScreenshot(JsonObject args)
    {
        return Forward("pane_screenshot");
    }

    private string HandlePaneDrag(JsonObject args)
    {
        var from = args["from"]?.GetValue<string>() ?? throw new ArgumentException("requires 'from'");
        var to = args["to"]?.GetValue<string>() ?? throw new ArgumentException("requires 'to'");
        return Forward("drag", new { from, to });
    }

    private string HandlePaneContextMenu(JsonObject args)
    {
        var control = args["control"]?.GetValue<string>() ?? throw new ArgumentException("requires 'control'");
        var action = args["action"]?.GetValue<string>() ?? throw new ArgumentException("requires 'action'");
        var item = args["item"]?.GetValue<string>();

        var data = new Dictionary<string, object?> { ["name"] = control, ["action"] = action };
        if (item != null) data["item"] = item;

        return Forward("context_menu", data);
    }
}
