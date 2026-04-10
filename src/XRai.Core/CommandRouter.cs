using System.Text.Json;
using System.Text.Json.Nodes;

namespace XRai.Core;

public delegate string CommandHandler(JsonObject args);

public class CommandRouter
{
    private readonly Dictionary<string, CommandHandler> _handlers = new(StringComparer.OrdinalIgnoreCase);
    private readonly EventStream _events;

    /// <summary>
    /// Global default timeout in milliseconds for every command. Can be overridden
    /// per-command by passing "timeout" in the command args, or via --timeout on the CLI.
    /// </summary>
    public int DefaultTimeoutMs { get; set; } = 15000;

    /// <summary>
    /// When a command times out and its background thread is still running, the
    /// process is "tainted" — no further commands should be trusted on this router
    /// because a COM call may still be in flight. The caller should exit and restart.
    /// </summary>
    public bool IsTainted { get; private set; }

    /// <summary>
    /// Optional STA worker for COM operations. When set (via SetStaInvoker), all
    /// command handlers run on the dedicated STA thread instead of spawning new
    /// background threads. This is required for IOleMessageFilter to take effect
    /// and for reliable cross-apartment COM marshalling.
    /// </summary>
    private Func<Func<string>, int, string>? _staInvoker;
    private ITimeoutDiagnostics? _timeoutDiagnostics;

    public void SetStaInvoker(Func<Func<string>, int, string>? invoker)
    {
        _staInvoker = invoker;
    }

    public void SetTimeoutDiagnostics(ITimeoutDiagnostics? diagnostics)
    {
        _timeoutDiagnostics = diagnostics;
    }

    public CommandRouter(EventStream events)
    {
        _events = events;
    }

    public void Register(string command, CommandHandler handler)
    {
        _handlers[command] = handler;
    }

    public IEnumerable<string> RegisteredCommands => _handlers.Keys.OrderBy(c => c);
    public int CommandCount => _handlers.Count;

    public string Dispatch(string jsonLine)
    {
        try
        {
            var node = JsonNode.Parse(jsonLine);
            if (node is not JsonObject obj)
                return Response.Error("Invalid JSON: expected object");

            return DispatchObject(obj);
        }
        catch (JsonException ex)
        {
            // Common issue: bash shells strip double-backslashes in paths,
            // turning valid JSON "C:\\Temp" into invalid "C:\Temp" by the
            // time it reaches stdin. Fix by re-escaping lone backslashes
            // inside JSON string values and retrying the parse.
            try
            {
                var repaired = FixWindowsPathBackslashes(jsonLine);
                if (repaired != jsonLine)
                {
                    var repairedNode = JsonNode.Parse(repaired);
                    if (repairedNode is JsonObject repairedObj)
                        return DispatchObject(repairedObj);
                }
            }
            catch { }

            return Response.Error($"JSON parse error: {ex.Message}", code: ErrorCodes.InvalidJson);
        }
        catch (Exception ex)
        {
            return Response.ErrorFromException(ex);
        }
    }

    private string DispatchObject(JsonObject obj)
    {
        var cmd = obj["cmd"]?.GetValue<string>();
        if (string.IsNullOrEmpty(cmd))
            return Response.Error("Missing 'cmd' field", code: ErrorCodes.MissingArgument);

        // Handle batch command (no timeout wrapper — batch coordinates its own timeouts)
        if (cmd == "batch")
            return HandleBatch(obj);

        if (!_handlers.TryGetValue(cmd, out var handler))
            return Response.Error($"Unknown command: {cmd}. Use {{\"cmd\":\"help\"}} to list all commands.", code: ErrorCodes.UnknownCommand);

        int timeoutMs = obj["timeout"]?.GetValue<int>() ?? DefaultTimeoutMs;
        return InvokeWithTimeout(cmd, handler, obj, timeoutMs);
    }

    // Slow commands that legitimately take longer than the default 15s timeout.
    // Workbook I/O can hit file dialogs, external links refresh, protected view
    // prompts, etc. These get their own default timeouts unless the caller
    // overrides via an explicit "timeout" arg.
    private static readonly Dictionary<string, int> SlowCommandDefaults = new(StringComparer.OrdinalIgnoreCase)
    {
        // File operations can block on protected view, external link refresh,
        // macro prompts, and other UI callbacks. The COM RPC layer can't
        // interrupt these from our MTA worker thread, so the handler thread
        // stays blocked until Excel finishes. 300s covers the reasonable worst
        // case for a legitimate file open; anything longer is a real hang that
        // deserves a timeout. Override via {"timeout": N} in the command args.
        ["workbook.open"] = 300_000,
        ["workbook.save"] = 300_000,
        ["workbook.saveas"] = 300_000,
        ["workbook.close"] = 300_000,
        ["calc"] = 120_000,
        ["time.calc"] = 180_000,
        ["reload"] = 120_000,
        // rebuild does: kill Excel → NuGet restore → dotnet build → launch .xll
        // → attach COM → reconnect hooks. The build step alone can take 30-60s
        // on a cold build. 180s covers the realistic worst case.
        ["rebuild"] = 180_000,
    };

    /// <summary>
    /// Run a command handler on a background thread with a timeout. If the thread
    /// doesn't return in time, the command returns a timeout error and the thread
    /// is abandoned (still running, will die with the process). This prevents
    /// indefinite hangs when Excel is stuck in a modal dialog or COM deadlock.
    ///
    /// Includes a grace-period race fix: after the Wait() times out, we check once
    /// more with a 250ms grace window in case the handler finished RIGHT as we
    /// timed out. This eliminates the common "phantom timeout" failure mode where
    /// workbook.open actually succeeds but the CLI reports a timeout error because
    /// the result arrived a handful of milliseconds after the deadline.
    /// </summary>
    private string InvokeWithTimeout(string cmdName, CommandHandler handler, JsonObject args, int timeoutMs)
    {
        // Apply slow-command default if the caller didn't explicitly override
        if (args["timeout"] == null && SlowCommandDefaults.TryGetValue(cmdName, out var slowDefault))
        {
            if (timeoutMs < slowDefault) timeoutMs = slowDefault;
        }

        // timeout:0 = fire-and-forget: dispatch to the STA worker (if available)
        // but return immediately with ok:true without waiting for the result.
        // Used for modal-opening Commands where the agent knows it'll drive
        // the dialog separately and doesn't want a 15s timeout error.
        if (timeoutMs == 0 && _staInvoker != null)
        {
            _ = Task.Run(() =>
            {
                try { _staInvoker(() =>
                {
                    try { return handler(args); }
                    catch { return ""; }
                }, 300_000); }
                catch { }
            });
            return Response.Ok(new { fire_and_forget = true, command = cmdName });
        }

        if (timeoutMs < 0)
        {
            // Negative timeout = no timeout, synchronous on calling thread (legacy/tests)
            try { return handler(args); }
            catch (Exception ex) { return Response.Error($"{ex.GetType().Name}: {ex.Message}"); }
        }

        // PREFERRED PATH: route through the STA worker if one is registered.
        // This is where IOleMessageFilter lives and where all COM calls should
        // happen. The worker serializes work via a single-threaded queue.
        if (_staInvoker != null)
        {
            try
            {
                return _staInvoker(() =>
                {
                    try { return handler(args); }
                    catch (Exception ex) { return Response.Error($"{ex.GetType().Name}: {ex.Message}"); }
                }, timeoutMs);
            }
            catch (TimeoutException)
            {
                IsTainted = true;

                // Probe for an open dialog that's likely blocking the STA thread.
                // Win32 EnumWindows is thread-agnostic, so this is safe to call
                // from the caller thread while the STA is still stuck.
                object? dialogSnapshot = null;
                try { dialogSnapshot = _timeoutDiagnostics?.GetDialogSnapshot(); }
                catch { }

                return Response.ErrorWithData(
                    $"Command '{cmdName}' timed out after {timeoutMs}ms on STA worker.",
                    new { dialog = dialogSnapshot }
                );
            }
            catch (Exception ex)
            {
                return Response.Error($"{ex.GetType().Name}: {ex.Message}");
            }
        }

        // LEGACY FALLBACK: when no STA worker is registered, spawn a background
        // thread with timeout. This path is used only when the router is
        // constructed without an STA worker (e.g., in unit tests) — it doesn't
        // get IOleMessageFilter protection.
        string? result = null;
        Exception? captured = null;
        var done = new ManualResetEventSlim(false);

        var thread = new Thread(() =>
        {
            try { result = handler(args); }
            catch (Exception ex) { captured = ex; }
            finally { done.Set(); }
        })
        {
            IsBackground = true,
            Name = $"xrai-cmd-{cmdName}"
        };
        thread.Start();

        if (done.Wait(timeoutMs))
        {
            if (captured != null)
                return Response.Error($"{captured.GetType().Name}: {captured.Message}");
            return result ?? Response.Error("Handler returned null");
        }

        // Grace period for the phantom-timeout race
        if (done.Wait(250))
        {
            if (captured != null)
                return Response.Error($"{captured.GetType().Name}: {captured.Message}");
            return result ?? Response.Error("Handler returned null");
        }

        IsTainted = true;
        return Response.Error(
            $"Command '{cmdName}' timed out after {timeoutMs}ms. " +
            "Suggestions: dialog.dismiss, win32.dialog.dismiss, kill-excel, or increase timeout."
        );
    }

    private string HandleBatch(JsonObject obj)
    {
        var commands = obj["commands"]?.AsArray();
        if (commands == null)
            return Response.Error("batch requires a 'commands' array");

        int batchTimeoutMs = obj["timeout"]?.GetValue<int>() ?? DefaultTimeoutMs;

        var results = new JsonArray();
        foreach (var cmdNode in commands)
        {
            if (cmdNode is not JsonObject cmdObj)
            {
                results.Add(JsonNode.Parse(Response.Error("Invalid command in batch", code: ErrorCodes.InvalidArgument)));
                continue;
            }

            var cmdStr = cmdObj["cmd"]?.GetValue<string>();
            if (string.IsNullOrEmpty(cmdStr) || !_handlers.TryGetValue(cmdStr, out var handler))
            {
                results.Add(JsonNode.Parse(Response.Error($"Unknown command: {cmdStr}", code: ErrorCodes.UnknownCommand)));
                continue;
            }

            // Per-command timeout overrides batch timeout
            int cmdTimeout = cmdObj["timeout"]?.GetValue<int>() ?? batchTimeoutMs;
            var result = InvokeWithTimeout(cmdStr, handler, cmdObj, cmdTimeout);
            results.Add(JsonNode.Parse(result));

            // If we taint mid-batch, abort the rest — process should exit
            if (IsTainted)
            {
                results.Add(JsonNode.Parse(Response.Error("Batch aborted: previous command timed out, subsequent commands skipped")));
                break;
            }
        }

        return Response.Ok(new { results });
    }

    /// <summary>
    /// Fix unescaped Windows path backslashes in JSON strings.
    /// Bash shells strip double-backslashes: echo '{"path":"C:\\Temp"}' arrives
    /// at stdin as {"path":"C:\Temp"} which is invalid JSON (\T is not a valid
    /// escape). This method re-escapes lone backslashes inside JSON string values.
    ///
    /// Only processes characters inside double-quoted strings, and only fixes
    /// backslashes NOT already followed by a valid JSON escape character
    /// (", \, /, b, f, n, r, t, u).
    /// </summary>
    private static string FixWindowsPathBackslashes(string json)
    {
        var sb = new System.Text.StringBuilder(json.Length + 32);
        bool inString = false;
        for (int i = 0; i < json.Length; i++)
        {
            char c = json[i];

            if (c == '"' && (i == 0 || json[i - 1] != '\\'))
            {
                inString = !inString;
                sb.Append(c);
                continue;
            }

            if (inString && c == '\\' && i + 1 < json.Length)
            {
                char next = json[i + 1];
                // Valid JSON escape chars: " \ / b f n r t u
                if (next == '"' || next == '\\' || next == '/' ||
                    next == 'b' || next == 'f' || next == 'n' ||
                    next == 'r' || next == 't' || next == 'u')
                {
                    // Already a valid escape — pass through as-is
                    sb.Append(c);
                }
                else
                {
                    // Lone backslash followed by non-escape char (e.g. \T, \C, \U)
                    // This is a Windows path backslash that bash stripped.
                    // Double it so JSON parser sees \\ → literal backslash.
                    sb.Append('\\');
                    sb.Append('\\');
                    // Skip appending c again since we just wrote \\
                    continue;
                }
            }
            else
            {
                sb.Append(c);
            }
        }
        return sb.ToString();
    }
}
