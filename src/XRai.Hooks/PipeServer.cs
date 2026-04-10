// Leak-audited: 2026-04-10 — Stop() now joins the background ServerLoop
// thread, disposes the CancellationTokenSource, and clears _activeWriter
// under the writer lock. The named pipe instance is disposed in the
// ServerLoop's per-iteration finally block, so each disconnected client
// frees its pipe handle immediately. There are no static caches, no
// long-lived event subscriptions, and no per-command allocations that
// outlive the response.

using System.Diagnostics;
using System.IO;
using System.IO.Pipes;
using System.Security.Principal;
using System.Text.Json;
using System.Text.Json.Nodes;
using System.Windows;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Threading;
using XRai.Core;

namespace XRai.Hooks;

public class PipeServer
{
    private readonly string _pipeName;
    private readonly ControlRegistry _controls;
    private readonly ModelRegistry _models;
    private CancellationTokenSource? _cts;
    private Thread? _thread;
    private StreamWriter? _activeWriter;
    private readonly object _writerLock = new();
    private Dispatcher? _uiDispatcher;

    /// <summary>
    /// True if the pipe was created with the restricted ACL (per-user + SYSTEM).
    /// Surfaced by the security.status command.
    /// </summary>
    public bool PipeAclRestricted { get; private set; }

    /// <summary>
    /// True if token-based authentication is currently enforced. Always true unless
    /// XRAI_ALLOW_UNAUTH=1 is set, in which case the server logs a loud warning.
    /// </summary>
    public bool TokenAuthEnabled { get; private set; } = true;

    public string PipeName => _pipeName;

    public PipeServer(string pipeName, ControlRegistry controls, ModelRegistry models)
    {
        _pipeName = pipeName;
        _controls = controls;
        _models = models;

        // UI dispatcher set later via SetDispatcher when WPF thread is ready
    }

    public void SetDispatcher(Dispatcher dispatcher)
    {
        _uiDispatcher = dispatcher;
    }

    public void Start()
    {
        // Generate a fresh auth token for this pipe. Stored in
        // %LOCALAPPDATA%\XRai\tokens\{pipeName}.token with a per-user ACL.
        try
        {
            PipeAuth.GenerateAndStoreToken(_pipeName);
            TokenAuthEnabled = !PipeAuth.AllowUnauthenticated;
            if (PipeAuth.AllowUnauthenticated)
            {
                Debug.WriteLine("[XRai.Hooks] WARNING: XRAI_ALLOW_UNAUTH=1 is set — pipe accepts unauthenticated clients. Do NOT use this outside of a trusted diagnostic environment.");
            }
        }
        catch (Exception ex)
        {
            Debug.WriteLine($"[XRai.Hooks] WARNING: Failed to provision auth token: {ex.Message}. Pipe will start without token auth.");
            TokenAuthEnabled = false;
        }

        _cts = new CancellationTokenSource();
        _thread = new Thread(ServerLoop) { IsBackground = true, Name = "XRai-PipeServer" };
        _thread.Start();
    }

    public void Stop()
    {
        // Leak-audited: 2026-04-10. Stop() now joins the server thread,
        // disposes the cancellation source, and clears the active writer
        // reference so a clean Stop() doesn't leak the worker thread, the
        // CTS handle, or the StreamWriter held under _writerLock.
        try { _cts?.Cancel(); } catch { }

        // Best-effort: give the server thread a moment to drain any in-flight
        // command and exit ServerLoop. We never block forever here — the
        // background thread is IsBackground=true so process shutdown will
        // reap it regardless. The Join is purely to surface clean shutdown
        // when the host has time to wait.
        try { _thread?.Join(2000); } catch { }

        try { _cts?.Dispose(); } catch { }
        _cts = null;
        _thread = null;

        lock (_writerLock) { _activeWriter = null; }

        // Delete the token file on graceful shutdown so no stale token survives.
        try { PipeAuth.ClearToken(_pipeName); } catch { }
    }

    // When true, a command is in flight (HandleCommand is executing, possibly
    // blocked on InvokeOnUI waiting for a modal to close). PushEvent must NOT
    // write to the pipe during this window — the client is blocked on ReadLine
    // expecting the command response, and any interleaved event line would be
    // consumed as the response. Since event JSON doesn't carry ok:true, the
    // client falls back to "Command failed: {cmd}", creating a false-negative
    // response even though the command actually succeeded.
    private volatile bool _commandInFlight;

    public void PushEvent(string eventType, object? data = null)
    {
        // Suppress event writes while a command is in flight. Events are
        // informational (log, error captures) and non-critical — dropping
        // them during a command call is safe. The alternative (queueing and
        // flushing after the response) adds complexity for zero value since
        // the client doesn't consume events today.
        if (_commandInFlight) return;

        lock (_writerLock)
        {
            if (_activeWriter == null) return;
            try
            {
                var node = new JsonObject { ["event"] = eventType };
                if (data != null)
                {
                    var json = JsonSerializer.Serialize(data);
                    var parsed = JsonNode.Parse(json);
                    if (parsed is JsonObject obj)
                    {
                        foreach (var kvp in obj)
                            node[kvp.Key] = kvp.Value?.DeepClone();
                    }
                }
                _activeWriter.WriteLine(node.ToJsonString());
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"PushEvent error: {ex.Message}");
            }
        }
    }

    private void ServerLoop()
    {
        while (_cts != null && !_cts.IsCancellationRequested)
        {
            NamedPipeServerStream? pipe = null;
            try
            {
                // Create the pipe with a restricted ACL (current user + SYSTEM).
                // Any local process running under a different user gets
                // ERROR_ACCESS_DENIED when it tries to open the pipe.
                try
                {
                    pipe = PipeAuth.CreateRestrictedServerPipe(_pipeName, maxInstances: 1);
                    PipeAclRestricted = true;
                }
                catch (Exception aclEx)
                {
                    // If the ACL factory fails (rare — e.g. broken SID resolution),
                    // fall back to the unsecured constructor so the add-in keeps
                    // working. The PipeAclRestricted flag surfaces the degraded
                    // state via security.status.
                    Debug.WriteLine($"[XRai.Hooks] WARNING: Could not create restricted pipe ACL ({aclEx.Message}). Falling back to unsecured pipe.");
                    pipe = new NamedPipeServerStream(_pipeName, PipeDirection.InOut, 1,
                        PipeTransmissionMode.Byte, PipeOptions.Asynchronous);
                    PipeAclRestricted = false;
                }

                var connectTask = pipe.WaitForConnectionAsync(_cts!.Token);
                connectTask.Wait(_cts.Token);

                Debug.WriteLine("XRai pipe client connected");

                using var reader = new StreamReader(pipe);
                var writer = new StreamWriter(pipe) { AutoFlush = true };

                // === Auth handshake ===
                // The first line from the client must be {"auth_token":"..."}.
                // Any mismatch closes the pipe immediately.
                if (!PerformAuthHandshake(reader, writer))
                {
                    Debug.WriteLine("XRai pipe auth failed — closing connection");
                    continue;
                }

                lock (_writerLock) { _activeWriter = writer; }

                while (pipe.IsConnected && !_cts.IsCancellationRequested)
                {
                    string? line = reader.ReadLine();
                    if (line == null) break;

                    // Suppress PushEvent writes for the duration of the command.
                    // HandleCommand can block for a long time (OnClick → ShowDialog
                    // → nested dispatcher frame). Any unsolicited PushEvent write
                    // during that window would interleave with the command response
                    // and desync the client's ReadLine.
                    _commandInFlight = true;
                    try
                    {
                        var response = HandleCommand(line);
                        writer.WriteLine(response);
                    }
                    finally
                    {
                        _commandInFlight = false;
                    }
                }

                lock (_writerLock) { _activeWriter = null; }
                Debug.WriteLine("XRai pipe client disconnected");
            }
            catch (OperationCanceledException) { break; }
            catch (Exception ex)
            {
                Debug.WriteLine($"Pipe error: {ex.Message}");
                lock (_writerLock) { _activeWriter = null; }
            }
            finally
            {
                pipe?.Dispose();
            }
        }
    }

    /// <summary>
    /// Read the first line from the client and validate it as an auth handshake.
    /// Returns true if the client is authorised to proceed. Logs rejected attempts
    /// to Debug.WriteLine and writes a structured error response to the client
    /// before returning false.
    ///
    /// If <see cref="PipeAuth.AllowUnauthenticated"/> is true (XRAI_ALLOW_UNAUTH=1),
    /// the handshake is waived but a warning is logged for every connection.
    /// </summary>
    private bool PerformAuthHandshake(StreamReader reader, StreamWriter writer)
    {
        string? firstLine;
        try { firstLine = reader.ReadLine(); }
        catch (Exception ex)
        {
            Debug.WriteLine($"[XRai.Hooks] Handshake read failed: {ex.Message}");
            return false;
        }

        if (firstLine == null) return false;

        var token = PipeAuth.TryExtractAuthToken(firstLine);

        if (PipeAuth.ValidateToken(_pipeName, token))
        {
            try { writer.WriteLine(PipeAuth.BuildAuthOkResponse()); } catch { }
            return true;
        }

        // Token missing or invalid. Honor backward-compat escape hatch.
        if (PipeAuth.AllowUnauthenticated)
        {
            Debug.WriteLine($"[XRai.Hooks] WARNING: Accepting unauthenticated client on pipe '{_pipeName}' because XRAI_ALLOW_UNAUTH=1. First line was: {Truncate(firstLine, 200)}");

            // The first line might have been a real command rather than a handshake
            // (legacy client that doesn't know about auth). Dispatch it so we don't
            // lose the command.
            try
            {
                _commandInFlight = true;
                var response = HandleCommand(firstLine);
                writer.WriteLine(response);
            }
            catch (Exception ex) { Debug.WriteLine($"[XRai.Hooks] Unauth legacy dispatch error: {ex.Message}"); }
            finally { _commandInFlight = false; }

            return true;
        }

        Debug.WriteLine($"[XRai.Hooks] REJECTED unauthenticated client on pipe '{_pipeName}'. First line was: {Truncate(firstLine, 200)}");
        try { writer.WriteLine(PipeAuth.BuildAuthFailedResponse()); } catch { }
        return false;
    }

    private static string Truncate(string s, int max) =>
        s.Length <= max ? s : s.Substring(0, max) + "...";

    private string HandleCommand(string json)
    {
        try
        {
            var node = JsonNode.Parse(json);
            var cmd = node?["cmd"]?.GetValue<string>();

            return cmd switch
            {
                "ping" => HandlePing(),
                "info" => HandleInfo(),
                "pane_status" => HandlePaneStatus(),
                "controls" => HandleControls(),
                "set_control" => HandleSetControl(node!),
                "click" => HandleClick(node!),
                "toggle_control" => HandleToggle(node!),
                "select_control" => HandleSelect(node!),
                "read_control" => HandleReadControl(node!),
                "model" => HandleModel(node!),
                "model_set" => HandleModelSet(node!),
                "functions" => HandleFunctions(),
                // Human simulation commands
                "double_click" => HandleDoubleClick(node!),
                "right_click" => HandleRightClick(node!),
                "hover" => HandleHover(node!),
                "focus" => HandleFocus(node!),
                "send_keys" => HandleSendKeys(node!),
                "scroll" => HandleScroll(node!),
                "control_info" => HandleControlInfo(node!),
                "pane_tree" => HandlePaneTree(),
                "pane_screenshot" => HandlePaneScreenshotCapture(),
                "datagrid_read" => HandleDataGridRead(node!),
                "datagrid_cell" => HandleDataGridCell(node!),
                "datagrid_select" => HandleDataGridSelectRow(node!),
                "tree_expand" => HandleTreeExpand(node!),
                "tab_select" => HandleTabSelect(node!),
                "list_select" => HandleListSelect(node!),
                "list_read" => HandleListRead(node!),
                "expand" => HandleExpand(node!),
                "wait_control" => HandleWaitControl(node!),
                "drag" => HandleDrag(node!),
                "context_menu" => HandleContextMenu(node!),
                "security.status" => HandleSecurityStatus(),
                "security_status" => HandleSecurityStatus(),
                "log_read" => HandleLogRead(node!),
                _ => Serialize(new { ok = false, error = $"Unknown command: {cmd}", code = "XRAI_UNKNOWN_COMMAND" }),
            };
        }
        catch (Exception ex)
        {
            return Serialize(new { ok = false, error = $"{ex.GetType().Name}: {ex.Message}", code = "XRAI_INTERNAL_ERROR" });
        }
    }

    private string HandlePing()
    {
        return Serialize(new
        {
            ok = true,
            message = "pong",
            pid = Process.GetCurrentProcess().Id,
            addin = "XRai.Hooks",
        });
    }

    private string HandleInfo()
    {
        string excelVersion = "unknown";
        try
        {
            dynamic app = ExcelDna.Integration.ExcelDnaUtil.Application;
            excelVersion = app.Version;
        }
        catch { }

        return Serialize(new
        {
            ok = true,
            excel_version = excelVersion,
            controls_count = _controls.Count,
            models_count = _models.All.Count(),
            hooks_active = true,
            pipe_name = $"xrai_{Process.GetCurrentProcess().Id}",
        });
    }

    /// <summary>
    /// Returns detailed pane/model exposure status so agents can diagnose whether:
    ///   - the pipe is alive but Pilot.Expose wasn't called yet
    ///   - Expose was called but controls are missing x:Name
    ///   - a model is exposed but controls aren't (or vice versa)
    /// </summary>
    private string HandlePaneStatus()
    {
        var asm = typeof(Pilot).Assembly;
        var version = asm.GetName().Version?.ToString() ?? "unknown";
        string buildTimestamp = "unknown";
        try
        {
            var metadata = asm.GetCustomAttributes(typeof(System.Reflection.AssemblyMetadataAttribute), false)
                .Cast<System.Reflection.AssemblyMetadataAttribute>()
                .FirstOrDefault(m => m.Key == "XRaiBuildTimestamp");
            if (metadata?.Value != null)
                buildTimestamp = metadata.Value;
        }
        catch { }

        // Best-effort: derive the consuming add-in's build timestamp from the
        // .xll's File.GetLastWriteTime. This is what users actually want to see
        // when they ask "when was my add-in built?" — the hooks DLL timestamp
        // is just XRai's library build, not theirs.
        string? addinBuildTimestamp = null;
        try
        {
            var addinPath = ExcelDna.Integration.ExcelDnaUtil.XllPath;
            if (!string.IsNullOrEmpty(addinPath) && File.Exists(addinPath))
            {
                addinBuildTimestamp = File.GetLastWriteTimeUtc(addinPath).ToString("o");
            }
        }
        catch { }

        return Serialize(new
        {
            ok = true,
            pipe_connected = true,
            exposed_controls = _controls.Count,
            exposed_models = _models.All.Count(),
            last_expose_at = Pilot.LastExposeAt?.ToString("o"),
            last_expose_model_at = Pilot.LastExposeModelAt?.ToString("o"),
            total_expose_calls = Pilot.TotalExposeCalls,
            total_expose_model_calls = Pilot.TotalExposeModelCalls,
            hooks_assembly_version = version,
            // Canonical: when the XRai.Hooks library DLL was compiled.
            hooks_library_build_timestamp = buildTimestamp,
            // Canonical: when the consuming add-in's .xll was last written.
            // Null if the .xll path can't be resolved (e.g. running outside Excel).
            addin_build_timestamp = addinBuildTimestamp,
            // DEPRECATED: kept for backward compatibility with older CLI binaries.
            // Use hooks_library_build_timestamp instead. Users frequently misread
            // this field as "when my add-in was built" — it's the library timestamp.
            hooks_build_timestamp = buildTimestamp,
            hint = _controls.Count == 0
                ? "Pipe is live but no controls exposed. Check Pilot.Expose() is called on your WPF task pane AFTER the visual tree is rendered. Controls must have x:Name in XAML to be discovered."
                : null,
        });
    }

    private string HandleSecurityStatus()
    {
        string currentUser;
        try { currentUser = WindowsIdentity.GetCurrent().Name; }
        catch { currentUser = "unknown"; }

        var tokenPath = PipeAuth.GetTokenFilePath(_pipeName);
        bool tokenExists = false;
        try { tokenExists = File.Exists(tokenPath); } catch { }

        return Serialize(new
        {
            ok = true,
            pipe_acl_restricted = PipeAclRestricted,
            token_auth_enabled = TokenAuthEnabled,
            token_file_exists = tokenExists,
            token_file_path = tokenPath,
            hooks_pipe_name = _pipeName,
            daemon_pipe_name = (string?)null,
            current_user = currentUser,
            allow_unauthenticated = PipeAuth.AllowUnauthenticated,
        });
    }

    private string HandleControls()
    {
        return InvokeOnUI(() =>
        {
            var list = _controls.All.Select(kvp => new
            {
                name = kvp.Key,
                type = kvp.Value.Type,
                value = kvp.Value.GetValue(),
                enabled = kvp.Value.IsEnabled,
                has_command = kvp.Value.HasCommand,
            }).ToArray();

            return Serialize(new { ok = true, controls = list });
        });
    }

    /// <summary>
    /// Resolve a control by name, optionally polling for it to appear if a
    /// non-zero <paramref name="timeoutMs"/> was supplied. Returns true and sets
    /// <paramref name="ctrl"/> if found within the budget; false otherwise.
    ///
    /// Eliminates the "ribbon.click → pane.click race" where the pane's WPF
    /// Loaded event hasn't fired yet by the time the next pane command arrives.
    /// Polls every 100ms via the UI dispatcher (so we observe newly registered
    /// controls without races against the registration thread).
    /// </summary>
    private bool TryGetControlWithWait(string name, int timeoutMs, out IControlAdapter? ctrl)
    {
        if (_controls.TryGet(name, out ctrl!)) return true;
        if (timeoutMs <= 0) { ctrl = null; return false; }

        var sw = Stopwatch.StartNew();
        while (sw.ElapsedMilliseconds < timeoutMs)
        {
            Thread.Sleep(100);
            if (_controls.TryGet(name, out ctrl!)) return true;
        }
        ctrl = null;
        return false;
    }

    /// <summary>
    /// Read the optional 'timeout' parameter (in milliseconds) used by pane
    /// commands to auto-wait for a control to appear. Defaults to 0 (no wait)
    /// to preserve existing behavior.
    /// </summary>
    private static int ReadTimeoutMs(JsonNode node) =>
        node?["timeout"]?.GetValue<int>() ?? 0;

    private string HandleSetControl(JsonNode node)
    {
        var name = node["name"]?.GetValue<string>()
            ?? throw new ArgumentException("set_control requires 'name'");
        var value = node["value"]?.GetValue<string>()
            ?? throw new ArgumentException("set_control requires 'value'");
        int timeoutMs = ReadTimeoutMs(node);

        if (!TryGetControlWithWait(name, timeoutMs, out var preCtrl))
            return Serialize(new
            {
                ok = false,
                error = $"Control not found: {name}",
                code = "XRAI_PANE_CONTROL_NOT_FOUND",
                timeout_ms = timeoutMs,
            });

        return InvokeOnUI(() =>
        {
            if (!_controls.TryGet(name, out var ctrl))
                return Serialize(new { ok = false, error = $"Control not found: {name}", code = "XRAI_PANE_CONTROL_NOT_FOUND" });

            var oldValue = ctrl.GetValue();
            ctrl.SetValue(value);
            return Serialize(new { ok = true, name, old_value = oldValue, new_value = value });
        });
    }

    private string HandleClick(JsonNode node)
    {
        var name = node["name"]?.GetValue<string>()
            ?? throw new ArgumentException("click requires 'name'");
        int timeoutMs = ReadTimeoutMs(node);

        if (!TryGetControlWithWait(name, timeoutMs, out var _))
        {
            var controlNames = string.Join(", ", _controls.All.Select(kvp => kvp.Key).Take(20));
            return Serialize(new
            {
                ok = false,
                error = $"Control not found: {name}",
                code = "XRAI_PANE_CONTROL_NOT_FOUND",
                hint = "Run {\"cmd\":\"pane\"} to see all exposed control names. Common causes: (1) missing x:Name in XAML, (2) Pilot.Expose called before visual tree built, (3) typo in control name. Use 'timeout' (ms) to auto-wait for the control to appear.",
                available_sample = controlNames,
                exposed_count = _controls.Count,
                timeout_ms = timeoutMs,
            });
        }

        return InvokeOnUI(() =>
        {
            if (!_controls.TryGet(name, out var ctrl))
            {
                var controlNames = string.Join(", ", _controls.All.Select(kvp => kvp.Key).Take(20));
                return Serialize(new
                {
                    ok = false,
                    error = $"Control not found: {name}",
                    code = "XRAI_PANE_CONTROL_NOT_FOUND",
                    hint = "Run {\"cmd\":\"pane\"} to see all exposed control names. Common causes: (1) missing x:Name in XAML, (2) Pilot.Expose called before visual tree built, (3) typo in control name.",
                    available_sample = controlNames,
                    exposed_count = _controls.Count,
                });
            }

            ControlAdapter.ClickResult result;
            try
            {
                result = ctrl.Click();
            }
            catch (Exception ex)
            {
                // OnClick handler threw (or reflection couldn't find the method,
                // or Command.Execute threw, etc.). Return verbose error with the
                // full exception so callers can debug without guessing.
                return Serialize(new
                {
                    ok = false,
                    error = $"pane.click failed: {ex.GetType().Name}: {ex.Message}",
                    exception_type = ex.GetType().Name,
                    exception_message = ex.Message,
                    inner_exception = ex.InnerException?.Message,
                    stack_frame = GetTopStackFrame(ex),
                    name,
                    hint = "The OnClick handler or bound Command threw an exception. Check the Command implementation in the ViewModel, or the button's Click event handler in the code-behind.",
                });
            }

            // Synchronous, single-invocation contract.
            // Return verbose result whether or not Click() succeeded — caller
            // needs method, resolved_target_type, command_can_execute to debug
            // silent no-ops (has_command=true but command_executed=false).
            return Serialize(new
            {
                ok = true,
                name,
                clicked = true,
                method = result.Method,
                resolved_to_button_base = result.ResolvedToButtonBase,
                resolved_target_type = result.ResolvedTargetType,
                has_command = result.HasCommand,
                command_can_execute = result.CommandCanExecute,
                command_executed = result.CommandExecuted,
                warning = result.ErrorHint,
            });
        });
    }

    private static string? GetTopStackFrame(Exception ex)
    {
        try
        {
            var stack = ex.StackTrace;
            if (string.IsNullOrEmpty(stack)) return null;
            var firstLine = stack.Split('\n')[0].Trim();
            if (firstLine.StartsWith("at ")) firstLine = firstLine.Substring(3);
            return firstLine.Length > 200 ? firstLine.Substring(0, 200) : firstLine;
        }
        catch { return null; }
    }

    private string HandleToggle(JsonNode node)
    {
        var name = node["name"]?.GetValue<string>()
            ?? throw new ArgumentException("toggle_control requires 'name'");
        int timeoutMs = ReadTimeoutMs(node);

        if (!TryGetControlWithWait(name, timeoutMs, out var _))
            return Serialize(new { ok = false, error = $"Control not found: {name}", code = "XRAI_PANE_CONTROL_NOT_FOUND", timeout_ms = timeoutMs });

        return InvokeOnUI(() =>
        {
            if (!_controls.TryGet(name, out var ctrl))
                return Serialize(new { ok = false, error = $"Control not found: {name}", code = "XRAI_PANE_CONTROL_NOT_FOUND" });

            ctrl.Toggle();
            return Serialize(new { ok = true, name, toggled = true });
        });
    }

    private string HandleSelect(JsonNode node)
    {
        var name = node["name"]?.GetValue<string>()
            ?? throw new ArgumentException("select_control requires 'name'");
        var value = node["value"]?.GetValue<string>()
            ?? throw new ArgumentException("select_control requires 'value'");
        int timeoutMs = ReadTimeoutMs(node);

        if (!TryGetControlWithWait(name, timeoutMs, out var _))
            return Serialize(new { ok = false, error = $"Control not found: {name}", code = "XRAI_PANE_CONTROL_NOT_FOUND", timeout_ms = timeoutMs });

        return InvokeOnUI(() =>
        {
            if (!_controls.TryGet(name, out var ctrl))
                return Serialize(new { ok = false, error = $"Control not found: {name}", code = "XRAI_PANE_CONTROL_NOT_FOUND" });

            ctrl.SetValue(value);
            return Serialize(new { ok = true, name, selected = value });
        });
    }

    private string HandleReadControl(JsonNode node)
    {
        var name = node["name"]?.GetValue<string>()
            ?? throw new ArgumentException("read_control requires 'name'");
        int timeoutMs = ReadTimeoutMs(node);

        if (!TryGetControlWithWait(name, timeoutMs, out var _))
            return Serialize(new { ok = false, error = $"Control not found: {name}", code = "XRAI_PANE_CONTROL_NOT_FOUND", timeout_ms = timeoutMs });

        return InvokeOnUI(() =>
        {
            if (!_controls.TryGet(name, out var ctrl))
                return Serialize(new { ok = false, error = $"Control not found: {name}", code = "XRAI_PANE_CONTROL_NOT_FOUND" });

            return Serialize(new { ok = true, name, type = ctrl.Type, value = ctrl.GetValue(), enabled = ctrl.IsEnabled });
        });
    }

    private string HandleModel(JsonNode node)
    {
        // Optional 'name' selects a specific keyed model. Without 'name', the
        // default model (most-recently exposed) is used. This lets callers
        // disambiguate when multiple models have been exposed via
        // ExposeModel(vm, "Foo") + ExposeModel(vm2, "Bar").
        var name = node?["name"]?.GetValue<string>();
        ModelAdapter? model;
        if (!string.IsNullOrEmpty(name))
        {
            if (!_models.TryGet(name, out model!))
            {
                var available = string.Join(", ", _models.All.Select(kvp => kvp.Key));
                return Serialize(new
                {
                    ok = false,
                    error = $"No model exposed under name: {name}",
                    available_models = available,
                });
            }
        }
        else
        {
            model = _models.Default;
        }

        if (model == null)
            return Serialize(new { ok = false, error = "No model exposed" });

        return InvokeOnUI(() =>
        {
            var props = model.GetAll();
            return Serialize(new { ok = true, name, properties = props });
        });
    }

    private string HandleModelSet(JsonNode node)
    {
        var property = node["property"]?.GetValue<string>()
            ?? throw new ArgumentException("model_set requires 'property'");
        var value = node["value"];
        // Optional 'name' to target a specific keyed model. Same lookup
        // semantics as HandleModel.
        var name = node["name"]?.GetValue<string>();

        ModelAdapter? model;
        if (!string.IsNullOrEmpty(name))
        {
            if (!_models.TryGet(name, out model!))
                return Serialize(new { ok = false, error = $"No model exposed under name: {name}" });
        }
        else
        {
            model = _models.Default;
        }

        if (model == null)
            return Serialize(new { ok = false, error = "No model exposed" });

        return InvokeOnUI(() =>
        {
            object? val = value switch
            {
                JsonValue jv when jv.TryGetValue<double>(out var d) => d,
                JsonValue jv when jv.TryGetValue<bool>(out var b) => b,
                _ => value?.ToString()
            };

            model.SetProperty(property, val);
            return Serialize(new { ok = true, property, value = val });
        });
    }

    private string HandleFunctions()
    {
        try
        {
            var functions = FunctionReporter.GetRegisteredFunctions();
            return Serialize(new { ok = true, functions });
        }
        catch (Exception ex)
        {
            return Serialize(new { ok = false, error = ex.Message });
        }
    }

    // === Human Simulation Handlers ===

    private string HandleDoubleClick(JsonNode node)
    {
        var name = node["name"]?.GetValue<string>() ?? throw new ArgumentException("double_click requires 'name'");
        return InvokeOnUI(() =>
        {
            if (!_controls.TryGet(name, out var ctrl))
                return Serialize(new { ok = false, error = $"Control not found: {name}" });
            ctrl.DoubleClick();
            return Serialize(new { ok = true, name, double_clicked = true });
        });
    }

    private string HandleRightClick(JsonNode node)
    {
        var name = node["name"]?.GetValue<string>() ?? throw new ArgumentException("right_click requires 'name'");
        return InvokeOnUI(() =>
        {
            if (!_controls.TryGet(name, out var ctrl))
                return Serialize(new { ok = false, error = $"Control not found: {name}" });
            ctrl.RightClick();
            return Serialize(new { ok = true, name, right_clicked = true });
        });
    }

    private string HandleHover(JsonNode node)
    {
        var name = node["name"]?.GetValue<string>() ?? throw new ArgumentException("hover requires 'name'");
        return InvokeOnUI(() =>
        {
            if (!_controls.TryGet(name, out var ctrl))
                return Serialize(new { ok = false, error = $"Control not found: {name}" });
            ctrl.Hover();
            return Serialize(new { ok = true, name, hovered = true });
        });
    }

    private string HandleFocus(JsonNode node)
    {
        var name = node["name"]?.GetValue<string>() ?? throw new ArgumentException("focus requires 'name'");
        return InvokeOnUI(() =>
        {
            if (!_controls.TryGet(name, out var ctrl))
                return Serialize(new { ok = false, error = $"Control not found: {name}" });
            ctrl.Focus();
            return Serialize(new { ok = true, name, focused = true });
        });
    }

    private string HandleSendKeys(JsonNode node)
    {
        var name = node["name"]?.GetValue<string>() ?? throw new ArgumentException("send_keys requires 'name'");
        // Accept both 'keys' (plural, canonical) and 'key' (singular, common typo)
        var keys = node["keys"]?.GetValue<string>()
                   ?? node["key"]?.GetValue<string>()
                   ?? throw new ArgumentException("send_keys requires 'keys' (or 'key' as alias)");

        return InvokeOnUI(() =>
        {
            if (!_controls.TryGet(name, out var ctrl))
                return Serialize(new { ok = false, error = $"Control not found: {name}" });

            var result = ctrl.SendKeys(keys);

            if (result.NoPresentationSource)
            {
                return Serialize(new
                {
                    ok = false,
                    error = result.ErrorHint,
                    name,
                    no_presentation_source = true,
                });
            }

            // Success = at least one key delivered. A post-delivery KeyUp failure
            // (e.g. because KeyDown triggered a modal dialog that invalidated the
            // source) is a WARNING, not a failure — the keystroke succeeded.
            bool anyDelivered = result.AnyDelivered;
            bool fullySuccessful = result.AllDelivered && result.PostDeliveryWarnings.Count == 0;

            // If nothing delivered, ensure an explanatory error field is present so
            // callers don't see a generic "Command failed: send_keys" fallback.
            string? errorMsg = null;
            if (!anyDelivered)
            {
                errorMsg = result.FailedKeys.Count > 0
                    ? $"No keys delivered. Failures: {string.Join("; ", result.FailedKeys)}"
                    : result.UnknownKeys.Count > 0
                        ? $"No keys delivered. Unknown key names: {string.Join(", ", result.UnknownKeys)}"
                        : "No keys delivered (unknown reason).";
            }

            return Serialize(new
            {
                ok = anyDelivered,
                error = errorMsg,
                name,
                keys_sent = keys,
                delivered = result.DeliveredKeys.ToArray(),
                failed = result.FailedKeys.ToArray(),
                unknown = result.UnknownKeys.ToArray(),
                warnings = result.PostDeliveryWarnings.ToArray(),
                fully_successful = fullySuccessful,
                focus_warning = result.FocusWarning,
            });
        });
    }

    private string HandleScroll(JsonNode node)
    {
        var name = node["name"]?.GetValue<string>() ?? throw new ArgumentException("scroll requires 'name'");
        var offset = node["offset"]?.GetValue<double>() ?? 0;
        return InvokeOnUI(() =>
        {
            if (!_controls.TryGet(name, out var ctrl))
                return Serialize(new { ok = false, error = $"Control not found: {name}" });
            ctrl.ScrollTo(offset);
            return Serialize(new { ok = true, name, scrolled_to = offset });
        });
    }

    private string HandleControlInfo(JsonNode node)
    {
        var name = node["name"]?.GetValue<string>() ?? throw new ArgumentException("control_info requires 'name'");
        return InvokeOnUI(() =>
        {
            if (!_controls.TryGet(name, out var ctrl))
                return Serialize(new { ok = false, error = $"Control not found: {name}" });
            var info = ctrl.GetDetailedInfo();
            return Serialize(new { ok = true, name, type = ctrl.Type, value = ctrl.GetValue(), enabled = ctrl.IsEnabled, visible = ctrl.IsVisible, details = info });
        });
    }

    private string HandlePaneTree()
    {
        return InvokeOnUI(() =>
        {
            var all = _controls.All.Select(kvp => new
            {
                name = kvp.Key,
                type = kvp.Value.Type,
                value = kvp.Value.GetValue(),
                enabled = kvp.Value.IsEnabled,
                visible = kvp.Value.IsVisible,
                details = kvp.Value.GetDetailedInfo(),
            }).ToArray();
            return Serialize(new { ok = true, controls = all, count = all.Length });
        });
    }

    private string HandleDataGridRead(JsonNode node)
    {
        var name = node["name"]?.GetValue<string>() ?? throw new ArgumentException("datagrid_read requires 'name'");
        int timeoutMs = ReadTimeoutMs(node);

        if (!TryGetControlWithWait(name, timeoutMs, out var _))
            return Serialize(new { ok = false, error = $"Control not found: {name}", code = "XRAI_PANE_CONTROL_NOT_FOUND", timeout_ms = timeoutMs });

        return InvokeOnUI(() =>
        {
            if (!_controls.TryGet(name, out var ctrl))
                return Serialize(new { ok = false, error = $"Control not found: {name}", code = "XRAI_PANE_CONTROL_NOT_FOUND" });
            var info = ctrl.GetDetailedInfo();
            var data = ctrl.GetDataGridAllData();
            return Serialize(new { ok = true, name, info, data });
        });
    }

    private string HandleDataGridCell(JsonNode node)
    {
        var name = node["name"]?.GetValue<string>() ?? throw new ArgumentException("datagrid_cell requires 'name'");
        var row = node["row"]?.GetValue<int>() ?? throw new ArgumentException("datagrid_cell requires 'row'");
        var col = node["col"]?.GetValue<int>() ?? throw new ArgumentException("datagrid_cell requires 'col'");
        return InvokeOnUI(() =>
        {
            if (!_controls.TryGet(name, out var ctrl))
                return Serialize(new { ok = false, error = $"Control not found: {name}" });
            var value = ctrl.GetDataGridCell(row, col);
            return Serialize(new { ok = true, name, row, col, value });
        });
    }

    private string HandleDataGridSelectRow(JsonNode node)
    {
        var name = node["name"]?.GetValue<string>() ?? throw new ArgumentException("datagrid_select requires 'name'");
        var row = node["row"]?.GetValue<int>() ?? throw new ArgumentException("datagrid_select requires 'row'");
        return InvokeOnUI(() =>
        {
            if (!_controls.TryGet(name, out var ctrl))
                return Serialize(new { ok = false, error = $"Control not found: {name}" });
            ctrl.SelectDataGridRow(row);
            return Serialize(new { ok = true, name, selected_row = row });
        });
    }

    private string HandleTreeExpand(JsonNode node)
    {
        var name = node["name"]?.GetValue<string>() ?? throw new ArgumentException("tree_expand requires 'name'");
        var path = node["path"]?.GetValue<string>() ?? throw new ArgumentException("tree_expand requires 'path'");
        return InvokeOnUI(() =>
        {
            if (!_controls.TryGet(name, out var ctrl))
                return Serialize(new { ok = false, error = $"Control not found: {name}" });
            ctrl.ExpandTreeNode(path);
            return Serialize(new { ok = true, name, expanded_path = path });
        });
    }

    private string HandleTabSelect(JsonNode node)
    {
        var name = node["name"]?.GetValue<string>() ?? throw new ArgumentException("tab_select requires 'name'");
        var tab = node["tab"]?.GetValue<string>() ?? throw new ArgumentException("tab_select requires 'tab'");
        return InvokeOnUI(() =>
        {
            if (!_controls.TryGet(name, out var ctrl))
                return Serialize(new { ok = false, error = $"Control not found: {name}" });
            ctrl.SetValue(tab);
            return Serialize(new { ok = true, name, selected_tab = tab });
        });
    }

    private string HandleExpand(JsonNode node)
    {
        var name = node["name"]?.GetValue<string>()
            ?? throw new ArgumentException("expand requires 'name'");
        var open = node["open"]?.GetValue<bool>() ?? true;

        return InvokeOnUI(() =>
        {
            if (!_controls.TryGet(name, out var ctrl))
                return Serialize(new { ok = false, error = $"Control not found: {name}" });

            var message = ctrl.Expand(open);
            return Serialize(new { ok = true, name, expanded = open, message });
        });
    }

    private string HandleListRead(JsonNode node)
    {
        var name = node["name"]?.GetValue<string>()
            ?? throw new ArgumentException("list_read requires 'name'");
        int timeoutMs = ReadTimeoutMs(node);

        if (!TryGetControlWithWait(name, timeoutMs, out var _))
            return Serialize(new { ok = false, error = $"Control not found: {name}", code = "XRAI_PANE_CONTROL_NOT_FOUND", timeout_ms = timeoutMs });

        return InvokeOnUI(() =>
        {
            if (!_controls.TryGet(name, out var ctrl))
                return Serialize(new { ok = false, error = $"Control not found: {name}", code = "XRAI_PANE_CONTROL_NOT_FOUND" });

            var (items, selectedIndex) = ctrl.ReadListItems();
            return Serialize(new
            {
                ok = true,
                name,
                items,
                count = items.Length,
                selected_index = selectedIndex,
            });
        });
    }

    private string HandleListSelect(JsonNode node)
    {
        var name = node["name"]?.GetValue<string>()
            ?? throw new ArgumentException("list_select requires 'name'");
        int? index = node["index"] is JsonNode indexNode ? indexNode.GetValue<int>() : null;
        var text = node["text"]?.GetValue<string>();
        int timeoutMs = ReadTimeoutMs(node);

        if (!TryGetControlWithWait(name, timeoutMs, out var _))
            return Serialize(new { ok = false, error = $"Control not found: {name}", code = "XRAI_PANE_CONTROL_NOT_FOUND", timeout_ms = timeoutMs });

        return InvokeOnUI(() =>
        {
            if (!_controls.TryGet(name, out var ctrl))
                return Serialize(new { ok = false, error = $"Control not found: {name}", code = "XRAI_PANE_CONTROL_NOT_FOUND" });

            var selected = ctrl.SelectListItem(index, text);
            return Serialize(new
            {
                ok = true,
                name,
                selected_index = index,
                selected_text = selected,
            });
        });
    }

    // === New AI-agent commands ===

    private string HandleWaitControl(JsonNode node)
    {
        var name = node["name"]?.GetValue<string>()
            ?? throw new ArgumentException("wait_control requires 'name'");
        var targetValue = node["value"]?.GetValue<string>();
        bool? targetEnabled = node["enabled"] is JsonNode en ? en.GetValue<bool>() : null;
        bool? targetExists = node["exists"] is JsonNode ex ? ex.GetValue<bool>() : null;
        var timeout = node["timeout"]?.GetValue<int>() ?? 5000;
        var pollMs = node["poll_ms"]?.GetValue<int>() ?? 200;

        // Default condition when none specified: exists=true. This is what
        // callers almost always want ("wait for this control to appear") and
        // matches the documented pane.wait quick-reference example.
        if (targetValue == null && targetEnabled == null && targetExists == null)
            targetExists = true;

        var sw = Stopwatch.StartNew();
        while (sw.ElapsedMilliseconds < timeout)
        {
            // Each poll iteration dispatches to the UI thread individually,
            // so the UI remains responsive between polls.
            var pollResult = InvokeOnUI(() =>
            {
                bool found = _controls.TryGet(name, out var ctrl);

                // exists condition
                if (targetExists.HasValue)
                {
                    if (found != targetExists.Value)
                        return (matched: false, response: (string?)null);
                    if (!found)
                    {
                        // exists:false and not found — condition met, no further checks
                        return (matched: true, response: Serialize(new { ok = true, name, exists = false, elapsed_ms = sw.ElapsedMilliseconds }));
                    }
                }
                else if (!found)
                {
                    return (matched: false, response: (string?)null);
                }

                // At this point, ctrl is resolved
                if (targetValue != null)
                {
                    var val = ctrl!.GetValue() ?? "";
                    if (!val.Contains(targetValue, StringComparison.OrdinalIgnoreCase))
                        return (matched: false, response: (string?)null);
                }

                if (targetEnabled.HasValue)
                {
                    if (ctrl!.IsEnabled != targetEnabled.Value)
                        return (matched: false, response: (string?)null);
                }

                // All conditions met
                return (matched: true, response: Serialize(new
                {
                    ok = true,
                    name,
                    value = ctrl!.GetValue(),
                    enabled = ctrl.IsEnabled,
                    exists = true,
                    elapsed_ms = sw.ElapsedMilliseconds,
                }));
            });

            if (pollResult.matched)
                return pollResult.response!;

            Thread.Sleep(pollMs);
        }

        // Timeout — return final state for diagnostics
        return InvokeOnUI(() =>
        {
            bool found = _controls.TryGet(name, out var ctrl);
            return Serialize(new
            {
                ok = false,
                error = $"Timeout after {timeout}ms waiting for condition on '{name}'",
                name,
                exists = found,
                value = found ? ctrl!.GetValue() : null,
                enabled = found ? ctrl!.IsEnabled : (bool?)null,
                elapsed_ms = sw.ElapsedMilliseconds,
            });
        });
    }

    private string HandlePaneScreenshotCapture()
    {
        return InvokeOnUI(() =>
        {
            var element = _controls.RootElement;
            if (element == null)
                return Serialize(new { ok = false, error = "No root element exposed. Call Pilot.Expose() first." });

            int width = (int)element.ActualWidth;
            int height = (int)element.ActualHeight;
            if (width == 0 || height == 0)
                return Serialize(new { ok = false, error = $"Root element has zero size ({width}x{height}). Pane may not be visible." });

            var rtb = new RenderTargetBitmap(width, height, 96, 96, PixelFormats.Pbgra32);
            rtb.Render(element);

            var encoder = new PngBitmapEncoder();
            encoder.Frames.Add(BitmapFrame.Create(rtb));

            var tempPath = Path.Combine(Path.GetTempPath(), $"xrai_pane_{DateTime.Now:yyyyMMdd_HHmmss}.png");
            using (var fs = new FileStream(tempPath, FileMode.Create))
            {
                encoder.Save(fs);
            }

            return Serialize(new { ok = true, path = tempPath, width, height });
        });
    }

    private string HandleDrag(JsonNode node)
    {
        var fromName = node["from"]?.GetValue<string>()
            ?? throw new ArgumentException("drag requires 'from'");
        var toName = node["to"]?.GetValue<string>()
            ?? throw new ArgumentException("drag requires 'to'");

        return InvokeOnUI(() =>
        {
            if (!_controls.TryGet(fromName, out var fromCtrl))
                return Serialize(new { ok = false, error = $"Source control not found: {fromName}" });
            if (!_controls.TryGet(toName, out var toCtrl))
                return Serialize(new { ok = false, error = $"Target control not found: {toName}" });

            if (toCtrl.Element == null)
                return Serialize(new { ok = false, error = $"Target control '{toName}' has no WPF element (WinForms controls do not support drag)." });

            fromCtrl.DragTo(toCtrl.Element);
            return Serialize(new { ok = true, from = fromName, to = toName, dragged = true });
        });
    }

    /// <summary>
    /// Read recent lines from an XRai log file. Returns the file path, the
    /// last N lines, the total line count, and an exists flag.
    ///
    /// Sources:
    ///   pilot  → %LOCALAPPDATA%\XRai\logs\pilot-{pid}.log (this process)
    ///   daemon → %LOCALAPPDATA%\XRai\logs\daemon.log
    /// </summary>
    private string HandleLogRead(JsonNode node)
    {
        var source = node["source"]?.GetValue<string>() ?? "pilot";
        var lines = node["lines"]?.GetValue<int>() ?? 100;

        var dir = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
            "XRai", "logs");

        string? path = source switch
        {
            "pilot" => Path.Combine(dir, $"pilot-{Process.GetCurrentProcess().Id}.log"),
            "daemon" => Path.Combine(dir, "daemon.log"),
            _ => null
        };

        if (path == null)
            return Serialize(new
            {
                ok = false,
                error = $"Unknown log source: {source}",
                code = "XRAI_INVALID_ARGUMENT",
                valid = new[] { "pilot", "daemon" }
            });

        if (!File.Exists(path))
            return Serialize(new { ok = true, path, lines = Array.Empty<string>(), exists = false });

        try
        {
            var allLines = File.ReadAllLines(path);
            var tail = allLines.Length > lines ? allLines[^lines..] : allLines;
            return Serialize(new
            {
                ok = true,
                path,
                lines = tail,
                total_lines = allLines.Length,
                exists = true,
            });
        }
        catch (Exception ex)
        {
            return Serialize(new
            {
                ok = false,
                error = $"Failed to read log: {ex.Message}",
                code = "XRAI_INTERNAL_ERROR",
                path,
            });
        }
    }

    private string HandleContextMenu(JsonNode node)
    {
        var name = node["name"]?.GetValue<string>()
            ?? throw new ArgumentException("context_menu requires 'name'");
        var action = node["action"]?.GetValue<string>()
            ?? throw new ArgumentException("context_menu requires 'action'");

        return InvokeOnUI(() =>
        {
            if (!_controls.TryGet(name, out var ctrl))
                return Serialize(new { ok = false, error = $"Control not found: {name}" });

            switch (action)
            {
                case "read":
                    var items = ctrl.GetContextMenuItems();
                    return Serialize(new { ok = true, name, items });

                case "click":
                    var item = node["item"]?.GetValue<string>()
                        ?? throw new ArgumentException("context_menu action 'click' requires 'item'");
                    ctrl.ClickContextMenuItem(item);
                    return Serialize(new { ok = true, name, clicked_item = item });

                case "open":
                    ctrl.OpenContextMenu();
                    return Serialize(new { ok = true, name, context_menu_opened = true });

                default:
                    return Serialize(new { ok = false, error = $"Unknown context_menu action: {action}. Use 'read', 'click', or 'open'." });
            }
        });
    }

    private string InvokeOnUI(Func<string> action)
    {
        if (_uiDispatcher == null || _uiDispatcher.CheckAccess())
            return action();

        string result = "";
        _uiDispatcher.Invoke(() => result = action());
        return result;
    }

    private T InvokeOnUI<T>(Func<T> action)
    {
        if (_uiDispatcher == null || _uiDispatcher.CheckAccess())
            return action();

        T result = default!;
        _uiDispatcher.Invoke(() => result = action());
        return result;
    }

    private static string Serialize(object obj)
    {
        return JsonSerializer.Serialize(obj, new JsonSerializerOptions
        {
            DefaultIgnoreCondition = System.Text.Json.Serialization.JsonIgnoreCondition.WhenWritingNull,
        });
    }
}
