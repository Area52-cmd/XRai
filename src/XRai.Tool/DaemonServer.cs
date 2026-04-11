using System.Collections.Concurrent;
using System.Diagnostics;
using System.IO.Pipes;
using System.Security.Principal;
using System.Text.Json;
using System.Text.Json.Nodes;
using XRai.Com;
using XRai.Core;
using XRai.HooksClient;
using XRai.UI;
using XRai.Vision;

namespace XRai.Tool;

/// <summary>
/// Long-running XRai process that owns a single ExcelSession, HookConnection, and
/// CommandRouter and serves commands from multiple short-lived CLI clients over
/// a named pipe. Commands are serialized into a single queue and executed one at
/// a time — eliminating OLE-wait race conditions from rapid successive CLI calls.
///
/// Pipe name: xrai_daemon_{SanitizedUserName} — per-user, so multiple users on
/// the same machine don't collide.
///
/// Protocol: newline-delimited JSON. Client writes one command per line, daemon
/// writes one response per line. Same JSON shapes as the direct CLI mode.
/// </summary>
public class DaemonServer
{
    private readonly CommandRouter _router;
    private readonly ExcelSession _session;
    private readonly HookConnection _hookConnection;
    private ReloadOrchestrator _reloadOrchestrator = null!;
    private readonly EventStream _events;
    private readonly StaComWorker _staWorker;
    private readonly CancellationTokenSource _cts = new();
    private bool _running;

    /// <summary>
    /// Version hash embedded in the daemon binary. Clients compare their binary's
    /// hash against the running daemon's hash — if they differ, the daemon is STALE
    /// and auto-restarts before forwarding commands. This prevents the "I shipped
    /// a fix but the daemon is still running old code" problem.
    /// </summary>
    public static string BuildVersion => _buildVersion ??= ComputeBuildVersion();
    private static string? _buildVersion;

    private static string ComputeBuildVersion()
    {
        try
        {
            var asm = typeof(DaemonServer).Assembly;
            var loc = asm.Location;
            if (!string.IsNullOrEmpty(loc) && File.Exists(loc))
            {
                var fi = new FileInfo(loc);
                // Use file size + last write time as a fast version signature.
                // Any rebuild changes both.
                return $"{fi.Length:x}_{fi.LastWriteTimeUtc.Ticks:x}";
            }
        }
        catch { }
        return asm_hash_fallback();
        static string asm_hash_fallback() => typeof(DaemonServer).Assembly.GetName().Version?.ToString() ?? "unknown";
    }

    public static string PipeName
    {
        get
        {
            var user = WindowsIdentity.GetCurrent().Name;
            // Sanitize: pipe names can't contain backslashes
            var safe = user.Replace("\\", "_").Replace("/", "_").Replace(":", "_");
            return $"xrai_daemon_{safe}";
        }
    }

    /// <summary>
    /// Path of the on-disk daemon log. Surfaced via the status command and
    /// readable via {"cmd":"log.read","source":"daemon"}.
    /// Path: %LOCALAPPDATA%\XRai\logs\daemon.log
    /// </summary>
    public static string LogPath
    {
        get
        {
            var dir = Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                "XRai", "logs");
            Directory.CreateDirectory(dir);
            return Path.Combine(dir, "daemon.log");
        }
    }

    private static readonly object _logLock = new();
    private const long LogRotateBytes = 10L * 1024 * 1024;

    /// <summary>
    /// Append a line to the daemon log file. Mirrors to Console.WriteLine
    /// so the existing daemon stdout output is preserved. Never throws.
    /// </summary>
    public static void DaemonLog(string message)
    {
        var line = $"[{DateTime.UtcNow:o}] [daemon] {message}";
        Console.WriteLine(line);
        lock (_logLock)
        {
            try
            {
                var path = LogPath;
                try
                {
                    var fi = new FileInfo(path);
                    if (fi.Exists && fi.Length > LogRotateBytes)
                    {
                        var lines = File.ReadAllLines(path);
                        var keep = lines.Length / 2;
                        File.WriteAllLines(path, lines[keep..]);
                    }
                }
                catch { /* rotation must not throw */ }

                File.AppendAllText(path, line + Environment.NewLine);
            }
            catch
            {
                // logging must never throw
            }
        }
    }

    public DaemonServer()
    {
        _events = new EventStream(Console.Out);
        _router = new CommandRouter(_events);
        _session = new ExcelSession();
        _hookConnection = new HookConnection();

        // Create the STA worker FIRST — it owns the IOleMessageFilter and is
        // the only thread that should make COM calls. The router routes every
        // command through it.
        _staWorker = new StaComWorker();
        _router.SetStaInvoker((func, timeoutMs) => _staWorker.Invoke(func, timeoutMs));

        WireRouter();
    }

    /// <summary>
    /// When true, the daemon starts the XRai.Studio web dashboard on startup.
    /// Set by Program.cs when --studio is passed on the command line, or
    /// flipped via the {"cmd":"studio"} command at runtime.
    /// </summary>
    public bool StudioEnabled { get; set; }

    /// <summary>
    /// When true (default), the Studio host auto-opens the user's default
    /// browser to the dashboard URL when it starts. Smoke tests pass false
    /// via --studio --no-browser so headless test runs don't pop browser
    /// windows on the developer.
    /// </summary>
    public bool StudioLaunchBrowser { get; set; } = true;

    /// <summary>
    /// The live Studio host, or null if studio was never started in this
    /// daemon. Kept on the server so {"cmd":"studio"} can report its URL
    /// back to callers without having to restart the daemon.
    /// </summary>
    private XRai.Studio.StudioHost? _studioHost;

    /// <summary>
    /// Decorator wrapper for _router.Dispatch that publishes command.start /
    /// command.end events onto the Studio event bus. When Studio is not
    /// enabled, this is null and HandleClient calls _router.Dispatch directly.
    /// </summary>
    private XRai.Studio.Sources.RouterEventSource? _studioRouterEventSource;

    /// <summary>
    /// True if the daemon pipe was created with the restricted per-user ACL.
    /// Reported via security.status.
    /// </summary>
    public bool PipeAclRestricted { get; private set; }

    /// <summary>
    /// True if token-based authentication is enforced. False only when
    /// XRAI_ALLOW_UNAUTH=1 is set (legacy compatibility).
    /// </summary>
    public bool TokenAuthEnabled { get; private set; } = true;

    public int Run()
    {
        DaemonLog($"Starting on pipe: {PipeName}");
        DaemonLog("Use Ctrl+C to stop. Daemon serializes all commands through a single queue.");
        DaemonLog($"Log file: {LogPath}");
        Console.WriteLine();

        // Check if another daemon is already running on this pipe
        if (IsDaemonRunning())
        {
            Console.Error.WriteLine($"[xrai-daemon] ERROR: Another daemon is already running on pipe '{PipeName}'.");
            Console.Error.WriteLine($"[xrai-daemon] Stop it with: XRai.Tool.exe daemon-stop");
            return 1;
        }

        // Provision the auth token for this daemon's pipe. Stored under
        // %LOCALAPPDATA%\XRai\tokens\{PipeName}.token with a per-user ACL.
        try
        {
            PipeAuth.GenerateAndStoreToken(PipeName);
            TokenAuthEnabled = !PipeAuth.AllowUnauthenticated;
            if (PipeAuth.AllowUnauthenticated)
            {
                Console.Error.WriteLine("[xrai-daemon] WARNING: XRAI_ALLOW_UNAUTH=1 is set — daemon accepts unauthenticated clients. Do NOT use this outside of a trusted diagnostic environment.");
            }
        }
        catch (Exception ex)
        {
            Console.Error.WriteLine($"[xrai-daemon] WARNING: Failed to provision auth token: {ex.Message}. Daemon will start without token auth.");
            TokenAuthEnabled = false;
        }

        // Clean up tokens for any previously crashed XRai processes — best effort.
        try
        {
            PipeAuth.CleanupOrphanedTokens(name =>
            {
                // A pipe is considered alive if we can open it as a client.
                // For our own pipe, skip the liveness check (we're about to create it).
                if (name == PipeName) return true;
                try
                {
                    using var probe = new NamedPipeClientStream(".", name, PipeDirection.InOut);
                    probe.Connect(100);
                    return probe.IsConnected;
                }
                catch { return false; }
            });
        }
        catch { }

        _running = true;
        Console.CancelKeyPress += (_, e) => { e.Cancel = true; Stop(); };

        // Command execution is serialized through the StaComWorker (created in
        // the constructor). Each client handler thread calls _router.Dispatch
        // which routes through the worker's single-threaded STA queue.
        Console.WriteLine($"[xrai-daemon] STA worker ready (IOleMessageFilter registered: {_staWorker.FilterRegistered})");
        Console.WriteLine($"[xrai-daemon] Ready. Clients should invoke XRai.Tool.exe and stdin/stdout will be transparently forwarded.");

        // Start the Studio web dashboard if --studio was passed. Runs in the
        // daemon process alongside the pipe server. Non-fatal if it fails to
        // start — the daemon still serves the CLI.
        if (StudioEnabled)
        {
            try
            {
                StartStudio();
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine($"[xrai-daemon] Studio failed to start: {ex.Message}");
                DaemonLog($"Studio start failure: {ex}");
            }
        }

        Console.WriteLine();

        // Server loop: accept client connections on the pipe
        try
        {
            while (_running && !_cts.IsCancellationRequested)
            {
                NamedPipeServerStream? pipe = null;
                try
                {
                    // Create the daemon pipe with a restricted per-user ACL so
                    // no other local user can open it. Falls back to the
                    // unsecured constructor if ACL creation fails.
                    try
                    {
                        pipe = PipeAuth.CreateRestrictedServerPipe(
                            PipeName,
                            maxInstances: 10);
                        PipeAclRestricted = true;
                    }
                    catch (Exception aclEx)
                    {
                        Console.Error.WriteLine($"[xrai-daemon] WARNING: Could not create restricted pipe ACL ({aclEx.Message}). Falling back to unsecured pipe.");
                        pipe = new NamedPipeServerStream(
                            PipeName,
                            PipeDirection.InOut,
                            maxNumberOfServerInstances: 10,
                            PipeTransmissionMode.Byte,
                            PipeOptions.Asynchronous);
                        PipeAclRestricted = false;
                    }

                    // Wait for a client with cancellation support
                    var connectTask = pipe.WaitForConnectionAsync(_cts.Token);
                    connectTask.Wait(_cts.Token);

                    if (!_running) break;

                    // Handle this client on a dedicated thread — multiple clients
                    // can be connected simultaneously, but their commands all queue
                    // into the single executor, so actual execution remains serial.
                    var clientPipe = pipe;
                    pipe = null; // prevent disposal in finally
                    var clientThread = new Thread(() => HandleClient(clientPipe))
                    {
                        IsBackground = true,
                        Name = "xrai-daemon-client"
                    };
                    clientThread.Start();
                }
                catch (OperationCanceledException)
                {
                    break;
                }
                catch (Exception ex)
                {
                    Console.Error.WriteLine($"[xrai-daemon] Pipe accept error: {ex.Message}");
                    try { pipe?.Dispose(); } catch { }
                }
            }
        }
        finally
        {
            try { _studioHost?.Dispose(); } catch { }
            _studioHost = null;
            try { _staWorker.Dispose(); } catch { }
            _hookConnection.Disconnect();
            _session.Dispose();
        }

        Console.WriteLine("[xrai-daemon] Stopped.");
        return 0;
    }

    /// <summary>
    /// Boot the XRai.Studio web host inside this daemon. Wires the event bus
    /// to the command router (for command.start/end events) and starts the
    /// CaptureLoop pointed at the current Excel window (best-effort — the
    /// capture loop rechecks the hwnd every frame so it survives a rebuild).
    /// </summary>
    private void StartStudio()
    {
        if (_studioHost != null)
        {
            Console.WriteLine("[xrai-daemon] Studio already running: " + _studioHost.Url);
            return;
        }

        _studioHost = new XRai.Studio.StudioHost(
            stateProvider: BuildStudioState,
            commandDispatcher: (cmd, args) =>
            {
                // Forward dashboard-originated commands through the router.
                var obj = (System.Text.Json.Nodes.JsonObject)args.DeepClone();
                obj["cmd"] = cmd;
                return _router.Dispatch(obj.ToJsonString());
            });

        var url = _studioHost.Start(launchBrowser: StudioLaunchBrowser);

        // Display the URL banner prominently so RDP / headless / locked-down
        // users can copy-paste it into a browser even if the auto-launch
        // failed silently. ASCII-only so it renders in any codepage —
        // box-drawing characters render as garbage in cmd.exe's default CP437.
        Console.WriteLine();
        Console.WriteLine("==================================================================");
        Console.WriteLine("  XRai Studio is ready");
        Console.WriteLine("------------------------------------------------------------------");
        Console.WriteLine($"  {url}");
        Console.WriteLine();
        Console.WriteLine("  Open this URL in your browser to watch your code come alive.");
        Console.WriteLine("  The token in the URL is one-time and tied to this daemon.");
        Console.WriteLine("==================================================================");
        Console.WriteLine();

        // Log only the port — never write the token to disk. Anyone with
        // file access could otherwise grab the URL and authenticate.
        DaemonLog($"Studio started on http://127.0.0.1:{_studioHost.Port}/ (token redacted)");

        // ── Source: add-in pipe events ──────────────────────────────
        // Wire the add-in's in-process events (via hooks pipe) to the bus.
        var pipeSource = new XRai.Studio.Sources.PipeEventSource(_studioHost.Bus);
        _studioHost.RegisterDisposable(pipeSource);

        // ── Source: live screenshot stream ──────────────────────────
        var captureLoop = new XRai.Studio.Sources.CaptureLoop(
            _studioHost.Bus,
            () =>
            {
                try
                {
                    if (!_session.IsAttached) return (nint?)null;
                    nint hwnd = 0;
                    try
                    {
                        _staWorker.Invoke(() => { hwnd = (nint)_session.App.Hwnd; return "ok"; }, 1000);
                    }
                    catch { hwnd = 0; }
                    return hwnd == 0 ? null : hwnd;
                }
                catch { return null; }
            });
        captureLoop.Start();
        _studioHost.RegisterDisposable(captureLoop);

        // ── Source: AI coding agent transcript tail ─────────────────
        // Auto-detect which agent is running (Claude Code today, Codex/Cursor/
        // Aider in the future) and tail its transcript for live agent events.
        try
        {
            var agentAdapter = XRai.Studio.Sources.Agents.AgentAdapterFactory.Detect(_studioHost.Bus);
            agentAdapter.Start();
            _studioHost.RegisterDisposable(agentAdapter);
            DaemonLog($"Studio agent adapter: {agentAdapter.AgentName} (connected: {agentAdapter.IsConnected})");
        }
        catch (Exception agentEx)
        {
            DaemonLog($"Studio agent adapter failed to start: {agentEx.Message}");
        }

        // ── Source: project file watcher ────────────────────────────
        // Watches the current working directory for .cs/.xaml/etc edits and
        // publishes file.changed events. Pairs with the agent transcript:
        // the agent announces the edit intent (agent.tool.use), the file
        // watcher confirms the actual file landing.
        try
        {
            var cwd = Directory.GetCurrentDirectory();
            if (Directory.Exists(cwd))
            {
                var fileWatcher = new XRai.Studio.Sources.FileWatcherSource(_studioHost.Bus, cwd);
                fileWatcher.Start();
                _studioHost.RegisterDisposable(fileWatcher);
                DaemonLog($"Studio file watcher: {cwd}");
            }
        }
        catch (Exception fwEx)
        {
            DaemonLog($"Studio file watcher failed to start: {fwEx.Message}");
        }

        // ── Source: command router events ──────────────────────────
        // Wraps every router dispatch with command.start / command.end events.
        // The daemon's HandleClient already logs commands, but this source
        // gives the dashboard a real-time stream of what's being dispatched.
        _studioRouterEventSource = new XRai.Studio.Sources.RouterEventSource(_studioHost.Bus);
        _studioHost.RegisterDisposable(_studioRouterEventSource);

        // ── Rebuild progress instrumentation ────────────────────────
        _reloadOrchestrator.StepReporterFactory = () =>
            new TeeStepReporter((step, status, elapsedMs, detail) =>
            {
                var data = new System.Text.Json.Nodes.JsonObject
                {
                    ["step"] = step,
                    ["status"] = status,
                    ["elapsedMs"] = elapsedMs,
                };
                if (detail != null) data["detail"] = detail;
                _studioHost.Bus.Publish(XRai.Studio.StudioEvent.Now("rebuild.step", "daemon", data));
            });

        // ── Auto-attach watchdog ────────────────────────────────────
        // Background loop that polls every 2s for an Excel process and
        // attaches to it the moment one appears. Survives Excel kills/
        // relaunches: when the current attachment dies, the watchdog
        // re-attaches automatically. This is the difference between a
        // "demo" (user has to type {"cmd":"connect"}) and a shippable
        // product (user just opens Excel and Studio's screenshot panel
        // springs to life).
        StartAutoAttachWatchdog();
    }

    private CancellationTokenSource? _autoAttachCts;

    private void StartAutoAttachWatchdog()
    {
        _autoAttachCts = new CancellationTokenSource();
        var cts = _autoAttachCts;
        int consecutiveFailures = 0;
        bool staStuckNotified = false;

        // Exponential backoff when the STA is wedged. Normal tick = 2s.
        // When the STA is stuck (modal dialog, COM deadlock, etc.) we back
        // off to 30s — hammering a stuck STA every 2s doesn't heal it and
        // just floods the log with identical "Already attached" errors.
        const int NormalIntervalMs = 2000;
        const int StuckIntervalMs = 30_000;

        var thread = new Thread(() =>
        {
            DaemonLog("Auto-attach watchdog started");
            while (!cts.IsCancellationRequested && _running)
            {
                int sleepMs = NormalIntervalMs;

                try
                {
                    // Fast-path: if the STA worker is reporting itself stuck
                    // (from its own watchdog, not ours), DO NOT dispatch any
                    // more work to it. Queueing Detach/Attach onto a stuck
                    // queue just piles up abandoned work items and times out
                    // in order. Instead back off and let the user call
                    // {"cmd":"sta.reset"} to recover.
                    if (_staWorker.IsStuck)
                    {
                        if (!staStuckNotified)
                        {
                            staStuckNotified = true;
                            DaemonLog("Auto-attach: STA worker is stuck, backing off until sta.reset");
                            try
                            {
                                _studioHost?.Bus.Publish(XRai.Studio.StudioEvent.Now(
                                    "target.sta_stuck", "daemon", new System.Text.Json.Nodes.JsonObject
                                    {
                                        ["consecutiveTimeouts"] = _staWorker.ConsecutiveTimeouts,
                                        ["hint"] = "STA worker is wedged — likely a modal dialog. Auto-attach is paused. Run sta.reset to recover.",
                                    }));
                            }
                            catch { }
                        }
                        sleepMs = StuckIntervalMs;
                        continue; // skip to sleep/continue
                    }

                    // STA is healthy again after a stuck spell — clear the flag
                    // so the next stuck episode logs once again.
                    if (staStuckNotified)
                    {
                        staStuckNotified = false;
                        DaemonLog("Auto-attach: STA worker recovered, resuming normal polling");
                        try
                        {
                            _studioHost?.Bus.Publish(XRai.Studio.StudioEvent.Now(
                                "target.sta_recovered", "daemon", null));
                        }
                        catch { }
                    }

                    bool isAttached = _session.IsAttached;

                    if (isAttached)
                    {
                        // Health probe — short timeout to fail fast if the STA
                        // starts wedging. We do NOT try to detach here on
                        // probe failure because a probe timeout probably means
                        // the STA is about to be marked stuck, and dispatching
                        // more work would just queue behind the stuck op.
                        bool probeOk = false;
                        try
                        {
                            _staWorker.Invoke(() =>
                            {
                                var _h = _session.App.Hwnd;
                                return "ok";
                            }, 800);
                            probeOk = true;
                        }
                        catch (TimeoutException)
                        {
                            // STA just got wedged. The next iteration will
                            // see IsStuck=true and enter the backoff path.
                            DaemonLog("Auto-attach: health probe timed out, deferring to stuck-STA backoff");
                        }
                        catch
                        {
                            // Real COM exception — Excel process is gone.
                            probeOk = false;
                        }

                        if (!probeOk && !_staWorker.IsStuck)
                        {
                            // Safe to detach: the failure was a COM exception
                            // (process gone), not an STA wedge.
                            DaemonLog("Auto-attach: existing session is dead, detaching");
                            try { _staWorker.Invoke(() => { _session.Detach(); return "ok"; }, 1000); }
                            catch { /* if even this fails, let the next loop catch it */ }
                            try { _hookConnection.Disconnect(); } catch { }

                            try
                            {
                                _studioHost?.Bus.Publish(XRai.Studio.StudioEvent.Now(
                                    "target.detached", "daemon", new System.Text.Json.Nodes.JsonObject
                                    {
                                        ["reason"] = "process gone",
                                    }));
                            }
                            catch { }
                            isAttached = _session.IsAttached; // re-read after Detach
                        }
                    }

                    if (!isAttached && !_staWorker.IsStuck)
                    {
                        var procs = System.Diagnostics.Process.GetProcessesByName("EXCEL");
                        try
                        {
                            if (procs.Length > 0)
                            {
                                int targetPid = procs[0].Id;
                                try
                                {
                                    _staWorker.Invoke(() => { _session.Attach(); return "ok"; }, 5000);
                                    if (_session.IsAttached)
                                    {
                                        DaemonLog($"Auto-attach: bound to Excel pid={targetPid}");
                                        consecutiveFailures = 0;
                                        try { TryConnectHooks(); } catch { }
                                        try
                                        {
                                            _studioHost?.Bus.Publish(XRai.Studio.StudioEvent.Now(
                                                "target.attached", "daemon", new System.Text.Json.Nodes.JsonObject
                                                {
                                                    ["pid"] = targetPid,
                                                    ["name"] = "Microsoft Excel",
                                                }));
                                        }
                                        catch { }
                                    }
                                }
                                catch (Exception ex)
                                {
                                    consecutiveFailures++;
                                    if (consecutiveFailures % 10 == 1)
                                    {
                                        DaemonLog($"Auto-attach: pid={targetPid} not yet ready ({ex.Message})");
                                    }
                                }
                            }
                        }
                        finally
                        {
                            foreach (var p in procs) { try { p.Dispose(); } catch { } }
                        }
                    }
                }
                catch (Exception ex)
                {
                    DaemonLog($"Auto-attach watchdog iteration error: {ex.Message}");
                }

                try { Thread.Sleep(sleepMs); }
                catch { break; }
            }
            DaemonLog("Auto-attach watchdog stopped");
        })
        {
            IsBackground = true,
            Name = "xrai-studio-autoattach",
        };
        thread.Start();
    }

    /// <summary>
    /// Build the JSON object returned by GET /state. Called synchronously
    /// from the web request thread — must not block on long operations.
    /// Generic shape: the dashboard treats "target" as the currently-attached
    /// application. Today's target is Excel (the first adapter shipped); the
    /// structure allows future targets (Word, SAP, AutoCAD, any Windows app)
    /// to populate the same fields without schema changes.
    /// </summary>
    private System.Text.Json.Nodes.JsonObject BuildStudioState()
    {
        var obj = new System.Text.Json.Nodes.JsonObject
        {
            ["attached"] = _session.IsAttached,
            ["hooks"] = _hookConnection.IsConnected,
            ["daemonPipe"] = PipeName,
            ["studioUrl"] = _studioHost?.Url,
        };

        if (_session.IsAttached)
        {
            // ProbeWorkbookState touches the Application RCW; route through the
            // STA worker like every other COM access. Without this, a Kestrel
            // request thread could touch the RCW concurrently with the STA
            // thread executing a router command. Short timeout — if the STA
            // is busy, the dashboard sees no target info this tick (the
            // periodic /state poll picks it up next time).
            try
            {
                bool hasWorkbook = false;
                string? wbName = null;
                int wbCount = 0;
                _staWorker.Invoke(() =>
                {
                    var state = _session.ProbeWorkbookState();
                    hasWorkbook = state.HasWorkbook;
                    wbName = state.Name;
                    wbCount = state.Count;
                    return "ok";
                }, 1000);

                obj["target"] = new System.Text.Json.Nodes.JsonObject
                {
                    ["kind"] = "excel",
                    ["name"] = "Microsoft Excel",
                    ["document"] = wbName,
                    ["hasDocument"] = hasWorkbook,
                    ["documentCount"] = wbCount,
                };
            }
            catch
            {
                // STA busy or attached state changed mid-call — leave target
                // out of this snapshot rather than risking a race.
            }
        }

        return obj;
    }

    public void Stop()
    {
        Console.WriteLine("[xrai-daemon] Stop requested.");
        _running = false;
        _cts.Cancel();

        // Stop the auto-attach watchdog so it exits cleanly.
        try { _autoAttachCts?.Cancel(); } catch { }
        try { _autoAttachCts?.Dispose(); } catch { }
        _autoAttachCts = null;

        // Stop the Studio host first so the token file is cleared and the
        // Kestrel server shuts down cleanly before the process exits.
        try { _studioHost?.Dispose(); } catch { }
        _studioHost = null;

        // Connect and disconnect on our own pipe to unblock any pending accept
        try
        {
            using var kicker = new NamedPipeClientStream(".", PipeName, PipeDirection.InOut);
            kicker.Connect(500);
        }
        catch { }

        // Delete our token file so no stale token survives daemon shutdown.
        try { PipeAuth.ClearToken(PipeName); } catch { }
    }

    private void HandleClient(NamedPipeServerStream pipe)
    {
        try
        {
            using (pipe)
            using (var reader = new StreamReader(pipe))
            using (var writer = new StreamWriter(pipe) { AutoFlush = true })
            {
                bool authenticated = false;

                while (pipe.IsConnected && _running)
                {
                    string? line;
                    try { line = reader.ReadLine(); }
                    catch { break; }

                    if (line == null) break;
                    line = line.Trim();
                    if (line.Length == 0) continue;

                    // Control messages (__daemon_ping__ / __daemon_stop__) are
                    // authorised by the pipe ACL alone — they carry no payload
                    // and exist precisely so the client can probe liveness
                    // BEFORE it has the token. The ACL restricts them to the
                    // current user which is enough for these low-impact ops.
                    if (line == "__daemon_stop__")
                    {
                        writer.WriteLine(Response.Ok(new { daemon = "stopping" }));
                        Stop();
                        break;
                    }
                    if (line == "__daemon_ping__")
                    {
                        writer.WriteLine(Response.Ok(new
                        {
                            daemon = "alive",
                            pid = Environment.ProcessId,
                            pipe = PipeName,
                            filter_registered = _staWorker.FilterRegistered,
                            build_version = BuildVersion,
                            pipe_acl_restricted = PipeAclRestricted,
                            token_auth_enabled = TokenAuthEnabled,
                        }));
                        continue;
                    }

                    // Auth handshake: the first non-control line MUST be a
                    // valid {"auth_token":"..."} message unless XRAI_ALLOW_UNAUTH=1.
                    if (!authenticated)
                    {
                        var providedToken = PipeAuth.TryExtractAuthToken(line);
                        if (PipeAuth.ValidateToken(PipeName, providedToken))
                        {
                            authenticated = true;
                            writer.WriteLine(PipeAuth.BuildAuthOkResponse());
                            continue;
                        }

                        if (PipeAuth.AllowUnauthenticated)
                        {
                            Console.Error.WriteLine($"[xrai-daemon] WARNING: Accepting unauthenticated client because XRAI_ALLOW_UNAUTH=1 is set.");
                            authenticated = true;
                            // Fall through and dispatch the first line as a real command.
                        }
                        else
                        {
                            Console.Error.WriteLine($"[xrai-daemon] REJECTED unauthenticated client (token missing or invalid).");
                            try { writer.WriteLine(PipeAuth.BuildAuthFailedResponse()); } catch { }
                            break;
                        }
                    }

                    // Normal command — dispatch through the router, which routes
                    // through the StaComWorker. Multiple client handler threads
                    // may call Dispatch concurrently; the worker serializes their
                    // work via its single-threaded queue, so COM calls never race.
                    //
                    // When Studio is enabled, we route through RouterEventSource
                    // which publishes command.start / command.end events to the
                    // Studio bus alongside the normal dispatch. The wrapper is
                    // transparent to the CLI contract.
                    string response;
                    try
                    {
                        if (_studioRouterEventSource != null)
                            response = _studioRouterEventSource.WrapDispatch(line, _router.Dispatch);
                        else
                            response = _router.Dispatch(line);
                    }
                    catch (Exception ex)
                    {
                        response = Response.Error($"Dispatch exception: {ex.GetType().Name}: {ex.Message}");
                    }

                    try { writer.WriteLine(response); }
                    catch { break; }
                }
            }
        }
        catch (Exception ex)
        {
            Console.Error.WriteLine($"[xrai-daemon] Client handler error: {ex.Message}");
        }
    }

    // ── Router wiring — mirrors Program.cs Main() setup ──────────────

    private void WireRouter()
    {
        _router.DefaultTimeoutMs = 15000;

        _router.Register("wait", cmdArgs =>
        {
            var ms = cmdArgs["ms"]?.GetValue<int>();
            if (ms.HasValue)
            {
                Thread.Sleep(Math.Max(0, ms.Value));
                return Response.Ok(new { slept_ms = ms.Value });
            }

            if (_session.IsAttached)
            {
                var existingState = _session.ProbeWorkbookState();
                return Response.Ok(new
                {
                    attached = true,
                    already_attached = true,
                    version = _session.ExcelVersion,
                    hooks = _hookConnection.IsConnected,
                    has_workbook = existingState.HasWorkbook,
                    active_workbook = existingState.Name,
                    workbook_count = existingState.Count,
                    daemon = true
                });
            }

            _session.WaitAndAttach();
            TryConnectHooks();
            var state = _session.ProbeWorkbookState();
            return Response.Ok(new
            {
                attached = true,
                version = _session.ExcelVersion,
                hooks = _hookConnection.IsConnected,
                has_workbook = state.HasWorkbook,
                active_workbook = state.Name,
                workbook_count = state.Count,
                daemon = true
            });
        });

        _router.Register("attach", args =>
        {
            var attachPid = args["pid"]?.GetValue<int>();
            _session.Attach(attachPid);
            TryConnectHooks();
            var state = _session.ProbeWorkbookState();
            return Response.Ok(new
            {
                attached = true,
                version = _session.ExcelVersion,
                hooks = _hookConnection.IsConnected,
                has_workbook = state.HasWorkbook,
                active_workbook = state.Name,
                workbook_count = state.Count,
                daemon = true
            });
        });

        _router.Register("detach", _ =>
        {
            _hookConnection.Disconnect();
            _session.Detach();
            return Response.Ok(new { detached = true, daemon = true });
        });

        _router.Register("status", _ =>
        {
            var studioInfo = _studioHost != null ? new
            {
                running = true,
                url = _studioHost.Url,
                port = _studioHost.Port,
            } : null;

            if (!_session.IsAttached)
            {
                return Response.Ok(new
                {
                    attached = false,
                    daemon = true,
                    daemon_pipe = PipeName,
                    log_path = LogPath,
                    studio = studioInfo,
                    hint = "Daemon is running but no COM session. Call {\"cmd\":\"connect\"} to attach."
                });
            }
            var state = _session.ProbeWorkbookState();
            return Response.Ok(new
            {
                attached = true,
                version = _session.ExcelVersion,
                hooks = _hookConnection.IsConnected,
                hooks_pipe = _hookConnection.PipeName,
                has_workbook = state.HasWorkbook,
                active_workbook = state.Name,
                workbook_count = state.Count,
                daemon = true,
                daemon_pipe = PipeName,
                log_path = LogPath,
                filter_registered = _staWorker.FilterRegistered,
                studio = studioInfo,
            });
        });

        _router.Register("ensure.workbook", _ =>
        {
            var state = _session.ProbeWorkbookState();
            if (state.HasWorkbook)
                return Response.Ok(new { already_open = true, name = state.Name, created = false });
            var wb = _session.EnsureWorkbook();
            string name = wb.Name;
            System.Runtime.InteropServices.Marshal.ReleaseComObject(wb);
            return Response.Ok(new { already_open = false, name, created = true });
        });

        // Studio web dashboard — lazy-start on demand. If the daemon was
        // launched with --studio the host is already running; this command
        // returns its URL + token. Otherwise, it boots the host now and
        // returns the freshly-minted URL. Idempotent.
        _router.Register("studio", _ =>
        {
            try
            {
                if (_studioHost == null)
                {
                    StudioEnabled = true;
                    StartStudio();
                }
                return Response.Ok(new
                {
                    running = _studioHost != null,
                    url = _studioHost?.Url,
                    port = _studioHost?.Port,
                    token_file = XRai.Studio.StudioToken.GetTokenFilePath(),
                    hint = "Open the url in a browser. The token is embedded in the query string and converted to a cookie on first load.",
                });
            }
            catch (Exception ex)
            {
                return Response.ErrorFromException(ex, "studio");
            }
        });

        _router.Register("sta.reset", _ =>
        {
            bool wasStuck = _staWorker.IsStuck;
            int timeouts = _staWorker.ConsecutiveTimeouts;
            try
            {
                try { _session.Detach(); } catch { }
                try { _hookConnection.Disconnect(); } catch { }
                _staWorker.Reset();
                return Response.Ok(new
                {
                    reset = true,
                    was_stuck = wasStuck,
                    consecutive_timeouts_before_reset = timeouts,
                    filter_registered = _staWorker.FilterRegistered,
                    hint = "STA thread recycled. Run {\"cmd\":\"connect\"} to reattach.",
                });
            }
            catch (Exception ex) { return Response.ErrorFromException(ex, "sta.reset"); }
        });

        _router.Register("sta.status", _ => Response.Ok(new
        {
            is_alive = _staWorker.IsAlive,
            is_stuck = _staWorker.IsStuck,
            filter_registered = _staWorker.FilterRegistered,
            consecutive_timeouts = _staWorker.ConsecutiveTimeouts,
            last_timeout_at = _staWorker.LastTimeoutAt?.ToString("o"),
        }));

        _router.Register("connect", args =>
        {
            var timeoutMs = args["timeout"]?.GetValue<int>() ?? 30000;
            try
            {
                if (!_session.IsAttached) _session.WaitAndAttach(timeoutMs);
            }
            catch (Exception ex) { return Response.Error($"Failed to attach: {ex.Message}"); }

            var state = _session.ProbeWorkbookState();
            bool createdWorkbook = false;
            if (!state.HasWorkbook)
            {
                try
                {
                    var wb = _session.EnsureWorkbook();
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(wb);
                    createdWorkbook = true;
                    state = _session.ProbeWorkbookState();
                }
                catch (Exception ex) { return Response.Error($"Attached but could not create workbook: {ex.Message}"); }
            }

            TryConnectHooks();

            return Response.Ok(new
            {
                attached = true,
                version = _session.ExcelVersion,
                hooks = _hookConnection.IsConnected,
                hooks_pipe = _hookConnection.PipeName,
                active_workbook = state.Name,
                workbook_count = state.Count,
                created_workbook = createdWorkbook,
                daemon = true
            });
        });

        // Phase 1: Core COM
        new CellOps(_session).Register(_router);
        new SheetOps(_session).Register(_router);
        new CalcOps(_session).Register(_router);
        var dialogDriver = new Win32DialogDriver();
        _router.SetTimeoutDiagnostics(dialogDriver);
        new WorkbookOps(_session, dialogDriver).Register(_router);

        // Phase A: Expanded COM
        new LayoutOps(_session).Register(_router);
        new DataOps(_session).Register(_router);
        new FormatOps(_session).Register(_router);
        new ChartOps(_session).Register(_router);
        new SparklineOps(_session).Register(_router);
        new TableOps(_session).Register(_router);
        new FilterOps(_session).Register(_router);
        new PivotOps(_session).Register(_router);
        new PrintOps(_session).Register(_router);
        new WindowOps(_session).Register(_router);
        new ShapeOps(_session).Register(_router);
        new AdvancedOps(_session).Register(_router);
        new PowerQueryOps(_session).Register(_router);
        new VbaOps(_session).Register(_router);
        new SlicerOps(_session).Register(_router);
        new ConnectionOps(_session).Register(_router);

        // Desktop automation (app-agnostic)
        new DesktopOps().Register(_router);
        new AppAttachOps().Register(_router);

        // Phase 2: Hooks
        new PaneClient(_hookConnection).Register(_router);
        new ModelClient(_hookConnection).Register(_router);

        // Phase 3: FlaUI + Vision
        // RibbonDriver takes the shared Win32 dialog driver for dialog.click
        // / dialog.dismiss Win32 fallback (closes the dialog.click ↔
        // win32.dialog.list desync).
        new RibbonDriver(dialogDriver).Register(_router);
        dialogDriver.Register(_router);
        new FileDialogDriver().Register(_router);
        var capture = new Capture();
        capture.SetComHwndProvider(() =>
        {
            try { return _session.IsAttached ? (nint)_session.App.Hwnd : null; }
            catch { return null; }
        });
        capture.Register(_router);
        new DiffOps(capture).Register(_router);
        new OcrOps().Register(_router);

        // Phase 3b: Intelligent Waits
        new WaitOps().Register(_router);

        // Phase 3c: Test Reporting + Assertions
        new TestReporter().Register(_router);
        new AssertOps(_router).Register(_router);

        // Phase 4: Reload + meta
        _reloadOrchestrator = new ReloadOrchestrator(_session, _hookConnection);
        _reloadOrchestrator.Register(_router);

        _router.Register("excel.kill", _ =>
        {
            _hookConnection.Disconnect();
            try { _session.Detach(); } catch { }

            var procs = Process.GetProcessesByName("EXCEL");
            var killed = new List<int>();
            foreach (var p in procs)
            {
                try
                {
                    p.Kill(entireProcessTree: true);
                    p.WaitForExit(5000);
                    killed.Add(p.Id);
                }
                catch { }
            }
            Thread.Sleep(500);
            var remaining = Process.GetProcessesByName("EXCEL").Length;
            return Response.Ok(new { killed_pids = killed, killed_count = killed.Count, remaining });
        });

        _router.Register("help", _ => Response.Ok(new
        {
            command_count = _router.RegisteredCommands.Count(),
            commands = _router.RegisteredCommands.ToArray(),
            daemon = true,
            daemon_pipe = PipeName
        }));

        _router.Register("commands", _ => Response.Ok(new { commands = _router.RegisteredCommands.ToArray() }));

        _router.Register("security.status", _ =>
        {
            string currentUser;
            try { currentUser = WindowsIdentity.GetCurrent().Name; }
            catch { currentUser = "unknown"; }

            var daemonTokenPath = PipeAuth.GetTokenFilePath(PipeName);
            bool daemonTokenExists = false;
            try { daemonTokenExists = File.Exists(daemonTokenPath); } catch { }

            // Hooks pipe security — reach across to the hooks pipe if connected.
            string? hooksPipeName = _hookConnection.PipeName;
            bool? hooksPipeAcl = null;
            bool? hooksTokenAuth = null;
            bool? hooksTokenExists = null;
            string? hooksTokenPath = null;
            if (!string.IsNullOrEmpty(hooksPipeName))
            {
                hooksTokenPath = PipeAuth.GetTokenFilePath(hooksPipeName);
                try { hooksTokenExists = File.Exists(hooksTokenPath); } catch { }

                // Query the hooks side for its self-reported security state.
                try
                {
                    var hooksStatus = _hookConnection.SendCommand("security.status");
                    if (hooksStatus != null)
                    {
                        hooksPipeAcl = hooksStatus["pipe_acl_restricted"]?.GetValue<bool>();
                        hooksTokenAuth = hooksStatus["token_auth_enabled"]?.GetValue<bool>();
                    }
                }
                catch { }
            }

            return Response.Ok(new
            {
                pipe_acl_restricted = PipeAclRestricted,
                token_auth_enabled = TokenAuthEnabled,
                token_file_exists = daemonTokenExists,
                token_file_path = daemonTokenPath,
                hooks_pipe_name = hooksPipeName,
                hooks_pipe_acl_restricted = hooksPipeAcl,
                hooks_token_auth_enabled = hooksTokenAuth,
                hooks_token_file_exists = hooksTokenExists,
                hooks_token_file_path = hooksTokenPath,
                daemon_pipe_name = PipeName,
                current_user = currentUser,
                allow_unauthenticated = PipeAuth.AllowUnauthenticated,
            });
        });
    }

    private void TryConnectHooks()
    {
        try
        {
            var processes = Process.GetProcessesByName("EXCEL");
            if (processes.Length > 0)
                _hookConnection.Connect(processes[0].Id, 2000);
        }
        catch { /* Hooks are optional */ }
    }

    public static bool IsDaemonRunning()
    {
        try
        {
            using var client = new NamedPipeClientStream(".", PipeName, PipeDirection.InOut);
            client.Connect(200);
            return client.IsConnected;
        }
        catch
        {
            return false;
        }
    }

}
