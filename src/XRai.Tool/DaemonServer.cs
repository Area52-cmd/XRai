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
            try { _staWorker.Dispose(); } catch { }
            _hookConnection.Disconnect();
            _session.Dispose();
        }

        Console.WriteLine("[xrai-daemon] Stopped.");
        return 0;
    }

    public void Stop()
    {
        Console.WriteLine("[xrai-daemon] Stop requested.");
        _running = false;
        _cts.Cancel();

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
                    string response;
                    try
                    {
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
            if (!_session.IsAttached)
            {
                return Response.Ok(new
                {
                    attached = false,
                    daemon = true,
                    daemon_pipe = PipeName,
                    log_path = LogPath,
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
        new ReloadOrchestrator(_session, _hookConnection).Register(_router);

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
