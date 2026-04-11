using XRai.Com;
using XRai.Core;
using XRai.HooksClient;
using XRai.Studio;
using XRai.UI;
using XRai.Vision;
using XRai.Tool;

class Program
{
    [STAThread]
    static int Main(string[] args)
    {
        int? pid = null;
        bool wait = false;
        int defaultTimeoutMs = 15000;
        bool repl = false;
        bool forceDirect = false;

        for (int i = 0; i < args.Length; i++)
        {
            switch (args[i])
            {
                case "--pid" when i + 1 < args.Length:
                    pid = int.Parse(args[++i]);
                    break;
                case "--wait":
                    wait = true;
                    break;
                case "--repl":
                    repl = true;
                    break;
                case "--no-daemon":
                    // Force direct mode even if a daemon is running (for debugging)
                    forceDirect = true;
                    break;
                case "--timeout" when i + 1 < args.Length:
                    defaultTimeoutMs = int.Parse(args[++i]);
                    break;
                case "--help" or "-h" or "help":
                    return PrintHelp();
                case "--daemon" or "daemon":
                    // Run as the long-lived daemon process
                    return new DaemonServer().Run();
                case "--studio" or "studio":
                    // Run as daemon with the Studio web dashboard enabled.
                    // Browser launches automatically by default. Pass
                    // --no-browser to suppress (used by smoke tests / RDP).
                    return new DaemonServer
                    {
                        StudioEnabled = true,
                        StudioLaunchBrowser = !args.Contains("--no-browser"),
                    }.Run();
                case "daemon-status":
                    return DaemonStatus();
                case "daemon-stop":
                    return DaemonStop();
                case "doctor":
                    return DoctorCommand.Run();
                case "init":
                    return InitCommand.Run(args[i..]);
                case "kill-excel":
                    return KillExcel();
                case "dump-commands" or "--dump-commands":
                    // Build router (no attach), print all registered commands, exit
                    return DumpCommands();
                case "set-ide":
                    return SetIdeCommand(args, i);
                case "get-ide":
                    return GetIdeCommand();
                case "ides":
                    return ListIdesCommand();
            }
        }

        // === Auto-detect daemon: if a daemon is running, forward stdin/stdout
        // through it. This eliminates per-invocation COM attach cost and prevents
        // OLE races from rapid successive CLI calls. Can be disabled with --no-daemon.
        //
        // DaemonClient.Run() returns -1 as a sentinel when it detects that the
        // running daemon is STALE (build_version mismatches the local binary).
        // In that case it auto-stops the stale daemon and we fall through to
        // direct mode so the caller gets fresh code.
        if (!forceDirect && !wait && !pid.HasValue && DaemonServer.IsDaemonRunning())
        {
            var daemonExit = DaemonClient.Run();
            if (daemonExit != -1) return daemonExit;
            // -1 means stale daemon was killed; continue to direct mode
        }

        var events = new EventStream(Console.Out);
        var router = new CommandRouter(events) { DefaultTimeoutMs = defaultTimeoutMs };

        // Create a dedicated STA worker thread for COM operations. It registers
        // IOleMessageFilter and owns a Windows message pump, which is what
        // allows Excel's busy-state rejections to retry silently instead of
        // popping the "Excel is waiting for another application" dialog.
        // The router routes every command through this worker via SetStaInvoker.
        var staWorker = new StaComWorker();
        router.SetStaInvoker((func, timeoutMs) => staWorker.Invoke(func, timeoutMs));

        var session = new ExcelSession();
        var hookConnection = new HookConnection();

        // Auto-recover the STA thread when it gets stuck, before the user has
        // to run {"cmd":"sta.reset"} by hand. The router invokes this callback
        // on TimeoutException and retries the command once against the fresh
        // thread if we return true.
        router.StaAutoRecover = () =>
        {
            try
            {
                // Detach session + hooks — the old COM references are bound
                // to the dead apartment and will be invalid after recycle.
                try { session.Detach(); } catch { }
                try { hookConnection.Disconnect(); } catch { }
                staWorker.Reset();

                // Best-effort reattach so the retried command has something
                // to work with. Failures here are swallowed — the retry will
                // surface its own clearer error.
                try { session.Attach(); } catch { }
                return true;
            }
            catch
            {
                return false;
            }
        };

        // Forward-declare drivers so `connect` / `sta.reset` handlers can
        // reference them in their closures (actual instantiation happens
        // further down, after Phase 1 Ops registration).
        RibbonDriver? ribbonDriver = null;

        // === Connection commands ===
        // wait: dual-mode command.
        //   {"cmd":"wait","ms":500}  → inert sleep, no COM touch, batch-safe
        //   {"cmd":"wait"}            → wait for Excel to appear, then attach
        //                               (legacy behavior, NOT safe inside a batch)
        router.Register("wait", cmdArgs =>
        {
            var ms = cmdArgs["ms"]?.GetValue<int>();
            if (ms.HasValue)
            {
                // Pure sleep mode — no COM, no attach. Safe in any batch.
                Thread.Sleep(Math.Max(0, ms.Value));
                return Response.Ok(new { slept_ms = ms.Value });
            }

            // Legacy: wait-and-attach. Only safe when not already attached.
            if (session.IsAttached)
            {
                // Already attached — treat as a no-op instead of throwing.
                // This eliminates the "Already attached. Call Detach() first."
                // cascade failure when wait is used inside a batch after connect.
                var existingState = session.ProbeWorkbookState();
                return Response.Ok(new
                {
                    attached = true,
                    already_attached = true,
                    version = session.ExcelVersion,
                    hooks = hookConnection.IsConnected,
                    has_workbook = existingState.HasWorkbook,
                    active_workbook = existingState.Name,
                    workbook_count = existingState.Count
                });
            }

            session.WaitAndAttach();
            TryConnectHooks(hookConnection);
            var state = session.ProbeWorkbookState();
            return Response.Ok(new
            {
                attached = true,
                version = session.ExcelVersion,
                hooks = hookConnection.IsConnected,
                has_workbook = state.HasWorkbook,
                active_workbook = state.Name,
                workbook_count = state.Count
            });
        });

        router.Register("attach", cmdArgs =>
        {
            var attachPid = cmdArgs["pid"]?.GetValue<int>();
            session.Attach(attachPid);
            TryConnectHooks(hookConnection);
            var state = session.ProbeWorkbookState();
            return Response.Ok(new
            {
                attached = true,
                version = session.ExcelVersion,
                hooks = hookConnection.IsConnected,
                has_workbook = state.HasWorkbook,
                active_workbook = state.Name,
                workbook_count = state.Count
            });
        });

        router.Register("detach", _ =>
        {
            hookConnection.Disconnect();
            session.Detach();
            return Response.Ok(new { detached = true });
        });

        router.Register("status", _ =>
        {
            if (!session.IsAttached)
            {
                return Response.Ok(new
                {
                    attached = false,
                    version = (string?)null,
                    hooks = false,
                    hint = "Not attached. Call {\"cmd\":\"attach\"} or {\"cmd\":\"wait\"} first."
                });
            }
            var state = session.ProbeWorkbookState();
            var hooksInfo = ProbeHooksVersion(hookConnection);
            var hooksStaleWarning = CheckHooksStaleness(hooksInfo);

            return Response.Ok(new
            {
                attached = true,
                version = session.ExcelVersion,
                hooks = hookConnection.IsConnected,
                hooks_pipe = hookConnection.PipeName,
                hooks_assembly_version = hooksInfo.AssemblyVersion,
                hooks_library_build_timestamp = hooksInfo.LibraryBuildTimestamp,
                addin_build_timestamp = hooksInfo.AddinBuildTimestamp,
                hooks_build_timestamp = hooksInfo.BuildTimestamp, // DEPRECATED, use hooks_library_build_timestamp
                hooks_stale = hooksStaleWarning != null,
                hooks_stale_warning = hooksStaleWarning,
                has_workbook = state.HasWorkbook,
                active_workbook = state.Name,
                workbook_count = state.Count,
                hint = state.HasWorkbook
                    ? hooksStaleWarning
                    : "Excel is on start screen with no workbook. Call {\"cmd\":\"ensure.workbook\"} or open a workbook first."
            });
        });

        router.Register("ensure.workbook", _ =>
        {
            var state = session.ProbeWorkbookState();
            if (state.HasWorkbook)
                return Response.Ok(new { already_open = true, name = state.Name, created = false });

            var wb = session.EnsureWorkbook();
            string name = wb.Name;
            System.Runtime.InteropServices.Marshal.ReleaseComObject(wb);
            return Response.Ok(new { already_open = false, name, created = true });
        });

        // STA worker recycle — recovers from a stuck STA thread without
        // requiring a full process restart. The old thread is abandoned
        // (leaks until process exit) and a fresh one is spawned with a new
        // IOleMessageFilter. Session state is NOT preserved — caller should
        // run {"cmd":"connect"} after reset to reattach.
        router.Register("sta.reset", _ =>
        {
            bool wasStuck = staWorker.IsStuck;
            int timeouts = staWorker.ConsecutiveTimeouts;
            try
            {
                // Detach the session first — the old COM reference is bound
                // to the old STA apartment and will be invalid after reset.
                try { session.Detach(); } catch { }
                try { hookConnection.Disconnect(); } catch { }

                staWorker.Reset();

                return Response.Ok(new
                {
                    reset = true,
                    was_stuck = wasStuck,
                    consecutive_timeouts_before_reset = timeouts,
                    filter_registered = staWorker.FilterRegistered,
                    hint = "STA thread recycled. Run {\"cmd\":\"connect\"} to reattach to Excel.",
                });
            }
            catch (Exception ex)
            {
                return Response.ErrorFromException(ex, "sta.reset");
            }
        });

        router.Register("sta.status", _ =>
        {
            return Response.Ok(new
            {
                is_alive = staWorker.IsAlive,
                is_stuck = staWorker.IsStuck,
                filter_registered = staWorker.FilterRegistered,
                consecutive_timeouts = staWorker.ConsecutiveTimeouts,
                last_timeout_at = staWorker.LastTimeoutAt?.ToString("o"),
            });
        });

        router.Register("connect", cmdArgs =>
        {
            // Batteries-included: wait for Excel, attach, ensure workbook, connect hooks
            var timeoutMs = cmdArgs["timeout"]?.GetValue<int>() ?? 30000;
            try
            {
                if (!session.IsAttached)
                    session.WaitAndAttach(timeoutMs);
            }
            catch (Exception ex)
            {
                return Response.Error($"Failed to attach: {ex.Message}");
            }

            // Ensure a workbook exists (auto-create if Excel is on start screen)
            var state = session.ProbeWorkbookState();
            bool createdWorkbook = false;
            if (!state.HasWorkbook)
            {
                try
                {
                    var wb = session.EnsureWorkbook();
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(wb);
                    createdWorkbook = true;
                    state = session.ProbeWorkbookState();
                }
                catch (Exception ex)
                {
                    return Response.Error($"Attached but could not create workbook: {ex.Message}");
                }
            }

            // Connect hooks (optional — no error if unavailable)
            TryConnectHooks(hookConnection);

            // Invalidate cached UIA/FlaUI state so ribbon/dialog commands
            // don't return stale empty results after a reattach
            try { ribbonDriver?.InvalidateCache(); } catch { }

            var hooksInfo = ProbeHooksVersion(hookConnection);
            var hooksStaleWarning = CheckHooksStaleness(hooksInfo);

            // Structured hooks diagnostics — included only if hooks failed to connect
            object? hooksDiagnostics = null;
            if (!hookConnection.IsConnected)
            {
                try { hooksDiagnostics = GetHooksDiagnostics(hookConnection); } catch { }
            }

            return Response.Ok(new
            {
                attached = true,
                version = session.ExcelVersion,
                hooks = hookConnection.IsConnected,
                hooks_pipe = hookConnection.PipeName,
                hooks_assembly_version = hooksInfo.AssemblyVersion,
                hooks_library_build_timestamp = hooksInfo.LibraryBuildTimestamp,
                addin_build_timestamp = hooksInfo.AddinBuildTimestamp,
                hooks_build_timestamp = hooksInfo.BuildTimestamp, // DEPRECATED, use hooks_library_build_timestamp
                hooks_stale = hooksStaleWarning != null,
                hooks_stale_warning = hooksStaleWarning,
                hooks_diagnostics = hooksDiagnostics,
                active_workbook = state.Name,
                workbook_count = state.Count,
                created_workbook = createdWorkbook
            });
        });

        // Shared dialog watchdog — one instance, injected into every op that
        // needs dialog dismissal during COM calls (WorkbookOps.Open, etc.)
        // and also registered on the router so excel.autodismiss / win32.dialog.*
        // commands share the same state.
        var dialogDriver = new Win32DialogDriver();
        router.SetTimeoutDiagnostics(dialogDriver);

        // === Phase 1: Core COM operations ===
        new CellOps(session).Register(router);
        new SheetOps(session).Register(router);
        new CalcOps(session).Register(router);
        new WorkbookOps(session, dialogDriver).Register(router);

        // === Phase A: Expanded COM operations ===
        new LayoutOps(session).Register(router);
        new DataOps(session).Register(router);
        new FormatOps(session).Register(router);
        new ChartOps(session).Register(router);
        new SparklineOps(session).Register(router);
        new TableOps(session).Register(router);
        new FilterOps(session).Register(router);
        new PivotOps(session).Register(router);
        new PrintOps(session).Register(router);
        new WindowOps(session).Register(router);
        new ShapeOps(session).Register(router);
        new AdvancedOps(session).Register(router);
        new PowerQueryOps(session).Register(router);
        new VbaOps(session).Register(router);
        new SlicerOps(session).Register(router);
        new ConnectionOps(session).Register(router);

        // === Desktop automation (app-agnostic) ===
        new DesktopOps().Register(router);
        new AppAttachOps().Register(router);

        // === Phase 2: Hooks ===
        new PaneClient(hookConnection).Register(router);
        new ModelClient(hookConnection).Register(router);

        // === Phase 3: FlaUI + Vision ===
        // RibbonDriver takes the shared Win32DialogDriver so dialog.click /
        // dialog.dismiss can fall through to Win32 EnumWindows when UIA's
        // ModalWindows enumeration finds nothing (closes the dialog.click ↔
        // win32.dialog.list desync on top-level #32770 / NUIDialog windows).
        ribbonDriver = new RibbonDriver(dialogDriver);
        ribbonDriver.Register(router);
        dialogDriver.Register(router);
        new FileDialogDriver().Register(router);
        var capture = new Capture();
        // Wire COM Hwnd provider so screenshot targets the correct XLMAIN window
        // (the one with the active workbook, not the blank start screen)
        capture.SetComHwndProvider(() =>
        {
            try { return session.IsAttached ? (nint)session.App.Hwnd : null; }
            catch { return null; }
        });
        capture.Register(router);
        new DiffOps(capture).Register(router);
        new OcrOps().Register(router);

        // === Phase 3b: Intelligent Waits ===
        new WaitOps().Register(router);

        // === Phase 3c: Test Reporting + Assertions ===
        new TestReporter().Register(router);
        new AssertOps(router).Register(router);

        // === Phase 4: Reload + Meta ===
        new ReloadOrchestrator(session, hookConnection).Register(router);

        // === kill-excel as JSON command (not just CLI) ===
        router.Register("excel.kill", _ =>
        {
            hookConnection.Disconnect();

            // Graceful quit first — prevents Document Recovery on next launch
            bool graceful = false;
            try
            {
                if (session.IsAttached)
                {
                    session.App.DisplayAlerts = false;
                    session.App.Quit();
                    graceful = true;
                }
            }
            catch { }
            try { session.Detach(); } catch { }
            if (graceful) Thread.Sleep(2000);

            // Force-kill any survivors
            var procs = System.Diagnostics.Process.GetProcessesByName("EXCEL");
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

            // Clean recovery files
            try
            {
                var recoveryDir = Path.Combine(
                    Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
                    "Microsoft", "Excel");
                if (Directory.Exists(recoveryDir))
                {
                    foreach (var f in Directory.GetFiles(recoveryDir, "*.xlsb", SearchOption.AllDirectories)
                        .Concat(Directory.GetFiles(recoveryDir, "*.tmp", SearchOption.AllDirectories)))
                    {
                        try { File.Delete(f); } catch { }
                    }
                }
            }
            catch { }

            var remaining = System.Diagnostics.Process.GetProcessesByName("EXCEL").Length;
            return Response.Ok(new { graceful, killed_pids = killed, killed_count = killed.Count, remaining });
        });

        // === Introspection commands ===
        router.Register("help", _ =>
        {
            var commands = router.RegisteredCommands.ToArray();
            return Response.Ok(new
            {
                command_count = commands.Length,
                commands,
                cli_subcommands = new[] { "doctor", "init", "kill-excel", "dump-commands", "help" },
                cli_flags = new[] { "--pid", "--wait", "--repl", "--timeout <ms>" }
            });
        });

        router.Register("commands", _ =>
        {
            return Response.Ok(new { commands = router.RegisteredCommands.ToArray() });
        });

        // === Auto-attach on startup ===
        // CRITICAL: attach MUST run on the STA worker so the COM reference is
        // owned by the worker's apartment. Otherwise every subsequent command
        // (which runs on the worker) would marshal across apartments.
        if (wait)
        {
            try
            {
                staWorker.Invoke(() =>
                {
                    session.WaitAndAttach();
                    TryConnectHooks(hookConnection);
                }, 60000);
                events.Write(Response.Ok(new { attached = true, version = session.ExcelVersion, hooks = hookConnection.IsConnected }));
            }
            catch (Exception ex) { events.Write(Response.Error($"Failed to attach: {ex.Message}")); }
        }
        else if (pid.HasValue)
        {
            try
            {
                staWorker.Invoke(() =>
                {
                    session.Attach(pid);
                    TryConnectHooks(hookConnection);
                }, 15000);
                events.Write(Response.Ok(new { attached = true, version = session.ExcelVersion, hooks = hookConnection.IsConnected }));
            }
            catch (Exception ex) { events.Write(Response.Error($"Failed to attach to PID {pid}: {ex.Message}")); }
        }

        // === Run REPL ===
        new Repl(router, events).Run();

        // === Cleanup ===
        hookConnection.Disconnect();
        session.Dispose();
        staWorker.Dispose();
        return 0;
    }

    private struct HooksVersionInfo
    {
        public bool Available;
        public string? AssemblyVersion;
        // The XRai.Hooks library DLL build timestamp. Same value as the
        // legacy BuildTimestamp; the renamed field makes it clear which
        // build this refers to.
        public string? LibraryBuildTimestamp;
        // The consuming add-in's .xll File.GetLastWriteTime, if resolvable.
        // Null on older XRai.Hooks builds that don't return this field.
        public string? AddinBuildTimestamp;
        // DEPRECATED: alias of LibraryBuildTimestamp. Kept so existing
        // call sites continue to surface a value during the transition.
        public string? BuildTimestamp;
    }

    /// <summary>
    /// Query the connected hooks pipe for its assembly version and build timestamp.
    /// Returns an Available=false struct if hooks aren't connected or the pipe
    /// server doesn't support the pane_status command (pre-Round-7 XRai.Hooks builds).
    /// </summary>
    private static HooksVersionInfo ProbeHooksVersion(HookConnection hooks)
    {
        if (!hooks.IsConnected) return default;
        try
        {
            var response = hooks.SendCommand("pane_status");
            if (response?["ok"]?.GetValue<bool>() != true) return default;
            // Prefer the new field name; fall back to the deprecated one so
            // we keep working against older XRai.Hooks builds in the wild.
            var libraryTs = response["hooks_library_build_timestamp"]?.GetValue<string>()
                            ?? response["hooks_build_timestamp"]?.GetValue<string>();
            return new HooksVersionInfo
            {
                Available = true,
                AssemblyVersion = response["hooks_assembly_version"]?.GetValue<string>(),
                LibraryBuildTimestamp = libraryTs,
                AddinBuildTimestamp = response["addin_build_timestamp"]?.GetValue<string>(),
                BuildTimestamp = libraryTs,
            };
        }
        catch
        {
            return default;
        }
    }

    /// <summary>
    /// Compare the hooks assembly's build timestamp against this CLI's build timestamp.
    /// If the hooks DLL is significantly older than the CLI, the add-in is running
    /// stale XRai.Hooks code — any bug fixes shipped since then will NOT be in effect
    /// inside the add-in. Returns a human-readable warning or null if the versions
    /// are compatible.
    ///
    /// This catches the failure mode where an XRai update ships, the CLI picks up
    /// the new code, but the add-in's bin/Debug still has the old XRai.Hooks.dll
    /// because NuGet cached 1.0.0 and didn't re-pull. Bumping XRai.Hooks.csproj
    /// version per-build (see Round 8) lets consumers with Version="1.0.*"
    /// automatically upgrade — this warning surfaces the problem when they don't.
    /// </summary>
    private static string? CheckHooksStaleness(HooksVersionInfo info)
    {
        if (!info.Available) return null;
        if (string.IsNullOrEmpty(info.BuildTimestamp) || info.BuildTimestamp == "unknown") return null;

        try
        {
            var cliAsm = typeof(DaemonServer).Assembly;
            var cliTimestampAttr = cliAsm.GetCustomAttributes(typeof(System.Reflection.AssemblyMetadataAttribute), false)
                .Cast<System.Reflection.AssemblyMetadataAttribute>()
                .FirstOrDefault(m => m.Key == "XRaiBuildTimestamp");

            // CLI doesn't embed its own timestamp yet — compare against a rough "days old" check instead
            if (!DateTime.TryParse(info.BuildTimestamp, System.Globalization.CultureInfo.InvariantCulture,
                System.Globalization.DateTimeStyles.RoundtripKind, out var hooksBuilt))
                return null;

            var ageHours = (DateTime.UtcNow - hooksBuilt).TotalHours;

            // If the hooks DLL is more than 24 hours older than "now", warn.
            // During active XRai development, hooks should always be within minutes.
            // This threshold is intentionally generous to avoid false positives for
            // stable installs.
            if (ageHours > 24)
            {
                return $"XRai.Hooks DLL loaded in the add-in is {ageHours:F0} hours old ({hooksBuilt:yyyy-MM-dd HH:mm} UTC). " +
                    "Recent XRai fixes may not be in effect. To upgrade: " +
                    "(1) Ensure your add-in's .csproj has <PackageReference Include=\"XRai.Hooks\" Version=\"1.0.*\" />, " +
                    "(2) Run 'dotnet nuget locals http-cache --clear' to clear NuGet cache, " +
                    "(3) Rebuild the add-in project, " +
                    "(4) Kill Excel and reload the .xll.";
            }
        }
        catch { }
        return null;
    }

    private static void TryConnectHooks(HookConnection hooks)
    {
        try
        {
            var processes = System.Diagnostics.Process.GetProcessesByName("EXCEL");
            if (processes.Length > 0)
                hooks.Connect(processes[0].Id, 2000);
        }
        catch { /* Hooks are optional */ }
    }

    /// <summary>
    /// Build a structured diagnostic object explaining why the hooks pipe is not
    /// connected. Called when hooks:false is returned so Claude can see the
    /// actual failure mode instead of flying blind.
    /// </summary>
    internal static object GetHooksDiagnostics(HookConnection hooks)
    {
        var procs = System.Diagnostics.Process.GetProcessesByName("EXCEL");
        var pids = procs.Select(p => p.Id).ToArray();

        // Try each Excel process's expected pipe name and report what we see
        var pipeAttempts = new List<object>();
        foreach (var p in procs)
        {
            var pipeName = $"xrai_{p.Id}";
            var pipePath = $@"\\.\pipe\{pipeName}";
            bool exists = false;
            string? connectError = null;
            try
            {
                // Check if the named pipe exists without consuming it
                var pipes = System.IO.Directory.GetFiles(@"\\.\pipe\");
                exists = pipes.Any(x => x.EndsWith(pipeName, StringComparison.OrdinalIgnoreCase));
            }
            catch { exists = false; }

            // Attempt a lightweight connect to capture the real error
            if (!hooks.IsConnected)
            {
                try
                {
                    hooks.Connect(p.Id, 500);
                }
                catch (Exception ex)
                {
                    connectError = $"{ex.GetType().Name}: {ex.Message}";
                }
            }

            pipeAttempts.Add(new
            {
                pid = p.Id,
                pipe_name = pipeName,
                pipe_path = pipePath,
                pipe_exists = exists,
                connect_error = connectError,
            });
        }

        return new
        {
            excel_process_count = procs.Length,
            excel_pids = pids,
            pipe_attempts = pipeAttempts,
            hint = procs.Length == 0
                ? "No Excel process running. Launch Excel and load your .xll add-in."
                : (pipeAttempts.Cast<dynamic>().All(p => !(bool)p.pipe_exists)
                    ? "Excel is running but no XRai.Hooks pipe was found. The add-in either (1) doesn't reference XRai.Hooks NuGet, (2) doesn't call Pilot.Start() in AutoOpen(), or (3) isn't loaded (.xll not registered)."
                    : "Pipe exists but connect failed. The add-in may be initializing — wait 1-2s and retry {\"cmd\":\"connect\"}."),
        };
    }

    private static int PrintHelp()
    {
        Console.WriteLine();
        Console.WriteLine("  xrai — Excel automation CLI for AI agents");
        Console.WriteLine();
        Console.WriteLine("USAGE:");
        Console.WriteLine("  XRai.Tool.exe [subcommand] [flags]");
        Console.WriteLine("  echo '<json-command>' | XRai.Tool.exe [flags]");
        Console.WriteLine();
        Console.WriteLine("CLI SUBCOMMANDS:");
        Console.WriteLine("  studio             Launch XRai Studio — live dashboard that watches your AI");
        Console.WriteLine("                     coding agent and your target app side-by-side.");
        Console.WriteLine("                     Equivalent to --studio. Auto-opens your browser.");
        Console.WriteLine("  ides               List every editor Studio detects (installed + running)");
        Console.WriteLine("  set-ide <kind>     Persist the user's preferred editor so Studio skips the");
        Console.WriteLine("                     onboarding overlay. Kind = VSCode | VisualStudio | Rider.");
        Console.WriteLine("                     Run this at the START of every greenfield session.");
        Console.WriteLine("  get-ide            Print the currently-persisted editor preference");
        Console.WriteLine("  daemon             Run as the long-lived XRai daemon (no Studio dashboard)");
        Console.WriteLine("  daemon-status      Check whether the daemon is running");
        Console.WriteLine("  daemon-stop        Stop a running daemon");
        Console.WriteLine("  doctor             Run system diagnostics (9 checks)");
        Console.WriteLine("  init <name>        Scaffold a new XRai-enabled Excel-DNA add-in project");
        Console.WriteLine("  kill-excel         Force-kill all Excel processes (zombie cleanup)");
        Console.WriteLine("  dump-commands      Print every registered JSON command (used to regenerate docs)");
        Console.WriteLine("  help               Show this message");
        Console.WriteLine();
        Console.WriteLine("CLI FLAGS:");
        Console.WriteLine("  --studio           Run the daemon WITH the Studio dashboard enabled");
        Console.WriteLine("  --no-browser       Suppress the auto browser launch (for headless / RDP use)");
        Console.WriteLine("  --pid <n>          Attach to a specific Excel PID on startup");
        Console.WriteLine("  --wait             Wait for Excel to appear, then attach, then enter REPL");
        Console.WriteLine("  --repl             Persistent REPL mode (stdin-driven, stays attached)");
        Console.WriteLine("  --timeout <ms>     Default timeout per command (default: 15000)");
        Console.WriteLine();
        Console.WriteLine("GREENFIELD QUICK START (AI agents: run these in order):");
        Console.WriteLine("  1. xrai ides                          # see which editors are detected");
        Console.WriteLine("  2. ASK the user which editor they use (VSCode / VisualStudio / Rider)");
        Console.WriteLine("  3. xrai set-ide <their choice>        # persist so Studio picks it up");
        Console.WriteLine("  4. xrai init MyAddin                  # scaffold the project");
        Console.WriteLine("  5. (optional) xrai --studio           # launch dashboard in separate terminal");
        Console.WriteLine();
        Console.WriteLine("STUDIO QUICK START:");
        Console.WriteLine("  XRai.Tool.exe --studio");
        Console.WriteLine("    → starts the daemon, launches the Studio web dashboard, opens your browser.");
        Console.WriteLine("    → if `xrai set-ide <kind>` was run first, the onboarding overlay is skipped.");
        Console.WriteLine("    → otherwise pick your IDE in the overlay, then watch your AI agent edit");
        Console.WriteLine("      code live as Excel updates alongside.");
        Console.WriteLine("    → zero impact on the agent — Studio passively reads transcript files.");
        Console.WriteLine();
        Console.WriteLine("JSON PROTOCOL:");
        Console.WriteLine("  Every stdin line must be one JSON object: {\"cmd\":\"...\",\"arg\":...}");
        Console.WriteLine("  Responses on stdout: {\"ok\":true,...} or {\"ok\":false,\"error\":\"...\"}");
        Console.WriteLine();
        Console.WriteLine("GOLDEN PATH (always start with this):");
        Console.WriteLine("  {\"cmd\":\"connect\"}");
        Console.WriteLine("    → attaches to Excel, waits for COM readiness, ensures a workbook,");
        Console.WriteLine("      connects hooks pipe, returns full state");
        Console.WriteLine();
        Console.WriteLine("DISCOVERY COMMANDS (use these to learn the API):");
        Console.WriteLine("  {\"cmd\":\"help\"}         → list every registered JSON command");
        Console.WriteLine("  {\"cmd\":\"commands\"}     → flat list of command names");
        Console.WriteLine("  {\"cmd\":\"status\"}       → attachment + workbook + hooks state");
        Console.WriteLine("  {\"cmd\":\"pane.status\"}  → pipe + exposed controls + exposed models");
        Console.WriteLine("  {\"cmd\":\"ribbon\"}       → list ribbon tabs");
        Console.WriteLine("  {\"cmd\":\"ribbon.buttons\",\"tab\":\"<tab>\"} → list buttons on a tab");
        Console.WriteLine();
        Console.WriteLine("For the full command catalog, run:  XRai.Tool.exe dump-commands");
        Console.WriteLine();
        return 0;
    }

    private static int DumpCommands()
    {
        // Spin up the router in the same shape as Main() but WITHOUT attaching to Excel.
        // This lets us introspect every registered command name.
        var events = new EventStream(Console.Out);
        var router = new CommandRouter(events);
        var session = new ExcelSession();
        var hookConnection = new HookConnection();

        router.Register("wait", _ => Response.Ok());
        router.Register("attach", _ => Response.Ok());
        router.Register("detach", _ => Response.Ok());
        router.Register("status", _ => Response.Ok());
        router.Register("ensure.workbook", _ => Response.Ok());
        router.Register("connect", _ => Response.Ok());

        new CellOps(session).Register(router);
        new SheetOps(session).Register(router);
        new CalcOps(session).Register(router);
        new WorkbookOps(session).Register(router);
        new LayoutOps(session).Register(router);
        new DataOps(session).Register(router);
        new FormatOps(session).Register(router);
        new ChartOps(session).Register(router);
        new SparklineOps(session).Register(router);
        new TableOps(session).Register(router);
        new FilterOps(session).Register(router);
        new PivotOps(session).Register(router);
        new PrintOps(session).Register(router);
        new WindowOps(session).Register(router);
        new ShapeOps(session).Register(router);
        new AdvancedOps(session).Register(router);
        new PowerQueryOps(session).Register(router);
        new VbaOps(session).Register(router);
        new SlicerOps(session).Register(router);
        new ConnectionOps(session).Register(router);
        new DesktopOps().Register(router);
        new AppAttachOps().Register(router);
        new PaneClient(hookConnection).Register(router);
        new ModelClient(hookConnection).Register(router);
        new RibbonDriver().Register(router);
        new Win32DialogDriver().Register(router);
        new FileDialogDriver().Register(router);
        new Capture().Register(router);
        new DiffOps(new Capture()).Register(router);
        new OcrOps().Register(router);
        new WaitOps().Register(router);
        new TestReporter().Register(router);
        new AssertOps(router).Register(router);
        new ReloadOrchestrator(session, hookConnection).Register(router);
        router.Register("excel.kill", _ => Response.Ok());
        router.Register("help", _ => Response.Ok());
        router.Register("commands", _ => Response.Ok());

        // Emit as structured markdown for commands.md auto-generation
        var cmds = router.RegisteredCommands.OrderBy(c => c).ToArray();
        Console.WriteLine("# XRai Command Reference (auto-generated)");
        Console.WriteLine();
        Console.WriteLine($"Total: {cmds.Length} commands. Regenerated from `XRai.Tool.exe dump-commands`.");
        Console.WriteLine();
        Console.WriteLine("## All Commands");
        Console.WriteLine();
        foreach (var c in cmds)
        {
            Console.WriteLine($"- `{c}`");
        }
        Console.WriteLine();
        Console.WriteLine("## CLI Subcommands (not JSON)");
        Console.WriteLine();
        Console.WriteLine("- `XRai.Tool.exe doctor` — system diagnostics");
        Console.WriteLine("- `XRai.Tool.exe init <name>` — scaffold new add-in");
        Console.WriteLine("- `XRai.Tool.exe kill-excel` — force-kill all Excel processes");
        Console.WriteLine("- `XRai.Tool.exe dump-commands` — regenerate this file");
        Console.WriteLine("- `XRai.Tool.exe help` — CLI help");
        Console.WriteLine();
        Console.WriteLine("## CLI Flags");
        Console.WriteLine();
        Console.WriteLine("- `--pid <n>` — attach to specific Excel PID");
        Console.WriteLine("- `--wait` — wait for Excel, then attach");
        Console.WriteLine("- `--repl` — persistent REPL mode");
        Console.WriteLine("- `--timeout <ms>` — default per-command timeout (default 15000)");
        return 0;
    }

    private static int DaemonStatus()
    {
        Console.WriteLine();
        Console.WriteLine($"  Daemon pipe: {DaemonServer.PipeName}");
        if (DaemonClient.Ping())
        {
            Console.WriteLine("  Status:      RUNNING");
            Console.WriteLine();
            Console.WriteLine("  Any xrai.exe invocation will automatically forward commands through the daemon.");
            Console.WriteLine("  Stop with: XRai.Tool.exe daemon-stop");
            return 0;
        }
        else
        {
            Console.WriteLine("  Status:      NOT RUNNING");
            Console.WriteLine();
            Console.WriteLine("  Start with: XRai.Tool.exe --daemon");
            return 1;
        }
    }

    private static int DaemonStop()
    {
        Console.WriteLine();
        if (!DaemonClient.Ping())
        {
            Console.WriteLine("  Daemon is not running.");
            return 0;
        }
        Console.WriteLine("  Sending stop signal to daemon...");
        if (DaemonClient.SendStop())
        {
            Console.WriteLine("  Daemon stopped.");
            return 0;
        }
        Console.WriteLine("  Failed to stop daemon.");
        return 1;
    }

    private static int KillExcel()
    {
        Console.WriteLine();
        Console.WriteLine("  xrai kill-excel — force-terminate all Excel processes");
        Console.WriteLine();

        var procs = System.Diagnostics.Process.GetProcessesByName("EXCEL");
        if (procs.Length == 0)
        {
            Console.WriteLine("  No Excel processes found. Nothing to kill.");
            return 0;
        }

        Console.WriteLine($"  Found {procs.Length} Excel process(es):");
        foreach (var p in procs)
            Console.WriteLine($"    PID {p.Id}  |  {p.MainWindowTitle}");

        Console.WriteLine();
        Console.WriteLine("  Killing...");

        int killed = 0;
        foreach (var p in procs)
        {
            try
            {
                p.Kill(entireProcessTree: true);
                p.WaitForExit(5000);
                killed++;
                Console.WriteLine($"    Killed PID {p.Id}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"    Failed to kill PID {p.Id}: {ex.Message}");
            }
        }

        // Wait a moment, then verify
        System.Threading.Thread.Sleep(500);
        var remaining = System.Diagnostics.Process.GetProcessesByName("EXCEL");
        Console.WriteLine();
        if (remaining.Length == 0)
            Console.WriteLine($"  Done. {killed} process(es) killed. No Excel processes remain.");
        else
            Console.WriteLine($"  Warning: {remaining.Length} Excel process(es) still running.");

        return remaining.Length == 0 ? 0 : 1;
    }

    // ── IDE preference CLI ──────────────────────────────────────

    /// <summary>
    /// Persist a user IDE preference to Studio's preferences file so Studio
    /// picks it up without showing the onboarding overlay. The typical flow:
    /// at the start of a greenfield session the agent asks the user which
    /// IDE they use, then runs `xrai set-ide &lt;kind&gt;`. Later when the user
    /// launches `xrai --studio`, the onboarding overlay is skipped entirely.
    ///
    /// Usage:
    ///   xrai set-ide VSCode
    ///   xrai set-ide VisualStudio   (for VS 2022 or VS 2026)
    ///   xrai set-ide Rider
    /// </summary>
    private static int SetIdeCommand(string[] args, int index)
    {
        if (index + 1 >= args.Length)
        {
            Console.Error.WriteLine();
            Console.Error.WriteLine("  xrai set-ide — set the preferred editor for Studio follow-mode");
            Console.Error.WriteLine();
            Console.Error.WriteLine("  Usage:  xrai set-ide <kind>");
            Console.Error.WriteLine("          where <kind> is VSCode | VisualStudio | Rider");
            Console.Error.WriteLine();
            Console.Error.WriteLine("  This writes preferredIde + onboarded=true to Studio's preferences");
            Console.Error.WriteLine("  file at %LOCALAPPDATA%\\XRai\\studio\\preferences.json. Studio picks");
            Console.Error.WriteLine("  it up the next time you run `xrai --studio` and skips the overlay.");
            Console.Error.WriteLine();
            Console.Error.WriteLine("  Tip: run `xrai ides` to see which editors Studio detects on this");
            Console.Error.WriteLine("  machine and which of them is currently running.");
            return 1;
        }

        var kindArg = args[index + 1];
        if (!Enum.TryParse<XRai.Studio.IdeLauncher.IdeKind>(kindArg, ignoreCase: true, out var kind) ||
            kind == XRai.Studio.IdeLauncher.IdeKind.None ||
            kind == XRai.Studio.IdeLauncher.IdeKind.Fallback)
        {
            Console.Error.WriteLine($"  Unknown IDE kind: '{kindArg}'");
            Console.Error.WriteLine($"  Valid values: VSCode, VisualStudio, Rider");
            return 1;
        }

        // Verify the chosen IDE is actually installed so the agent and user
        // can see the problem immediately instead of later when follow-mode
        // silently falls back to the default editor.
        var all = XRai.Studio.IdeLauncher.DetectAll();
        var info = all.FirstOrDefault(i => i.Kind == kind);
        if (info == null || !info.Installed)
        {
            Console.Error.WriteLine();
            Console.Error.WriteLine($"  WARNING: {kind} is not installed on this machine.");
            if (info?.InstallUrl != null)
                Console.Error.WriteLine($"  Install it from: {info.InstallUrl}");
            Console.Error.WriteLine($"  Preference saved anyway — Studio will fall back to Windows file association");
            Console.Error.WriteLine($"  for file opens until you install {kind}.");
            Console.Error.WriteLine();
        }

        var prefs = XRai.Studio.StudioPreferences.Load();
        prefs.PreferredIde = kind.ToString();
        prefs.Onboarded = true;
        if (!prefs.FollowMode) prefs.FollowMode = true;  // opt into follow mode on explicit set
        prefs.Save();

        Console.WriteLine();
        Console.WriteLine($"  Editor preference saved: {kind}");
        Console.WriteLine($"  Follow mode: {(prefs.FollowMode ? "on" : "off")}");
        Console.WriteLine($"  Onboarded:   yes (Studio will skip the overlay)");
        Console.WriteLine();
        Console.WriteLine($"  Stored at: %LOCALAPPDATA%\\XRai\\studio\\preferences.json");
        if (info != null)
        {
            Console.WriteLine($"  Detected:  {info.DisplayName} ({(info.Running ? "running" : info.Installed ? "installed" : "not installed")})");
            if (info.ExecutablePath != null)
                Console.WriteLine($"  At:        {info.ExecutablePath}");
        }
        Console.WriteLine();
        Console.WriteLine("  Launch Studio with: xrai --studio");
        Console.WriteLine();
        return 0;
    }

    /// <summary>
    /// Print the currently-persisted IDE preference. Used by the agent to
    /// verify what Studio will pick up without having to launch it.
    /// </summary>
    private static int GetIdeCommand()
    {
        var prefs = XRai.Studio.StudioPreferences.Load();
        Console.WriteLine();
        Console.WriteLine("  Studio preferences:");
        Console.WriteLine($"    preferredIde: {prefs.PreferredIde ?? "(not set)"}");
        Console.WriteLine($"    followMode:   {prefs.FollowMode}");
        Console.WriteLine($"    onboarded:    {prefs.Onboarded}");
        Console.WriteLine($"    theme:        {prefs.Theme}");
        Console.WriteLine();
        return 0;
    }

    /// <summary>
    /// List every IDE Studio knows about and its installed / running status.
    /// Used by the agent to show the user their options before calling set-ide.
    /// </summary>
    private static int ListIdesCommand()
    {
        var all = XRai.Studio.IdeLauncher.DetectAll();
        Console.WriteLine();
        Console.WriteLine("  Editors detected on this machine:");
        Console.WriteLine();
        foreach (var i in all)
        {
            var status = i.Running ? "RUNNING"
                       : i.Installed ? "installed"
                       : "NOT installed";
            Console.WriteLine($"    {i.Kind,-14} {status,-14} {i.DisplayName}");
            if (i.ExecutablePath != null)
                Console.WriteLine($"                    {i.ExecutablePath}");
            else if (!i.Installed && i.InstallUrl != null)
                Console.WriteLine($"                    install: {i.InstallUrl}");
        }
        Console.WriteLine();
        Console.WriteLine("  Set the user's preference with: xrai set-ide <Kind>");
        Console.WriteLine();
        return 0;
    }
}
