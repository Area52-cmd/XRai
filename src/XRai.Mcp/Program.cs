using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.Hosting;
using Microsoft.Extensions.Logging;
using ModelContextProtocol;
using XRai.Core;
using XRai.Com;
using XRai.HooksClient;
using XRai.Mcp;
using XRai.UI;
using XRai.Vision;

// Handle "setup" subcommand before starting the MCP host
if (args.Length > 0 && args[0].Equals("setup", StringComparison.OrdinalIgnoreCase))
{
    XRai.Mcp.SetupCommand.Run();
    return;
}

var builder = Host.CreateApplicationBuilder(args);

// MCP protocol uses stdout — all logging goes to stderr
builder.Logging.ClearProviders();
builder.Logging.AddConsole(o => o.LogToStandardErrorThreshold = LogLevel.Trace);

// Register XRai infrastructure as singletons
builder.Services.AddSingleton<StaComWorker>();
builder.Services.AddSingleton<ExcelSession>();
builder.Services.AddSingleton<HookConnection>();
builder.Services.AddSingleton<Win32DialogDriver>();

builder.Services.AddSingleton(sp =>
{
    var events = new EventStream(Console.Error);
    var router = new CommandRouter(events);
    var staWorker = sp.GetRequiredService<StaComWorker>();
    var session = sp.GetRequiredService<ExcelSession>();
    var hooks = sp.GetRequiredService<HookConnection>();
    var dialogDriver = sp.GetRequiredService<Win32DialogDriver>();

    router.SetStaInvoker((func, timeout) => staWorker.Invoke(func, timeout));
    router.SetTimeoutDiagnostics(dialogDriver);

    // Auto-recover the STA worker on timeout so MCP sessions don't get
    // permanently stuck. Same pattern as the daemon.
    router.StaAutoRecover = () =>
    {
        try
        {
            try { session.Detach(); } catch { }
            try { hooks.Disconnect(); } catch { }
            staWorker.Reset();
            try { session.Attach(); } catch { }
            return true;
        }
        catch { return false; }
    };

    // === Connection commands (replicated from XRai.Tool Program.cs) ===
    router.Register("wait", _ =>
    {
        session.WaitAndAttach();
        TryConnectHooks(hooks);
        var state = session.ProbeWorkbookState();
        return Response.Ok(new
        {
            attached = true,
            version = session.ExcelVersion,
            hooks = hooks.IsConnected,
            has_workbook = state.HasWorkbook,
            active_workbook = state.Name,
            workbook_count = state.Count
        });
    });

    router.Register("attach", cmdArgs =>
    {
        var attachPid = cmdArgs["pid"]?.GetValue<int>();
        session.Attach(attachPid);
        TryConnectHooks(hooks);
        var state = session.ProbeWorkbookState();
        return Response.Ok(new
        {
            attached = true,
            version = session.ExcelVersion,
            hooks = hooks.IsConnected,
            has_workbook = state.HasWorkbook,
            active_workbook = state.Name,
            workbook_count = state.Count
        });
    });

    router.Register("detach", _ =>
    {
        hooks.Disconnect();
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
        return Response.Ok(new
        {
            attached = true,
            version = session.ExcelVersion,
            hooks = hooks.IsConnected,
            hooks_pipe = hooks.PipeName,
            has_workbook = state.HasWorkbook,
            active_workbook = state.Name,
            workbook_count = state.Count,
            hint = state.HasWorkbook
                ? (string?)null
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

    router.Register("connect", cmdArgs =>
    {
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

        TryConnectHooks(hooks);

        return Response.Ok(new
        {
            attached = true,
            version = session.ExcelVersion,
            hooks = hooks.IsConnected,
            hooks_pipe = hooks.PipeName,
            active_workbook = state.Name,
            workbook_count = state.Count,
            created_workbook = createdWorkbook
        });
    });

    // === Core COM operations ===
    new CellOps(session).Register(router);
    new SheetOps(session).Register(router);
    new CalcOps(session).Register(router);
    new WorkbookOps(session, dialogDriver).Register(router);

    // === Expanded COM operations ===
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

    // === Hooks ===
    new PaneClient(hooks).Register(router);
    new ModelClient(hooks).Register(router);

    // === FlaUI + Vision ===
    new RibbonDriver(dialogDriver).Register(router);
    dialogDriver.Register(router);
    new FileDialogDriver().Register(router);
    var capture = new Capture();
    capture.SetComHwndProvider(() =>
    {
        try { return session.IsAttached ? (nint)session.App.Hwnd : null; }
        catch { return null; }
    });
    capture.Register(router);
    new DiffOps(capture).Register(router);
    new OcrOps().Register(router);

    // === Intelligent Waits ===
    new WaitOps().Register(router);

    // === Test Reporting + Assertions ===
    new TestReporter().Register(router);
    new AssertOps(router).Register(router);

    // === Reload + Meta ===
    new ReloadOrchestrator(session, hooks).Register(router);

    // === kill-excel as JSON command ===
    router.Register("excel.kill", _ =>
    {
        hooks.Disconnect();
        try { session.Detach(); } catch { }

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
        var remaining = System.Diagnostics.Process.GetProcessesByName("EXCEL").Length;
        return Response.Ok(new { killed_pids = killed, killed_count = killed.Count, remaining });
    });

    // === STA recovery ===
    router.Register("sta.reset", _ =>
    {
        bool wasStuck = staWorker.IsStuck;
        try
        {
            try { session.Detach(); } catch { }
            try { hooks.Disconnect(); } catch { }
            staWorker.Reset();

            // Auto-reattach so the user gets a fully-restored environment in a
            // single call — saves two follow-up round-trips.
            bool excelAttached = false;
            bool hooksAttached = false;
            string? excelError = null;
            string? hooksError = null;
            try
            {
                staWorker.Invoke(() => { session.Attach(); return "ok"; }, 5000);
                excelAttached = session.IsAttached;
            }
            catch (Exception ex) { excelError = ex.Message; }

            if (excelAttached)
            {
                try { hooks.TryAutoConnect(); hooksAttached = hooks.IsConnected; }
                catch (Exception ex) { hooksError = ex.Message; }
            }

            return Response.Ok(new
            {
                reset = true,
                was_stuck = wasStuck,
                attached = excelAttached,
                hooks_connected = hooksAttached,
                attach_error = excelError,
                hooks_error = hooksError,
                hint = excelAttached
                    ? (hooksAttached ? "STA recycled and fully reattached." : "STA recycled; Excel reattached, hooks not found.")
                    : "STA recycled but Excel not attached — launch Excel and run connect.",
            });
        }
        catch (Exception ex) { return Response.ErrorFromException(ex, "sta.reset"); }
    });

    router.Register("sta.status", _ => Response.Ok(new
    {
        is_alive = staWorker.IsAlive,
        is_stuck = staWorker.IsStuck,
        filter_registered = staWorker.FilterRegistered,
        consecutive_timeouts = staWorker.ConsecutiveTimeouts,
    }));

    // === Generic hooks connect (non-Excel apps) ===
    router.RegisterNoSta("hooks.connect", args =>
    {
        var pid = args["pid"]?.GetValue<int?>();
        if (!pid.HasValue)
            return Response.Error("hooks.connect requires 'pid'", code: ErrorCodes.MissingArgument);
        try
        {
            hooks.Connect(pid.Value, 3000);
            return hooks.IsConnected
                ? Response.Ok(new { hooks = true, pid = pid.Value, pipe = hooks.PipeName })
                : Response.Error($"Hooks pipe xrai_{pid.Value} not responding");
        }
        catch (Exception ex) { return Response.ErrorFromException(ex, "hooks.connect"); }
    });

    // === Introspection commands ===
    router.Register("help", _ =>
    {
        var commands = router.RegisteredCommands.ToArray();
        return Response.Ok(new { command_count = commands.Length, commands });
    });

    router.Register("commands", _ =>
    {
        return Response.Ok(new { commands = router.RegisteredCommands.ToArray() });
    });

    return router;
});

// Register MCP server with stdio transport and discover tools from this assembly
builder.Services
    .AddMcpServer()
    .WithStdioServerTransport()
    .WithToolsFromAssembly();

await builder.Build().RunAsync();

// === Helpers ===
static void TryConnectHooks(HookConnection hooks)
{
    try
    {
        var processes = System.Diagnostics.Process.GetProcessesByName("EXCEL");
        if (processes.Length > 0)
            hooks.Connect(processes[0].Id, 2000);
    }
    catch { /* Hooks are optional */ }
}
