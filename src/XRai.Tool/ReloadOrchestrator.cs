using System.Diagnostics;
using System.Text.Json.Nodes;
using XRai.Core;
using XRai.HooksClient;

namespace XRai.Tool;

public class ReloadOrchestrator
{
    private readonly Com.ExcelSession _session;
    private readonly HookConnection _hooks;

    public ReloadOrchestrator(Com.ExcelSession session, HookConnection hooks)
    {
        _session = session;
        _hooks = hooks;
    }

    public void Register(CommandRouter router)
    {
        router.Register("reload", HandleReload);
        router.Register("rebuild", HandleRebuild);
    }

    /// <summary>
    /// Full kill → build → launch → connect cycle in one command.
    /// Requires "project" (path to .csproj). Optionally accepts "xll" to
    /// override the .xll path — otherwise auto-discovers the *-AddIn64-packed.xll
    /// in the project's publish output.
    /// </summary>
    private string HandleRebuild(JsonObject args)
    {
        var project = args["project"]?.GetValue<string>();
        if (string.IsNullOrEmpty(project))
            return Response.Error("rebuild requires 'project' (path to .csproj)", code: ErrorCodes.MissingArgument);

        if (!File.Exists(project))
            return Response.Error($"Project not found: {project}", code: ErrorCodes.ProjectNotFound);

        var xllOverride = args["xll"]?.GetValue<string>();
        var config = args["config"]?.GetValue<string>() ?? "Debug";
        var steps = new List<string>();
        var sw = Stopwatch.StartNew();

        try
        {
            // Step 1: Quit Excel gracefully (prevents Document Recovery on next launch)
            _hooks.Disconnect();

            // Try graceful quit first: Application.Quit with DisplayAlerts=false
            // so Excel doesn't prompt to save and doesn't leave recovery files.
            bool gracefulQuit = false;
            try
            {
                if (_session.IsAttached)
                {
                    _session.App.DisplayAlerts = false;
                    _session.App.Quit();
                    gracefulQuit = true;
                }
            }
            catch { }

            try { _session.Detach(); } catch { }

            // Wait for graceful quit, then force-kill any survivors
            if (gracefulQuit) Thread.Sleep(2000);

            var procs = Process.GetProcessesByName("EXCEL");
            if (procs.Length > 0)
            {
                foreach (var p in procs)
                {
                    try { p.Kill(entireProcessTree: true); p.WaitForExit(5000); } catch { }
                }
                Thread.Sleep(500);
            }
            steps.Add($"kill-excel: graceful={gracefulQuit}, remaining_killed={procs.Length}");

            // Clean up recovery files so Document Recovery panel doesn't appear
            try
            {
                var recoveryDir = Path.Combine(
                    Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
                    "Microsoft", "Excel");
                foreach (var f in Directory.GetFiles(recoveryDir, "*.xlsb", SearchOption.AllDirectories)
                    .Concat(Directory.GetFiles(recoveryDir, "*.tmp", SearchOption.AllDirectories)))
                {
                    try { File.Delete(f); } catch { }
                }
                // Also clean the XLSTART recovery folder
                var xlstartDir = Path.Combine(recoveryDir, "XLSTART");
                if (Directory.Exists(xlstartDir))
                {
                    foreach (var f in Directory.GetFiles(xlstartDir, "*", SearchOption.AllDirectories))
                    {
                        try { File.Delete(f); } catch { }
                    }
                }
            }
            catch { }

            // Step 2: Ensure XRai-Skill-Local NuGet source exists
            // (idempotent — if already configured, dotnet returns error which we ignore)
            var skillPackagesDir = Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.UserProfile),
                ".claude", "skills", "xrai-excel", "packages");
            if (Directory.Exists(skillPackagesDir))
            {
                RunDotnet($"nuget add source \"{skillPackagesDir}\" --name XRai-Skill-Local", ignoreExit: true);
                steps.Add("nuget-source: ensured XRai-Skill-Local");
            }

            // Step 3: Clear NuGet HTTP cache so wildcard Version="1.0.*" re-resolves
            // to the latest XRai.Hooks package. Without this, NuGet serves a cached
            // older version even when a newer .nupkg is in the local source folder.
            RunDotnet("nuget locals http-cache --clear", ignoreExit: true);
            steps.Add("nuget-cache: cleared");

            // Step 4: Restore (pulls latest XRai.Hooks via the wildcard)
            var restoreResult = RunDotnet($"restore \"{project}\" --force --verbosity quiet");
            if (restoreResult.ExitCode != 0)
            {
                // Non-fatal — build may still succeed if packages are already present
                steps.Add($"dotnet restore: warning (exit {restoreResult.ExitCode})");
            }
            else
            {
                steps.Add("dotnet restore: success");
            }

            // Step 5: Build
            var buildResult = RunDotnet($"build \"{project}\" -c {config} --nologo --verbosity quiet");
            if (buildResult.ExitCode != 0)
            {
                return Response.Error($"Build failed (exit code {buildResult.ExitCode}). " +
                    $"stderr: {buildResult.Stderr.Trim()}. stdout: {buildResult.Stdout.Trim()}",
                    code: ErrorCodes.BuildFailed);
            }
            steps.Add("dotnet build: success");

            // Step 3: Find the .xll
            string xllPath;
            if (!string.IsNullOrEmpty(xllOverride))
            {
                xllPath = Path.GetFullPath(xllOverride);
            }
            else
            {
                // Auto-discover: look for *-AddIn64-packed.xll in publish output
                var projectDir = Path.GetDirectoryName(Path.GetFullPath(project))!;
                var publishDir = Path.Combine(projectDir, "bin", config, "net8.0-windows", "publish");
                var candidates = Directory.Exists(publishDir)
                    ? Directory.GetFiles(publishDir, "*-AddIn64-packed.xll")
                    : Array.Empty<string>();

                if (candidates.Length == 0)
                {
                    // Fall back to non-packed 64-bit
                    var binDir = Path.Combine(projectDir, "bin", config, "net8.0-windows");
                    candidates = Directory.Exists(binDir)
                        ? Directory.GetFiles(binDir, "*-AddIn64.xll")
                        : Array.Empty<string>();
                }

                if (candidates.Length == 0)
                    return Response.Error("Build succeeded but no .xll found. " +
                        "Pass \"xll\":\"path/to/file.xll\" explicitly.");

                xllPath = candidates[0];
            }

            if (!File.Exists(xllPath))
                return Response.Error($"XLL not found after build: {xllPath}");

            steps.Add($"xll: {Path.GetFileName(xllPath)}");

            // Step 4: Launch Excel with the .xll
            Process.Start(new ProcessStartInfo
            {
                FileName = xllPath,
                UseShellExecute = true,
            });
            steps.Add("launch: started");

            // Step 5: Wait for Excel + connect
            int maxWaitMs = 20000;
            int waited = 0;
            bool attached = false;
            while (waited < maxWaitMs)
            {
                Thread.Sleep(1000);
                waited += 1000;
                try
                {
                    _session.Attach();
                    attached = true;
                    break;
                }
                catch { }
            }

            if (!attached)
                return Response.Error("Excel launched but COM attach failed after 20s. " +
                    "Excel may still be loading — try {\"cmd\":\"connect\"} manually.");

            steps.Add($"attach: ok ({waited}ms)");

            // Step 6: Ensure workbook
            var state = _session.ProbeWorkbookState();
            bool created = false;
            if (!state.HasWorkbook)
            {
                try
                {
                    var wb = _session.EnsureWorkbook();
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(wb);
                    created = true;
                    state = _session.ProbeWorkbookState();
                }
                catch { }
            }

            // Step 7: Connect hooks — poll up to `hooks_wait_ms` ms for the
            // pipe + auth token to come online. Default 20s is plenty for a
            // fresh add-in load; increase via {"hooks_wait_ms": 30000} if the
            // add-in has heavy AutoOpen work.
            var hooksWaitMs = args["hooks_wait_ms"]?.GetValue<int?>() ?? 20_000;
            bool hooksOk = false;
            string? lastHooksError = null;
            bool sawToken = false;
            int hooksWaited = 0;
            var hooksSw = Stopwatch.StartNew();

            while (hooksSw.ElapsedMilliseconds < hooksWaitMs)
            {
                try
                {
                    var excelProcs = Process.GetProcessesByName("EXCEL");
                    if (excelProcs.Length > 0)
                    {
                        // Pre-check: does the token file exist yet? If not, Pilot.Start
                        // hasn't run, so the handshake will fail. Skip the actual connect
                        // and wait for the token to appear.
                        var pipeName = $"xrai_{excelProcs[0].Id}";
                        var tokenPath = PipeAuth.GetTokenFilePath(pipeName);
                        if (File.Exists(tokenPath))
                        {
                            sawToken = true;
                            _hooks.Connect(excelProcs[0].Id, 2000);
                            if (_hooks.IsConnected) { hooksOk = true; break; }
                        }
                    }
                }
                catch (Exception ex)
                {
                    lastHooksError = ex.Message;
                }
                Thread.Sleep(500);
            }
            hooksSw.Stop();
            hooksWaited = (int)hooksSw.ElapsedMilliseconds;

            if (hooksOk)
            {
                steps.Add($"hooks: connected ({hooksWaited}ms)");
            }
            else
            {
                var diag = sawToken
                    ? $"hooks: not connected after {hooksWaited}ms — auth token found but handshake failed" +
                      (lastHooksError != null ? $" ({lastHooksError})" : "")
                    : $"hooks: not connected after {hooksWaited}ms — no auth token at %LOCALAPPDATA%\\XRai\\tokens\\xrai_{{pid}}.token. Pilot.Start() may have crashed — check {{\"cmd\":\"log.read\",\"source\":\"startup\"}}";
                steps.Add(diag);
            }

            sw.Stop();
            return Response.Ok(new
            {
                rebuilt = true,
                total_ms = sw.ElapsedMilliseconds,
                steps,
                xll = xllPath,
                hooks = hooksOk,
                hooks_wait_ms = hooksWaited,
                hooks_saw_token = sawToken,
                hooks_last_error = lastHooksError,
                active_workbook = state.Name,
                created_workbook = created,
            });
        }
        catch (Exception ex)
        {
            return Response.Error($"Rebuild failed: {ex.Message}");
        }
    }

    private string HandleReload(JsonObject args)
    {
        var xllPath = args["xll"]?.GetValue<string>();
        var timeoutMs = args["timeout"]?.GetValue<int>() ?? 30000;

        try
        {
            // Step 1: Disconnect hooks
            _hooks.Disconnect();

            // Step 2: Unregister the .xll from Excel
            if (!string.IsNullOrEmpty(xllPath))
            {
                _session.App.RegisterXLL(xllPath); // Toggle off
                Thread.Sleep(500);
            }

            // Step 3: Wait for rebuild (file change or timeout)
            if (!string.IsNullOrEmpty(xllPath))
            {
                var dir = Path.GetDirectoryName(xllPath)!;
                var file = Path.GetFileName(xllPath);
                using var watcher = new FileSystemWatcher(dir, file);
                watcher.EnableRaisingEvents = true;

                var changed = new ManualResetEventSlim(false);
                watcher.Changed += (_, _) => changed.Set();

                if (!changed.Wait(timeoutMs))
                {
                    // Timeout — try to reload anyway
                }
            }

            // Step 4: Re-register the .xll
            if (!string.IsNullOrEmpty(xllPath))
            {
                var result = _session.App.RegisterXLL(xllPath);
                if (!result)
                    return Response.Error($"Failed to re-register XLL: {xllPath}");
            }

            Thread.Sleep(1000); // Let the add-in initialize

            // Step 5: Reconnect hooks
            var processes = Process.GetProcessesByName("EXCEL");
            if (processes.Length > 0)
            {
                try { _hooks.Connect(processes[0].Id, 5000); } catch { }
            }

            return Response.Ok(new
            {
                reloaded = true,
                hooks_connected = _hooks.IsConnected,
                xll = xllPath,
            });
        }
        catch (Exception ex)
        {
            return Response.Error($"Reload failed: {ex.Message}");
        }
    }

    // ── Helpers ──────────────────────────────────────────────────────

    private record DotnetResult(int ExitCode, string Stdout, string Stderr);

    private static DotnetResult RunDotnet(string arguments, bool ignoreExit = false)
    {
        var proc = Process.Start(new ProcessStartInfo
        {
            FileName = "dotnet",
            Arguments = arguments,
            RedirectStandardOutput = true,
            RedirectStandardError = true,
            UseShellExecute = false,
            CreateNoWindow = true,
        });
        proc!.WaitForExit(120000);
        return new DotnetResult(
            proc.ExitCode,
            proc.StandardOutput.ReadToEnd(),
            proc.StandardError.ReadToEnd());
    }
}
