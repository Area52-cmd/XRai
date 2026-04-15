using System.Diagnostics;
using System.Text.Json.Nodes;
using XRai.Core;
using XRai.HooksClient;

namespace XRai.Tool;

public class ReloadOrchestrator
{
    private readonly Com.ExcelSession _session;
    private readonly HookConnection _hooks;

    /// <summary>
    /// Optional factory for a step reporter — called at the start of each
    /// rebuild. When Studio is enabled, DaemonServer sets this to a factory
    /// that returns a TeeStepReporter publishing to the event bus in addition
    /// to the in-memory list. Default (null) means ReloadOrchestrator uses a
    /// plain ListStepReporter and the behavior matches the pre-Studio era.
    /// </summary>
    public Func<IStepReporter>? StepReporterFactory { get; set; }

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

        // Use the injected reporter factory when Studio is active; otherwise
        // fall back to the plain list reporter that matches the pre-Studio
        // behavior exactly. This keeps the rebuild response shape unchanged.
        var reporter = (StepReporterFactory ?? (() => (IStepReporter)new ListStepReporter()))();
        var stepSw = Stopwatch.StartNew();

        // Local helper: report a step with elapsed time since the last Step()
        // call, then restart the per-step stopwatch. Mirrors the old free-form
        // steps.Add() API while giving Studio per-step timing for the dashboard.
        void Step(string name, string status, string? detail = null)
        {
            var elapsed = stepSw.ElapsedMilliseconds;
            reporter.Report(name, status, elapsed, detail);
            stepSw.Restart();
        }

        var sw = Stopwatch.StartNew();

        try
        {
            reporter.Starting("kill-excel");
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
            Step("kill-excel", "ok", $"graceful={gracefulQuit}, remaining_killed={procs.Length}");

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
            reporter.Starting("nuget-source");
            var skillPackagesDir = Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.UserProfile),
                ".claude", "skills", "xrai-excel", "packages");
            if (Directory.Exists(skillPackagesDir))
            {
                RunDotnet($"nuget add source \"{skillPackagesDir}\" --name XRai-Skill-Local", ignoreExit: true);
                Step("nuget-source", "ok", "ensured XRai-Skill-Local");
            }
            else
            {
                Step("nuget-source", "skip", "skill packages dir not present");
            }

            // Step 3: Clear NuGet HTTP cache so wildcard Version="1.0.0-*" re-resolves
            // to the latest XRai.Hooks package. Without this, NuGet serves a cached
            // older version even when a newer .nupkg is in the local source folder.
            reporter.Starting("nuget-cache-clear");
            RunDotnet("nuget locals http-cache --clear", ignoreExit: true);
            Step("nuget-cache-clear", "ok");

            // Step 4: Restore (pulls latest XRai.Hooks via the wildcard)
            reporter.Starting("dotnet-restore");
            var restoreResult = RunDotnet($"restore \"{project}\" --force --verbosity quiet");
            if (restoreResult.ExitCode != 0)
            {
                // Non-fatal — build may still succeed if packages are already present
                Step("dotnet-restore", "warning", $"exit {restoreResult.ExitCode}");
            }
            else
            {
                Step("dotnet-restore", "ok");
            }

            // Step 5: Build. --verbosity normal surfaces MSBuild diagnostics in
            // the single-line parseable format so ExtractCompilerErrors can
            // return a structured errors[] array alongside the raw text.
            reporter.Starting("dotnet-build");
            var buildResult = RunDotnet($"build \"{project}\" -c {config} --nologo --verbosity normal -clp:NoSummary");
            if (buildResult.ExitCode != 0)
            {
                Step("dotnet-build", "error", $"exit {buildResult.ExitCode}");
                var combined = buildResult.Stdout + "\n" + buildResult.Stderr;
                var errors = ExtractCompilerErrors(combined);
                var lines = combined.Split('\n');
                var tail = string.Join("\n", lines.Skip(Math.Max(0, lines.Length - 40)));
                var msg = errors.Length > 0
                    ? $"Build failed: {errors.Length} compiler error(s). First: {((dynamic)errors[0]).message}"
                    : $"Build failed (exit code {buildResult.ExitCode}).";
                return Response.ErrorWithData(msg,
                    data: new { exit_code = buildResult.ExitCode, errors, raw_tail = tail },
                    code: ErrorCodes.BuildFailed);
            }
            Step("dotnet-build", "ok");

            // Step 3: Find the .xll
            string xllPath;
            if (!string.IsNullOrEmpty(xllOverride))
            {
                xllPath = Path.GetFullPath(xllOverride);
            }
            else
            {
                // Auto-discover. Selection rules (most-reliable first):
                //   1. Exact basename match against {csproj-name}-AddIn64-packed.xll
                //      — protects against stale .xlls from a renamed project.
                //   2. Otherwise freshest by LastWriteTime (just-built wins).
                //   3. Warn if two candidates landed within 10s of each other.
                var projectDir = Path.GetDirectoryName(Path.GetFullPath(project))!;
                var csprojName = Path.GetFileNameWithoutExtension(project);
                var publishDir = Path.Combine(projectDir, "bin", config, "net8.0-windows", "publish");
                var candidates = Directory.Exists(publishDir)
                    ? Directory.GetFiles(publishDir, "*-AddIn64-packed.xll")
                    : Array.Empty<string>();

                if (candidates.Length == 0)
                {
                    var binDir = Path.Combine(projectDir, "bin", config, "net8.0-windows");
                    candidates = Directory.Exists(binDir)
                        ? Directory.GetFiles(binDir, "*-AddIn64.xll")
                        : Array.Empty<string>();
                }

                if (candidates.Length == 0)
                    return Response.Error("Build succeeded but no .xll found. " +
                        "Pass \"xll\":\"path/to/file.xll\" explicitly.");

                var expected = $"{csprojName}-AddIn64-packed.xll";
                var nameMatch = candidates.FirstOrDefault(p =>
                    string.Equals(Path.GetFileName(p), expected, StringComparison.OrdinalIgnoreCase));

                if (nameMatch != null)
                {
                    xllPath = nameMatch;
                }
                else
                {
                    var sorted = candidates
                        .Select(p => new { Path = p, Mtime = File.GetLastWriteTimeUtc(p) })
                        .OrderByDescending(x => x.Mtime)
                        .ToArray();
                    xllPath = sorted[0].Path;

                    if (sorted.Length > 1 &&
                        (sorted[0].Mtime - sorted[1].Mtime).TotalSeconds < 10)
                    {
                        Step("xll-resolve", "warning",
                            $"multiple .xll candidates within 10s ({sorted.Length}): picked '{Path.GetFileName(xllPath)}'. " +
                            $"Pass \"xll\" explicitly to disambiguate.");
                    }
                }
            }

            if (!File.Exists(xllPath))
                return Response.Error($"XLL not found after build: {xllPath}");

            Step("xll-resolve", "ok", Path.GetFileName(xllPath));

            // Step 4: Launch Excel with the .xll
            reporter.Starting("launch-excel");
            Process.Start(new ProcessStartInfo
            {
                FileName = xllPath,
                UseShellExecute = true,
            });
            Step("launch-excel", "ok");

            // Step 5: Wait for Excel + connect
            reporter.Starting("attach-com");
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
            {
                Step("attach-com", "error", $"COM attach failed after {waited}ms");
                return Response.Error("Excel launched but COM attach failed after 20s. " +
                    "Excel may still be loading — try {\"cmd\":\"connect\"} manually.");
            }

            Step("attach-com", "ok", $"{waited}ms");

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
                Step("hooks-connect", "ok", $"{hooksWaited}ms");
            }
            else
            {
                var diag = sawToken
                    ? $"auth token found but handshake failed" +
                      (lastHooksError != null ? $" ({lastHooksError})" : "")
                    : $"no auth token — Pilot.Start() may have crashed. Check log.read source=startup";
                Step("hooks-connect", "error", diag);
            }

            sw.Stop();
            return Response.Ok(new
            {
                rebuilt = true,
                total_ms = sw.ElapsedMilliseconds,
                steps = reporter.Lines,
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

    /// <summary>
    /// Called by the daemon after a rebuild returns to ensure the STA worker
    /// isn't left in a stuck state from a timed-out attach step. Without this,
    /// a cold-build timeout poisons the STA and every subsequent command fails
    /// until the user manually runs daemon-stop or sta.reset — pure friction.
    /// </summary>
    public void CleanupAfterRebuild(Action? staResetAction)
    {
        if (staResetAction == null) return;
        // Only fire if the STA actually needs it — the calling site checks
        // _staWorker.IsStuck before invoking this.
        try { staResetAction(); }
        catch { /* best effort — if even the reset fails, the daemon log
                   will show it and the user can restart */ }
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

    /// <summary>
    /// Runs a dotnet command, draining stdout+stderr ASYNCHRONOUSLY so the child
    /// never stalls on a full pipe buffer (default 4KB on Windows). The previous
    /// implementation called WaitForExit BEFORE reading the streams, which
    /// deadlocked whenever a build produced >4KB of compiler errors — the child
    /// blocked writing to stderr and we hit the 120s hard timeout returning
    /// truncated output. BeginOutputReadLine + BeginErrorReadLine drain
    /// continuously into StringBuilders while WaitForExit blocks on exit only.
    /// </summary>
    private static DotnetResult RunDotnet(string arguments, bool ignoreExit = false)
    {
        var proc = new Process
        {
            StartInfo = new ProcessStartInfo
            {
                FileName = "dotnet",
                Arguments = arguments,
                RedirectStandardOutput = true,
                RedirectStandardError = true,
                UseShellExecute = false,
                CreateNoWindow = true,
            }
        };

        var stdout = new System.Text.StringBuilder();
        var stderr = new System.Text.StringBuilder();
        proc.OutputDataReceived += (_, e) => { if (e.Data != null) lock (stdout) stdout.AppendLine(e.Data); };
        proc.ErrorDataReceived  += (_, e) => { if (e.Data != null) lock (stderr) stderr.AppendLine(e.Data); };

        proc.Start();
        proc.BeginOutputReadLine();
        proc.BeginErrorReadLine();

        if (!proc.WaitForExit(600_000))
        {
            try { proc.Kill(entireProcessTree: true); } catch { }
            return new DotnetResult(-1, stdout.ToString(), stderr.ToString() + "\n[xrai] dotnet exceeded 10 min, killed.");
        }
        proc.WaitForExit(); // flush trailing async output

        return new DotnetResult(proc.ExitCode, stdout.ToString(), stderr.ToString());
    }

    /// <summary>
    /// Parse dotnet build output (at -v normal) into structured compiler errors.
    /// Format matches the MSBuild single-line diagnostic:
    ///   Foo.cs(42,13): error CS0103: The name 'bar' does not exist...
    /// </summary>
    public static object[] ExtractCompilerErrors(string output)
    {
        if (string.IsNullOrEmpty(output)) return Array.Empty<object>();

        // Line-by-line parse. We want only lines of the shape
        //   {optional "N>"}path\to\file.cs(line,col): error CS1234: message [project]
        // MSBuild sometimes prefixes project-ref markers like "1>" and suffixes
        // the project path in square brackets. We strip both.
        var rx = new System.Text.RegularExpressions.Regex(
            @"^\s*(?:\d+(?::\d+)?>)?\s*(?<file>[A-Za-z]:[^()\r\n]+?|[^()\r\n]+?)\((?<line>\d+),(?<col>\d+)\):\s+error\s+(?<code>[A-Z]+\d+):\s+(?<msg>.+?)(?:\s+\[[^\]]+\])?\s*$",
            System.Text.RegularExpressions.RegexOptions.Compiled);
        // Strip any stray MSBuild prefix that leaked into the file capture
        // (e.g. if whitespace layout defeated the non-capturing group).
        var prefixRx = new System.Text.RegularExpressions.Regex(@"^\s*\d+(?::\d+)?>\s*", System.Text.RegularExpressions.RegexOptions.Compiled);

        var seen = new HashSet<string>();
        var list = new List<object>();
        foreach (var rawLine in output.Split('\n'))
        {
            var line = rawLine.TrimEnd('\r');
            var m = rx.Match(line);
            if (!m.Success) continue;
            var key = $"{m.Groups["file"].Value}|{m.Groups["line"].Value}|{m.Groups["code"].Value}";
            if (!seen.Add(key)) continue;
            var file = prefixRx.Replace(m.Groups["file"].Value.Trim(), "");
            list.Add(new
            {
                file,
                line = int.Parse(m.Groups["line"].Value),
                col = int.Parse(m.Groups["col"].Value),
                code = m.Groups["code"].Value,
                message = m.Groups["msg"].Value.Trim(),
            });
        }
        return list.ToArray();
    }
}
