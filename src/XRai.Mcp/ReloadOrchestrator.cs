using System.Diagnostics;
using System.IO;
using System.Text.Json.Nodes;
using XRai.Core;
using XRai.HooksClient;

namespace XRai.Mcp;

/// <summary>
/// Copied from XRai.Tool.ReloadOrchestrator — the Tool project is an Exe and
/// cannot be referenced as a project dependency. This duplicate registers the
/// reload/rebuild commands on the CommandRouter for the MCP server.
/// </summary>
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
            _hooks.Disconnect();
            try { _session.Detach(); } catch { }

            var procs = Process.GetProcessesByName("EXCEL");
            foreach (var p in procs)
            {
                try { p.Kill(entireProcessTree: true); p.WaitForExit(5000); } catch { }
            }
            Thread.Sleep(500);
            steps.Add($"kill-excel: {procs.Length} process(es)");

            var skillPackagesDir = Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.UserProfile),
                ".claude", "skills", "xrai-excel", "packages");
            if (Directory.Exists(skillPackagesDir))
            {
                RunDotnet($"nuget add source \"{skillPackagesDir}\" --name XRai-Skill-Local", ignoreExit: true);
                steps.Add("nuget-source: ensured XRai-Skill-Local");
            }

            RunDotnet("nuget locals http-cache --clear", ignoreExit: true);
            steps.Add("nuget-cache: cleared");

            var restoreResult = RunDotnet($"restore \"{project}\" --force --verbosity quiet");
            if (restoreResult.ExitCode != 0)
                steps.Add($"dotnet restore: warning (exit {restoreResult.ExitCode})");
            else
                steps.Add("dotnet restore: success");

            var buildResult = RunDotnet($"build \"{project}\" -c {config} --nologo --verbosity quiet");
            if (buildResult.ExitCode != 0)
            {
                return Response.Error($"Build failed (exit code {buildResult.ExitCode}). " +
                    $"stderr: {buildResult.Stderr.Trim()}. stdout: {buildResult.Stdout.Trim()}",
                    code: ErrorCodes.BuildFailed);
            }
            steps.Add("dotnet build: success");

            string xllPath;
            if (!string.IsNullOrEmpty(xllOverride))
            {
                xllPath = Path.GetFullPath(xllOverride);
            }
            else
            {
                var projectDir = Path.GetDirectoryName(Path.GetFullPath(project))!;
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

                xllPath = candidates[0];
            }

            if (!File.Exists(xllPath))
                return Response.Error($"XLL not found after build: {xllPath}");

            steps.Add($"xll: {Path.GetFileName(xllPath)}");

            Process.Start(new ProcessStartInfo
            {
                FileName = xllPath,
                UseShellExecute = true,
            });
            steps.Add("launch: started");

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

            bool hooksOk = false;
            for (int attempt = 0; attempt < 10; attempt++)
            {
                try
                {
                    var excelProcs = Process.GetProcessesByName("EXCEL");
                    if (excelProcs.Length > 0)
                    {
                        _hooks.Connect(excelProcs[0].Id, 2000);
                        if (_hooks.IsConnected) { hooksOk = true; break; }
                    }
                }
                catch { }
                Thread.Sleep(1000);
            }

            steps.Add(hooksOk ? "hooks: connected" : "hooks: not connected (add-in may still be loading)");

            sw.Stop();
            return Response.Ok(new
            {
                rebuilt = true,
                total_ms = sw.ElapsedMilliseconds,
                steps,
                xll = xllPath,
                hooks = hooksOk,
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
            _hooks.Disconnect();

            if (!string.IsNullOrEmpty(xllPath))
            {
                _session.App.RegisterXLL(xllPath);
                Thread.Sleep(500);
            }

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

            if (!string.IsNullOrEmpty(xllPath))
            {
                var result = _session.App.RegisterXLL(xllPath);
                if (!result)
                    return Response.Error($"Failed to re-register XLL: {xllPath}");
            }

            Thread.Sleep(1000);

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
