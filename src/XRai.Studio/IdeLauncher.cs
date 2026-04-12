using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Text.Json.Nodes;
using Microsoft.Win32;

namespace XRai.Studio;

/// <summary>
/// Opens a source file in the user's preferred IDE and probes the system
/// to discover which IDEs are installed, running, and available to launch.
///
/// Studio's design philosophy: instead of embedding a code editor and
/// competing with VS Code / VS 2026 / Rider, Studio hands off to whichever
/// IDE the user already has. The code stays where it belongs — in their
/// IDE — and Studio just "aims" the IDE at the file the agent is editing.
///
/// Detection layers:
///   1. Running processes — preferred, because the user is already there
///   2. Installed IDEs — detected via PATH + Windows registry + common
///      install locations
///   3. Windows file association fallback — whatever the user has set
/// </summary>
public static class IdeLauncher
{
    public enum IdeKind
    {
        None,
        VisualStudio,   // devenv.exe — VS 2022, VS 2026, any future
        VSCode,         // code.exe / code.cmd
        Rider,          // rider64.exe
        Fallback,       // Windows file association
    }

    /// <summary>
    /// Describes an IDE that Studio can use — installed, optionally running,
    /// optionally with a launcher on the PATH. Returned by
    /// <see cref="DetectAll"/> for the dashboard's startup "choose your IDE"
    /// overlay.
    /// </summary>
    public sealed class IdeInfo
    {
        public required IdeKind Kind { get; init; }
        public required string DisplayName { get; init; }
        public required bool Installed { get; init; }
        public required bool Running { get; init; }
        public string? ExecutablePath { get; init; }
        public string? Version { get; init; }
        public string? InstallUrl { get; init; }
        public string? InstallTagline { get; init; }

        public JsonObject ToJson() => new()
        {
            ["kind"] = Kind.ToString(),
            ["name"] = DisplayName,
            ["installed"] = Installed,
            ["running"] = Running,
            ["executablePath"] = ExecutablePath,
            ["version"] = Version,
            ["installUrl"] = InstallUrl,
            ["installTagline"] = InstallTagline,
        };
    }

    // ── Detection cache ─────────────────────────────────────────

    // Install paths basically never change during a daemon session — caching
    // them for 30 seconds eliminates the 4+ shell-outs per Open call. The
    // Running flag is volatile (the user might launch / close their IDE
    // mid-session) so it's recomputed on a tighter 3-second TTL.
    private static readonly object _cacheLock = new();
    private static List<IdeInfo>? _cachedDetect;
    private static long _cacheStampMs;
    private const int InstallCacheTtlMs = 30_000;
    private const int RunningCacheTtlMs = 3_000;

    /// <summary>
    /// Reset the detection cache. Called by tests and after `studio` setup
    /// flows that install a new IDE.
    /// </summary>
    public static void InvalidateCache()
    {
        lock (_cacheLock)
        {
            _cachedDetect = null;
            _cacheStampMs = 0;
        }
    }

    // ── Public API ──────────────────────────────────────────────

    /// <summary>
    /// Return a list of all supported IDEs with their install / running status.
    /// Used by the dashboard startup overlay to offer the user a choice.
    /// Cached for 30 seconds (install paths) with a 3-second refresh on the
    /// Running flag, so repeat calls in a hot-path don't shell out.
    /// </summary>
    public static List<IdeInfo> DetectAll()
    {
        var nowMs = Environment.TickCount64;

        lock (_cacheLock)
        {
            if (_cachedDetect != null)
            {
                var ageMs = nowMs - _cacheStampMs;
                if (ageMs < InstallCacheTtlMs)
                {
                    if (ageMs < RunningCacheTtlMs)
                    {
                        // Fresh — return as-is
                        return _cachedDetect;
                    }
                    // Install paths still fresh, but refresh Running flag.
                    // Build a NEW list to avoid "Collection was modified
                    // during enumeration" — IdeInfo is immutable (init props).
                    var runningSet = GetRunningIdes();
                    var refreshed = new List<IdeInfo>(_cachedDetect.Count);
                    foreach (var i in _cachedDetect)
                    {
                        refreshed.Add(new IdeInfo
                        {
                            Kind = i.Kind,
                            DisplayName = i.DisplayName,
                            Installed = i.Installed,
                            Running = runningSet.Contains(i.Kind),
                            ExecutablePath = i.ExecutablePath,
                            Version = i.Version,
                            InstallUrl = i.InstallUrl,
                            InstallTagline = i.InstallTagline,
                        });
                    }
                    _cachedDetect = refreshed;
                    _cacheStampMs = nowMs - InstallCacheTtlMs + RunningCacheTtlMs;
                    return _cachedDetect;
                }
            }
        }

        // Cold path — full detection
        var fresh = DetectAllUncached();
        lock (_cacheLock)
        {
            _cachedDetect = fresh;
            _cacheStampMs = nowMs;
        }
        return fresh;
    }

    private static List<IdeInfo> DetectAllUncached()
    {
        var runningSet = GetRunningIdes();
        var list = new List<IdeInfo>();

        // VS Code — usually the best first recommendation (free, fastest setup)
        var vscPath = FindVSCode();
        list.Add(new IdeInfo
        {
            Kind = IdeKind.VSCode,
            DisplayName = "Visual Studio Code",
            Installed = vscPath != null,
            Running = runningSet.Contains(IdeKind.VSCode),
            ExecutablePath = vscPath,
            InstallUrl = "https://code.visualstudio.com/Download",
            InstallTagline = "Free, fast, cross-platform. Recommended if you don't already have an IDE.",
        });

        // Visual Studio 2026 / 2022 — the Excel-DNA traditional choice
        var vsPath = FindVisualStudio();
        list.Add(new IdeInfo
        {
            Kind = IdeKind.VisualStudio,
            DisplayName = "Visual Studio",
            Installed = vsPath != null,
            Running = runningSet.Contains(IdeKind.VisualStudio),
            ExecutablePath = vsPath,
            InstallUrl = "https://visualstudio.microsoft.com/downloads/",
            InstallTagline = "The traditional Microsoft IDE. Best WPF / XAML designer integration.",
        });

        // JetBrains Rider
        var riderPath = FindRider();
        list.Add(new IdeInfo
        {
            Kind = IdeKind.Rider,
            DisplayName = "JetBrains Rider",
            Installed = riderPath != null,
            Running = runningSet.Contains(IdeKind.Rider),
            ExecutablePath = riderPath,
            InstallUrl = "https://www.jetbrains.com/rider/download/",
            InstallTagline = "Paid. Loved by .NET power users. Strong refactoring + test runner.",
        });

        return list;
    }

    /// <summary>
    /// Open the given file in the best available IDE. Optionally jumps to
    /// a specific line (1-indexed). Preference order:
    ///   1. If <paramref name="preferredKind"/> is supplied and that IDE is
    ///      installed (running or not), use it.
    ///   2. Otherwise prefer a running IDE.
    ///   3. Otherwise prefer any installed IDE.
    ///   4. Otherwise fall back to Windows file association.
    /// Returns a JSON object describing what happened.
    /// </summary>
    public static JsonObject Open(string filePath, int? line = null, int? column = null, IdeKind? preferredKind = null)
    {
        if (string.IsNullOrWhiteSpace(filePath))
        {
            return new JsonObject { ["ok"] = false, ["error"] = "file_path is required" };
        }

        filePath = Path.GetFullPath(filePath);
        if (!File.Exists(filePath))
        {
            return new JsonObject
            {
                ["ok"] = false,
                ["error"] = $"File does not exist: {filePath}",
            };
        }

        var all = DetectAll();

        // 1. User-preferred IDE (if installed)
        if (preferredKind.HasValue)
        {
            var pref = all.FirstOrDefault(i => i.Kind == preferredKind.Value && i.Installed);
            if (pref != null)
            {
                var result = LaunchIde(pref, filePath, line, column);
                if (result != null) return result;
            }
        }

        // 2. Running IDE (any)
        var running = all.FirstOrDefault(i => i.Running);
        if (running != null)
        {
            var result = LaunchIde(running, filePath, line, column);
            if (result != null) return result;
        }

        // 3. Any installed IDE
        var installed = all.FirstOrDefault(i => i.Installed);
        if (installed != null)
        {
            var result = LaunchIde(installed, filePath, line, column);
            if (result != null) return result;
        }

        // 4. Windows file association
        try
        {
            Process.Start(new ProcessStartInfo
            {
                FileName = filePath,
                UseShellExecute = true,
            });
            return new JsonObject
            {
                ["ok"] = true,
                ["ide"] = "fallback",
                ["method"] = "windows-file-association",
                ["filePath"] = filePath,
                ["note"] = "No IDE detected; opened via Windows file association.",
            };
        }
        catch (Exception ex)
        {
            return new JsonObject
            {
                ["ok"] = false,
                ["error"] = $"All launch paths failed: {ex.Message}",
                ["filePath"] = filePath,
            };
        }
    }

    /// <summary>
    /// Launch an IDE without opening any particular file — used by the
    /// startup overlay when the user wants to boot their IDE before the
    /// agent starts working.
    /// </summary>
    public static JsonObject LaunchBlank(IdeKind kind)
    {
        var all = DetectAll();
        var info = all.FirstOrDefault(i => i.Kind == kind);
        if (info == null || !info.Installed || info.ExecutablePath == null)
        {
            return new JsonObject
            {
                ["ok"] = false,
                ["error"] = $"{kind} is not installed.",
                ["kind"] = kind.ToString(),
            };
        }

        try
        {
            Process.Start(new ProcessStartInfo
            {
                FileName = info.ExecutablePath,
                UseShellExecute = true,
            });
            return new JsonObject
            {
                ["ok"] = true,
                ["kind"] = kind.ToString(),
                ["executablePath"] = info.ExecutablePath,
            };
        }
        catch (Exception ex)
        {
            return new JsonObject
            {
                ["ok"] = false,
                ["kind"] = kind.ToString(),
                ["error"] = ex.Message,
            };
        }
    }

    // ── Detection internals ─────────────────────────────────────

    private static HashSet<IdeKind> GetRunningIdes()
    {
        var set = new HashSet<IdeKind>();

        try
        {
            foreach (var p in Process.GetProcessesByName("devenv"))
            {
                set.Add(IdeKind.VisualStudio);
                try { p.Dispose(); } catch { }
            }
        }
        catch { }

        try
        {
            foreach (var p in Process.GetProcessesByName("Code"))
            {
                set.Add(IdeKind.VSCode);
                try { p.Dispose(); } catch { }
            }
        }
        catch { }

        try
        {
            foreach (var p in Process.GetProcessesByName("rider64"))
            {
                set.Add(IdeKind.Rider);
                try { p.Dispose(); } catch { }
            }
        }
        catch { }

        return set;
    }

    private static string? FindVSCode()
    {
        // 1. PATH lookup
        var path = FindOnPath("code.cmd") ?? FindOnPath("code.exe");
        if (path != null) return path;

        // 2. Common install locations
        var candidates = new[]
        {
            Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "Programs", "Microsoft VS Code", "Code.exe"),
            Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ProgramFiles), "Microsoft VS Code", "Code.exe"),
            Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ProgramFilesX86), "Microsoft VS Code", "Code.exe"),
        };
        foreach (var c in candidates)
        {
            try { if (File.Exists(c)) return c; } catch { }
        }

        return null;
    }

    private static string? FindVisualStudio()
    {
        // 1. PATH lookup (rare — devenv isn't usually on PATH by default)
        var path = FindOnPath("devenv.exe");
        if (path != null) return path;

        // 2. vswhere.exe in its standard install location — Microsoft's
        // recommended way to discover Visual Studio installations since VS 2017.
        var vswhere = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.ProgramFilesX86),
            "Microsoft Visual Studio", "Installer", "vswhere.exe");
        if (File.Exists(vswhere))
        {
            try
            {
                var psi = new ProcessStartInfo
                {
                    FileName = vswhere,
                    Arguments = "-latest -property productPath -format value",
                    RedirectStandardOutput = true,
                    UseShellExecute = false,
                    CreateNoWindow = true,
                };
                using var proc = Process.Start(psi);
                if (proc != null)
                {
                    var output = proc.StandardOutput.ReadToEnd().Trim();
                    proc.WaitForExit(1500);
                    if (!string.IsNullOrEmpty(output) && File.Exists(output))
                        return output;
                }
            }
            catch { }
        }

        // 3. Common install locations
        var programFiles = Environment.GetFolderPath(Environment.SpecialFolder.ProgramFiles);
        foreach (var edition in new[] { "Enterprise", "Professional", "Community" })
        {
            foreach (var year in new[] { "2026", "2022" })
            {
                var candidate = Path.Combine(programFiles, "Microsoft Visual Studio", year, edition, "Common7", "IDE", "devenv.exe");
                if (File.Exists(candidate)) return candidate;
            }
        }

        return null;
    }

    private static string? FindRider()
    {
        // 1. PATH lookup
        var path = FindOnPath("rider64.exe");
        if (path != null) return path;

        // 2. JetBrains Toolbox install location (typical)
        var localApps = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
            "Programs");
        if (Directory.Exists(localApps))
        {
            try
            {
                foreach (var dir in Directory.EnumerateDirectories(localApps, "Rider*"))
                {
                    var candidate = Path.Combine(dir, "bin", "rider64.exe");
                    if (File.Exists(candidate)) return candidate;
                }
            }
            catch { }
        }

        // 3. Registry — JetBrains registers an uninstall key
        try
        {
            if (OperatingSystem.IsWindows())
            {
                using var key = Registry.CurrentUser.OpenSubKey(
                    @"SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall");
                if (key != null)
                {
                    foreach (var name in key.GetSubKeyNames())
                    {
                        if (!name.Contains("Rider", StringComparison.OrdinalIgnoreCase)) continue;
                        using var sub = key.OpenSubKey(name);
                        var loc = sub?.GetValue("InstallLocation") as string;
                        if (!string.IsNullOrEmpty(loc))
                        {
                            var candidate = Path.Combine(loc, "bin", "rider64.exe");
                            if (File.Exists(candidate)) return candidate;
                        }
                    }
                }
            }
        }
        catch { }

        return null;
    }

    private static string? FindOnPath(string executableName)
    {
        try
        {
            var psi = new ProcessStartInfo
            {
                FileName = "cmd.exe",
                Arguments = $"/c where {executableName}",
                RedirectStandardOutput = true,
                RedirectStandardError = true,
                UseShellExecute = false,
                CreateNoWindow = true,
            };
            using var proc = Process.Start(psi);
            if (proc == null) return null;
            var output = proc.StandardOutput.ReadToEnd().Trim();
            proc.WaitForExit(1500);
            if (proc.ExitCode != 0 || string.IsNullOrEmpty(output)) return null;
            var firstLine = output.Split('\n')[0].Trim();
            return File.Exists(firstLine) ? firstLine : null;
        }
        catch
        {
            return null;
        }
    }

    // ── Launch internals ────────────────────────────────────────

    private static JsonObject? LaunchIde(IdeInfo info, string filePath, int? line, int? column)
    {
        try
        {
            ProcessStartInfo psi;
            string method;

            switch (info.Kind)
            {
                case IdeKind.VSCode:
                    // code --goto file:line:col — native jump-to-line support
                    var gotoArg = line.HasValue
                        ? $"\"{filePath}:{line.Value}{(column.HasValue ? $":{column.Value}" : "")}\""
                        : $"\"{filePath}\"";
                    psi = new ProcessStartInfo
                    {
                        FileName = "cmd.exe",
                        Arguments = $"/c code --goto {gotoArg}",
                        UseShellExecute = false,
                        CreateNoWindow = true,
                    };
                    method = "code-cli";
                    break;

                case IdeKind.VisualStudio:
                    // devenv /Edit reuses the running instance if any.
                    psi = new ProcessStartInfo
                    {
                        FileName = info.ExecutablePath ?? "devenv",
                        Arguments = $"/Edit \"{filePath}\"",
                        UseShellExecute = false,
                        CreateNoWindow = true,
                    };
                    method = "devenv-edit";
                    break;

                case IdeKind.Rider:
                    var riderLine = line.HasValue ? $"--line {line.Value} " : "";
                    var riderCol = column.HasValue ? $"--column {column.Value} " : "";
                    psi = new ProcessStartInfo
                    {
                        FileName = info.ExecutablePath ?? "rider64",
                        Arguments = $"{riderLine}{riderCol}\"{filePath}\"",
                        UseShellExecute = false,
                        CreateNoWindow = true,
                    };
                    method = "rider-cli";
                    break;

                default:
                    return null;
            }

            Process.Start(psi);

            return new JsonObject
            {
                ["ok"] = true,
                ["ide"] = info.Kind.ToString(),
                ["name"] = info.DisplayName,
                ["method"] = method,
                ["filePath"] = filePath,
                ["line"] = line,
                ["column"] = column,
            };
        }
        catch (Exception ex)
        {
            return new JsonObject
            {
                ["ok"] = false,
                ["ide"] = info.Kind.ToString(),
                ["error"] = ex.Message,
                ["filePath"] = filePath,
            };
        }
    }
}
