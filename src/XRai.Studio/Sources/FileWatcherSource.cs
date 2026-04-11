using System.Collections.Concurrent;
using System.Diagnostics;
using System.Text.Json.Nodes;

namespace XRai.Studio.Sources;

/// <summary>
/// Watches a project directory for file changes and publishes debounced
/// <c>file.changed</c> events to the bus. Designed for the common case where
/// the user (or Claude) edits source files in their IDE while Studio is running
/// in the background — the dashboard file tree flashes whenever something
/// lands on disk.
///
/// Debouncing: editors typically do a rename-storm on save (temp file -> rename
/// -> delete temp). A naive watcher emits 3-5 events per logical save. We
/// coalesce by path with a 200 ms idle window and fire one event carrying the
/// final content.
///
/// Scope: watches the entire subtree under the root, filtered by a default
/// include list (*.cs, *.xaml, *.csproj, *.json, *.md, *.ps1, *.ts, *.js, *.css,
/// *.html, *.yml, *.yaml). Ignores bin/, obj/, dist/, .git/, .vs/, node_modules/.
/// </summary>
public sealed class FileWatcherSource : IDisposable
{
    private readonly EventBus _bus;
    private readonly string _rootPath;
    private readonly FileSystemWatcher _watcher;
    private readonly ConcurrentDictionary<string, Timer> _debouncers = new();
    private readonly int _debounceMs;
    private bool _disposed;

    /// <summary>
    /// Hard cap on the number of in-flight debouncer Timers. A long session
    /// editing many distinct files would otherwise leak Timer objects + GC
    /// handles. When the cap is hit, the oldest debouncer is force-fired
    /// (so its event still lands) before the new one is added.
    /// </summary>
    private const int MaxDebouncers = 256;

    private static readonly string[] IncludeExtensions = new[]
    {
        ".cs", ".xaml", ".csproj", ".sln", ".json", ".md", ".ps1",
        ".ts", ".tsx", ".js", ".jsx", ".mjs", ".css", ".html", ".htm",
        ".yml", ".yaml", ".toml", ".config", ".props", ".targets", ".xml",
    };

    private static readonly string[] IgnoreDirSegments = new[]
    {
        "\\bin\\", "/bin/",
        "\\obj\\", "/obj/",
        "\\dist\\", "/dist/",
        "\\.git\\", "/.git/",
        "\\.vs\\", "/.vs/",
        "\\node_modules\\", "/node_modules/",
        "\\packages\\", "/packages/",
        "\\TestResults\\", "/TestResults/",
    };

    public string RootPath => _rootPath;

    public FileWatcherSource(EventBus bus, string rootPath, int debounceMs = 200)
    {
        _bus = bus;
        _rootPath = Path.GetFullPath(rootPath);
        _debounceMs = debounceMs;

        if (!Directory.Exists(_rootPath))
            throw new DirectoryNotFoundException($"FileWatcherSource root does not exist: {_rootPath}");

        _watcher = new FileSystemWatcher(_rootPath)
        {
            IncludeSubdirectories = true,
            NotifyFilter = NotifyFilters.FileName
                         | NotifyFilters.LastWrite
                         | NotifyFilters.Size
                         | NotifyFilters.CreationTime,
            // Larger buffer reduces overflow on rename storms
            InternalBufferSize = 64 * 1024,
        };

        _watcher.Changed += OnFileSystemEvent;
        _watcher.Created += OnFileSystemEvent;
        _watcher.Renamed += OnFileSystemEvent;
        _watcher.Error += OnWatcherError;
    }

    public void Start()
    {
        _watcher.EnableRaisingEvents = true;

        // Emit an initial event so the dashboard knows the watcher is live
        // and which directory it's watching.
        _bus.Publish(StudioEvent.Now("filewatcher.started", "filewatcher", new JsonObject
        {
            ["root"] = _rootPath,
            ["debounceMs"] = _debounceMs,
        }));
    }

    private void OnWatcherError(object sender, ErrorEventArgs e)
    {
        Debug.WriteLine($"FileWatcher error: {e.GetException().Message}");
        _bus.Publish(StudioEvent.Now("filewatcher.error", "filewatcher", new JsonObject
        {
            ["error"] = e.GetException().Message,
        }));
    }

    private void OnFileSystemEvent(object sender, FileSystemEventArgs e)
    {
        if (_disposed) return;
        if (string.IsNullOrEmpty(e.FullPath)) return;

        // Filter out ignored directories
        foreach (var seg in IgnoreDirSegments)
        {
            if (e.FullPath.Contains(seg, StringComparison.OrdinalIgnoreCase)) return;
        }

        // Filter to tracked extensions — avoid flooding the bus with every
        // temp file an editor creates during save
        var ext = Path.GetExtension(e.FullPath).ToLowerInvariant();
        if (Array.IndexOf(IncludeExtensions, ext) < 0) return;

        // Skip directories (FileSystemWatcher can fire on both)
        try
        {
            if (Directory.Exists(e.FullPath)) return;
        }
        catch { return; }

        // Cap the debouncer dictionary so a long session editing thousands of
        // distinct files doesn't leak Timer objects. When over the cap, force-
        // fire the oldest pending debouncer (so its event still gets emitted)
        // before adding the new one. Cheap O(N) sweep on overflow only.
        if (_debouncers.Count >= MaxDebouncers && !_debouncers.ContainsKey(e.FullPath))
        {
            try
            {
                var victim = _debouncers.Keys.FirstOrDefault();
                if (victim != null && _debouncers.TryRemove(victim, out var oldTimer))
                {
                    try { oldTimer.Dispose(); } catch { }
                    // Force-fire so the user still sees the event
                    FireDebounced(victim);
                }
            }
            catch { }
        }

        // Debounce: reset the timer for this path
        _debouncers.AddOrUpdate(
            e.FullPath,
            _ => new Timer(FireDebounced, e.FullPath, _debounceMs, Timeout.Infinite),
            (_, existing) =>
            {
                try { existing.Change(_debounceMs, Timeout.Infinite); } catch { }
                return existing;
            });
    }

    private void FireDebounced(object? state)
    {
        if (_disposed) return;
        var fullPath = state as string;
        if (string.IsNullOrEmpty(fullPath)) return;

        // Remove + dispose the debouncer for this path
        if (_debouncers.TryRemove(fullPath, out var timer))
        {
            try { timer.Dispose(); } catch { }
        }

        try
        {
            if (!File.Exists(fullPath)) return;
            var fi = new FileInfo(fullPath);
            var relPath = Path.GetRelativePath(_rootPath, fullPath);

            // Keep path separators consistent for the UI (forward slashes)
            var displayPath = relPath.Replace('\\', '/');

            _bus.Publish(StudioEvent.Now("file.changed", "filewatcher", new JsonObject
            {
                ["path"] = displayPath,
                ["absolute"] = fullPath,
                ["sizeBytes"] = fi.Length,
                ["lastWrite"] = fi.LastWriteTimeUtc.ToString("o"),
                ["kind"] = "modified",
                ["extension"] = Path.GetExtension(fullPath).TrimStart('.'),
            }));
        }
        catch (Exception ex)
        {
            Debug.WriteLine($"FileWatcher debounced fire failed: {ex.Message}");
        }
    }

    public void Dispose()
    {
        if (_disposed) return;
        _disposed = true;

        try { _watcher.EnableRaisingEvents = false; } catch { }
        try { _watcher.Changed -= OnFileSystemEvent; } catch { }
        try { _watcher.Created -= OnFileSystemEvent; } catch { }
        try { _watcher.Renamed -= OnFileSystemEvent; } catch { }
        try { _watcher.Error -= OnWatcherError; } catch { }
        try { _watcher.Dispose(); } catch { }

        foreach (var kvp in _debouncers)
        {
            try { kvp.Value.Dispose(); } catch { }
        }
        _debouncers.Clear();
    }
}
