// Leak-audited: 2026-04-10 — Repeat ExposeModel/Expose calls overwrite the
// prior registration by key, so they do not accumulate. The registries are
// the only long-lived references on this static class; PipeServer is replaced
// (not chained) by Start(), and Stop() now joins/disposes its background
// thread + cancellation source. Static state survives until process exit by
// design — Pilot is a process-singleton inside the loaded .xll.

using System.ComponentModel;
using System.Diagnostics;
using System.IO;
using System.Windows;

namespace XRai.Hooks;

public static class Pilot
{
    private static PipeServer? _server;
    private static readonly ControlRegistry _controls = new();
    private static readonly ModelRegistry _models = new();
    private static DateTime? _lastExposeAt;
    private static DateTime? _lastExposeModelAt;
    private static int _totalExposeCalls;
    private static int _totalExposeModelCalls;

    public static bool IsRunning => _server != null;
    public static DateTime? LastExposeAt => _lastExposeAt;
    public static DateTime? LastExposeModelAt => _lastExposeModelAt;
    public static int TotalExposeCalls => _totalExposeCalls;
    public static int TotalExposeModelCalls => _totalExposeModelCalls;
    public static int ControlCount => _controls.Count;
    public static int ModelCount => _models.All.Count();

    public static void Start()
    {
        if (_server != null) return;

        int pid = Process.GetCurrentProcess().Id;
        string pipeName = $"xrai_{pid}";

        _server = new PipeServer(pipeName, _controls, _models);
        _server.Start();

        // Install error capture
        ErrorCapture.Install(_server);
        LogCapture.Install(_server);

        Debug.WriteLine($"XRai Pilot started on pipe: {pipeName}");
    }

    public static void Stop()
    {
        LogCapture.Uninstall();

        _server?.Stop();
        _server = null;

        Debug.WriteLine("XRai Pilot stopped");
    }

    /// <summary>
    /// Expose a WPF control (typically a UserControl / task pane) for inspection and interaction.
    /// Walks the visual tree and registers all named controls.
    /// </summary>
    public static void Expose(FrameworkElement element)
    {
        // Capture the WPF dispatcher from the element's thread
        _server?.SetDispatcher(element.Dispatcher);

        _controls.RootElement = element;
        ControlDiscovery.Walk(element, _controls);
        _lastExposeAt = DateTime.UtcNow;
        _totalExposeCalls++;
        Debug.WriteLine($"XRai: Exposed {_controls.Count} controls from {element.GetType().Name}");
    }

    /// <summary>
    /// Expose a WinForms control (typically a Form or UserControl) for inspection and interaction.
    /// Walks the control tree and registers all named controls.
    /// </summary>
    public static void Expose(System.Windows.Forms.Control control)
    {
        if (control == null) throw new ArgumentNullException(nameof(control));
        WinFormsDiscovery.Walk(control, _controls);
        _lastExposeAt = DateTime.UtcNow;
        _totalExposeCalls++;
        Debug.WriteLine($"XRai: Exposed {_controls.Count} controls from WinForms {control.GetType().Name}");
    }

    /// <summary>
    /// Expose a ViewModel (any INotifyPropertyChanged) for property read/write.
    /// The model is registered by <paramref name="name"/> (or the model's type
    /// name if null) and ALSO marked as the default. This way
    /// <c>{"cmd":"model"}</c> with no name still resolves the most-recently
    /// exposed model — fixing the prior bug where calling
    /// <c>ExposeModel(vm, "SomeName")</c> registered the model under
    /// "SomeName" but the unkeyed default lookup never found it, causing
    /// <c>{"cmd":"model"}</c> to fail or appear to hang.
    ///
    /// To look up by key explicitly, use <c>{"cmd":"model","name":"SomeName"}</c>.
    /// </summary>
    public static void ExposeModel(INotifyPropertyChanged viewModel, string? name = null)
    {
        if (viewModel == null) throw new ArgumentNullException(nameof(viewModel));

        var key = name ?? viewModel.GetType().Name;
        _models.Register(viewModel, key);
        _models.SetDefault(viewModel);  // always set default so unkeyed model lookup works

        _lastExposeModelAt = DateTime.UtcNow;
        _totalExposeModelCalls++;
        Debug.WriteLine($"XRai: Exposed model {key}");
    }

    /// <summary>
    /// Send a log message through the hooks pipe AND append to the on-disk
    /// pilot log so callers can use {"cmd":"log.read"} to retrieve recent
    /// activity even after the pipe has dropped events.
    ///
    /// File path: %LOCALAPPDATA%\XRai\logs\pilot-{pid}.log
    /// Auto-rotates at ~10 MB by truncating to half on next write.
    /// Logging never throws — file IO failures are swallowed.
    /// </summary>
    public static void Log(string message, string source = "Hooks")
    {
        // Push live event to any connected pipe client (best effort).
        _server?.PushEvent("log", new { message, source, timestamp = DateTime.UtcNow.ToString("o") });

        // Persist to disk so log.read works even when no client is attached.
        WriteToLogFile(message, source);
    }

    private const long LogRotateBytes = 10L * 1024 * 1024; // 10 MB
    private static readonly object _logLock = new();
    private static string? _logPath;

    private static void WriteToLogFile(string message, string source)
    {
        lock (_logLock)
        {
            try
            {
                _logPath ??= GetLogPath();
                var line = $"[{DateTime.UtcNow:o}] [{source}] {message}";
                Debug.WriteLine(line);

                // Auto-rotate if too large: keep tail half.
                try
                {
                    var fi = new FileInfo(_logPath);
                    if (fi.Exists && fi.Length > LogRotateBytes)
                    {
                        var lines = File.ReadAllLines(_logPath);
                        var keep = lines.Length / 2;
                        File.WriteAllLines(_logPath, lines[keep..]);
                    }
                }
                catch { /* rotation must not throw */ }

                File.AppendAllText(_logPath, line + Environment.NewLine);
            }
            catch
            {
                // Logging must never throw.
            }
        }
    }

    /// <summary>
    /// Returns the absolute path of the on-disk pilot log for this process.
    /// Used by the log.read command to locate the file.
    /// </summary>
    public static string GetLogPath()
    {
        var dir = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
            "XRai", "logs");
        Directory.CreateDirectory(dir);
        var pid = Process.GetCurrentProcess().Id;
        return Path.Combine(dir, $"pilot-{pid}.log");
    }
}
