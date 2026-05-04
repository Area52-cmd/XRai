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
    // Every WPF dispatcher we've been told about via Expose. Shutdown()
    // walks this list to InvokeShutdown each one and join its thread, so
    // consumers don't have to do that orchestration in AutoClose.
    private static readonly List<WeakReference<System.Windows.Threading.Dispatcher>> _trackedDispatchers = new();
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

        // Scrub Excel's Resiliency state from any prior crashed session so
        // the user never sees the "Excel didn't shut down properly — Safe
        // Mode?" prompt and the addin is never auto-disabled. This is the
        // user-visible mitigation for the documented .NET 8 + Excel-DNA +
        // WPF teardown FailFast (0x80131506). The runtime crash itself is
        // logged to Event Viewer (cosmetic) but Excel forgets it ever
        // happened from the user's perspective.
        try { ScrubOfficeResiliency(); } catch { }

        int pid = Process.GetCurrentProcess().Id;
        string pipeName = $"xrai_{pid}";

        _server = new PipeServer(pipeName, _controls, _models);
        _server.Start();

        // Install error capture
        ErrorCapture.Install(_server);
        LogCapture.Install(_server);

        // Safety net for both Excel-DNA addin unload (collectible
        // AssemblyLoadContext.Unloading) and process exit. Either fires
        // Stop() so consumers who forget to call it manually still get a
        // clean teardown. Subscribe ONCE — Pilot is a process-singleton.
        if (!_processExitHooked)
        {
            try
            {
                AppDomain.CurrentDomain.ProcessExit += OnProcessExit;
                var alc = System.Runtime.Loader.AssemblyLoadContext.GetLoadContext(typeof(Pilot).Assembly);
                if (alc != null) alc.Unloading += OnLoadContextUnloading;
                _processExitHooked = true;
            }
            catch { }
        }

        Debug.WriteLine($"XRai Pilot started on pipe: {pipeName}");
    }

    private static bool _processExitHooked;
    private static int _bypassStrictShutdown = 1; // 1 = on by default

    /// <summary>
    /// Disable the post-AutoClose strict-shutdown bypass. ONLY relevant if
    /// you're hosting XRai.Hooks in a context where CoreCLR's strict
    /// AssemblyLoadContext unload actually works (i.e. NOT Excel-DNA + WPF
    /// on .NET 8). The default (enabled) makes Excel exit cleanly via
    /// TerminateProcess after AutoClose, sidestepping the documented
    /// Excel-DNA/.NET 8/WPF FailFast 0x80131506 issue.
    /// </summary>
    public static void DisableStrictShutdownBypass() { _bypassStrictShutdown = 0; }

    // SAFETY-NET hooks. These are best-effort cleanup paths if the consumer
    // forgets to call Pilot.Shutdown() in AutoClose. They run Stop() only —
    // never the TerminateProcess bypass — because we can't reliably tell
    // whether ProcessExit / ALC.Unloading fired because the HOST is exiting
    // or because the runtime is doing a mid-session unload. Killing Excel
    // mid-session would be far worse than letting the documented .NET 8 +
    // WPF teardown crash happen on host exit. The TerminateProcess path
    // ONLY fires from an explicit Pilot.Shutdown() call (from AutoClose).
    private static void OnProcessExit(object? sender, EventArgs e) { try { Stop(); } catch { } }
    private static void OnLoadContextUnloading(System.Runtime.Loader.AssemblyLoadContext alc) { try { Stop(); } catch { } }

    public static void Stop()
    {
        // Order matters for Excel-DNA / .NET 8 hosted teardown:
        //   1) Drop AppDomain.UnhandledException subscription so the rooted
        //      delegate stops pinning this assembly across LoadContext unload.
        //   2) Remove the Trace listener for the same reason.
        //   3) Clear static events that would otherwise hold subscriber
        //      delegates from this OR consumer assemblies. Setting to null
        //      detaches every handler in one shot.
        //   4) Dispose every registered control/model adapter — releases their
        //      DependencyPropertyDescriptor subscriptions and Unloaded /
        //      Dispatcher.ShutdownStarted handlers.
        //   5) Stop the pipe server (force-closes the live pipe, joins the
        //      worker thread). Must happen LAST so any final shutdown event
        //      writes still land on a live writer.
        // Skipping any of these used to manifest as CoreCLR FailFast 0x80131506
        // (COR_E_EXECUTIONENGINE) on Excel exit when XRai.Hooks is hosted
        // inside an Excel-DNA addin's AssemblyLoadContext.
        try { ErrorCapture.Uninstall(); } catch { }
        try { LogCapture.Uninstall(); } catch { }

        // Detach our own ProcessExit / ALC.Unloading hooks so the rooted
        // delegates do not pin this assembly's load context. Same root-cause
        // class as the ErrorCapture leak.
        if (_processExitHooked)
        {
            try { AppDomain.CurrentDomain.ProcessExit -= OnProcessExit; } catch { }
            try
            {
                var alc = System.Runtime.Loader.AssemblyLoadContext.GetLoadContext(typeof(Pilot).Assembly);
                if (alc != null) alc.Unloading -= OnLoadContextUnloading;
            }
            catch { }
            _processExitHooked = false;
        }

        // Clear static events. These are public so consumer code (e.g. Studio)
        // may have attached. Null assignment removes ALL handlers atomically.
        try { PipeServer.ClearOnEventEmitted(); } catch { }
        try { ControlAdapter.ClearOnControlChanged(); } catch { }
        try { ModelAdapter.ClearOnModelChanged(); } catch { }

        // Dispose all adapters before stopping the pipe so DPD callbacks etc.
        // never fire into half-disposed state. Clear() disposes each adapter.
        try { _controls.Clear(); } catch { }

        _server?.Stop();
        _server = null;

        Debug.WriteLine("XRai Pilot stopped");
    }

    /// <summary>
    /// One-call clean teardown for Excel-DNA / WPF / WinForms host addins.
    /// THIS is what consumer AutoClose should call. Does, in order:
    ///   1. Pilot.Stop() — drops every static event subscription, the
    ///      AppDomain.UnhandledException hook, the pipe-server thread, the
    ///      ProcessExit / ALC.Unloading hooks, every DPD subscription on
    ///      exposed controls, and every model PropertyChanged subscription.
    ///   2. InvokeShutdown on every WPF dispatcher we've seen via Expose,
    ///      and Join the dispatcher thread (up to <paramref name="threadJoinMs"/>)
    ///      so it actually exits before AutoClose returns. Without this,
    ///      Excel-DNA initiates AssemblyLoadContext unload while a managed
    ///      thread is still in WPF native interop and CoreCLR FailFasts
    ///      with 0x80131506 (COR_E_EXECUTIONENGINE) on Excel exit.
    ///   3. Force GC + finalizer drain so any addin-side finalizers run
    ///      BEFORE the load-context unload sweep.
    ///
    /// Recommended consumer usage:
    /// <code>
    ///   public void AutoClose() => Pilot.Shutdown();
    /// </code>
    /// </summary>
    public static void Shutdown(int threadJoinMs = 5000)
    {
        try { Stop(); } catch { }

        // Pre-emptively scrub Resiliency: if the strict-shutdown sweep does
        // crash after this point, the user still won't see the recovery
        // prompt on the next launch. Belt-and-braces with the Pilot.Start
        // scrub.
        try { ScrubOfficeResiliency(); } catch { }

        // Snapshot then clear so the iteration is decoupled from concurrent
        // shutdown signals (e.g. ProcessExit firing while we're already here).
        WeakReference<System.Windows.Threading.Dispatcher>[] refs;
        lock (_trackedDispatchers)
        {
            refs = _trackedDispatchers.ToArray();
            _trackedDispatchers.Clear();
        }

        foreach (var wr in refs)
        {
            if (!wr.TryGetTarget(out var disp)) continue;
            try
            {
                if (disp.Thread.IsAlive)
                {
                    disp.InvokeShutdown();
                    disp.Thread.Join(threadJoinMs);
                }
            }
            catch { /* dispatcher already shut down or thread already ended */ }
        }

        // Drain finalizers before LoadContext unload begins.
        try
        {
            GC.Collect();
            GC.WaitForPendingFinalizers();
            GC.Collect();
        }
        catch { }

        Debug.WriteLine("XRai Pilot shutdown complete");

        // Strict-shutdown bypass. .NET 8 CoreCLR's AssemblyLoadContext-unload
        // sweep is fundamentally incompatible with WPF static state hosted
        // inside an Excel-DNA addin's collectible load context — CoreCLR
        // FailFasts with 0x80131506 (COR_E_EXECUTIONENGINE) when it tries
        // to drain WPF's process-wide statics that root types in the
        // unloading context. We can't fix that from inside the addin (the
        // crash happens after AutoClose returns, regardless of how thorough
        // our cleanup is — verified empirically: identical crash with an
        // empty AutoClose). The pragmatic fix is to terminate the host
        // process cleanly BEFORE CoreCLR begins its strict sweep. Excel is
        // exiting anyway, so this just skips the broken sweep.
        //
        // Only fires when:
        //   1. The bypass hasn't been explicitly disabled, AND
        //   2. The current process appears to be exiting (Excel is killing
        //      the host — not a mid-session unload).
        if (_bypassStrictShutdown != 0 && IsHostExiting())
        {
            // Flush stdio before TerminateProcess (a no-op for Excel but cheap).
            try { Console.Out.Flush(); Console.Error.Flush(); } catch { }

            // Direct TerminateProcess — the only path that does NOT trigger
            // CoreCLR's strict ALC unload sweep. Environment.Exit(0) does
            // trigger it (verified) and re-produces 0x80131506. This call
            // returns 0 with no possibility of throwing or hanging.
            try
            {
                NativeKill.TerminateProcess(NativeKill.GetCurrentProcess(), 0u);
            }
            catch
            {
                // P/Invoke literally cannot fail in any realistic way, but
                // belt-and-braces fall through to managed Kill if it does.
                try { System.Diagnostics.Process.GetCurrentProcess().Kill(); } catch { }
            }
        }
    }

    /// <summary>
    /// Win32 TerminateProcess — the only clean way to exit a host that has
    /// loaded WPF into a collectible AssemblyLoadContext on .NET 8. CoreCLR's
    /// shutdown sweep cannot drain WPF's process-wide statics that root
    /// types in the unloading context, so it FailFasts with 0x80131506.
    /// We bypass the sweep entirely with TerminateProcess from the host's
    /// own AutoClose path. Excel is already exiting; this just makes the
    /// exit clean instead of crash-logged.
    /// </summary>
    private static class NativeKill
    {
        [System.Runtime.InteropServices.DllImport("kernel32.dll", SetLastError = true)]
        public static extern IntPtr GetCurrentProcess();

        [System.Runtime.InteropServices.DllImport("kernel32.dll", SetLastError = true)]
        public static extern bool TerminateProcess(IntPtr hProcess, uint exitCode);
    }

    /// <summary>
    /// TARGETED scrub: only deletes Resiliency entries whose binary value
    /// contains the filename of the .xll currently hosting this Pilot
    /// instance. Office's protection of every OTHER addin on the system is
    /// preserved. The caller (Pilot.Start) already knows we're inside an
    /// XRai-enabled addin (the very fact that Pilot.Start is running proves
    /// it), so clearing Resiliency for OUR specific .xll is defensible.
    ///
    /// Without this, the .NET 8 + Excel-DNA + WPF teardown FailFast
    /// (0x80131506) writes a Resiliency entry that on the NEXT launch causes
    ///   1. "Excel didn't shut down properly — Safe Mode?" prompt
    ///   2. Auto-disable of the addin (DisabledItems)
    /// Both are agentic-workflow blockers. The CLR crash itself remains
    /// logged in Event Viewer (cosmetic) until the underlying runtime issue
    /// is fixed by Microsoft / Excel-DNA.
    /// </summary>
    private static void ScrubOfficeResiliency()
    {
        string? xllPath = null;
        try { xllPath = ExcelDna.Integration.ExcelDnaUtil.XllPath; } catch { }
        if (string.IsNullOrEmpty(xllPath)) return;
        var fileName = System.IO.Path.GetFileName(xllPath).ToLowerInvariant();
        if (fileName.Length == 0) return;
        var needle = System.Text.Encoding.Unicode.GetBytes(fileName);

        string[] apps = { "Excel", "Word", "PowerPoint" };
        string[] versions = { "16.0", "15.0", "14.0" };
        foreach (var app in apps)
        {
            foreach (var ver in versions)
            {
                var keyPath = $@"Software\Microsoft\Office\{ver}\{app}\Resiliency";
                try
                {
                    using var rootKey = Microsoft.Win32.Registry.CurrentUser.OpenSubKey(keyPath);
                    if (rootKey == null) continue;
                    foreach (var bucketName in rootKey.GetSubKeyNames())
                    {
                        try
                        {
                            using var bucket = Microsoft.Win32.Registry.CurrentUser.OpenSubKey(
                                $@"{keyPath}\{bucketName}", writable: true);
                            if (bucket == null) continue;
                            foreach (var valName in bucket.GetValueNames())
                            {
                                try
                                {
                                    if (bucket.GetValue(valName) is byte[] raw &&
                                        BlobContainsLowercased(raw, needle))
                                    {
                                        bucket.DeleteValue(valName, throwOnMissingValue: false);
                                    }
                                }
                                catch { }
                            }
                        }
                        catch { }
                    }
                }
                catch { /* missing key, no permission, etc. — silent */ }
            }
        }
    }

    private static bool BlobContainsLowercased(byte[] haystack, byte[] needleLower)
    {
        if (needleLower.Length == 0 || haystack.Length < needleLower.Length) return false;
        for (int i = 0; i + needleLower.Length <= haystack.Length; i += 2)
        {
            bool match = true;
            for (int j = 0; j < needleLower.Length; j += 2)
            {
                ushort hChar = (ushort)(haystack[i + j] | (haystack[i + j + 1] << 8));
                ushort nChar = (ushort)(needleLower[j] | (needleLower[j + 1] << 8));
                if (hChar < 128) hChar = char.ToLowerInvariant((char)hChar);
                if (hChar != nChar) { match = false; break; }
            }
            if (match) return true;
        }
        return false;
    }

    /// <summary>
    /// Heuristic: is the host process currently being torn down? We only
    /// engage the strict-shutdown bypass when this returns true, so
    /// mid-session calls to Pilot.Shutdown (e.g. addin reload without
    /// process exit) don't kill Excel.
    /// </summary>
    private static bool IsHostExiting()
    {
        try
        {
            // Excel is the textbook case. If we're hosted inside EXCEL.EXE,
            // and AutoClose has been called, the process IS shutting down.
            var procName = System.Diagnostics.Process.GetCurrentProcess().ProcessName;
            if (procName.Equals("EXCEL", StringComparison.OrdinalIgnoreCase) ||
                procName.Equals("WINWORD", StringComparison.OrdinalIgnoreCase) ||
                procName.Equals("POWERPNT", StringComparison.OrdinalIgnoreCase))
                return true;
        }
        catch { }
        return false;
    }

    /// <summary>
    /// Internal: registered by Expose so Shutdown can shut every dispatcher
    /// down deterministically. Idempotent — re-exposing the same dispatcher
    /// does not duplicate the entry.
    /// </summary>
    internal static void TrackDispatcher(System.Windows.Threading.Dispatcher dispatcher)
    {
        if (dispatcher == null) return;
        lock (_trackedDispatchers)
        {
            // Drop dead refs while we're here.
            for (int i = _trackedDispatchers.Count - 1; i >= 0; i--)
            {
                if (!_trackedDispatchers[i].TryGetTarget(out var existing))
                {
                    _trackedDispatchers.RemoveAt(i);
                    continue;
                }
                if (ReferenceEquals(existing, dispatcher)) return; // already tracked
            }
            _trackedDispatchers.Add(new WeakReference<System.Windows.Threading.Dispatcher>(dispatcher));
        }
    }

    /// <summary>
    /// Expose a WPF control (typically a UserControl / task pane) for inspection and interaction.
    /// Walks the visual tree and registers all named controls.
    /// </summary>
    public static void Expose(FrameworkElement element)
    {
        // Capture the WPF dispatcher from the element's thread
        _server?.SetDispatcher(element.Dispatcher);

        // Track this dispatcher so Pilot.Shutdown can InvokeShutdown + Join
        // its thread on AutoClose. Critical for clean .NET 8 / Excel-DNA
        // teardown when the addin owns a dedicated WPF thread that's blocked
        // on Dispatcher.Run().
        TrackDispatcher(element.Dispatcher);

        // Clear the old registry (and dispose its adapters' value-change
        // subscriptions) before walking the new visual tree. Prevents stale
        // ControlAdapter instances from holding references to a disposed tree.
        _controls.Clear();
        _controls.RootElement = element;
        ControlDiscovery.Walk(element, _controls);
        _lastExposeAt = DateTime.UtcNow;
        _totalExposeCalls++;
        Debug.WriteLine($"XRai: Exposed {_controls.Count} controls from {element.GetType().Name}");

        // Emit a pane.exposed event so Studio sees the initial state — this
        // is the primary trigger for the dashboard to render the control tree.
        try
        {
            _server?.PushEvent("pane.exposed", new
            {
                rootType = element.GetType().Name,
                controlCount = _controls.Count,
                controls = _controls.All.Select(kvp => new
                {
                    name = kvp.Key,
                    type = kvp.Value.Type,
                    enabled = kvp.Value.IsEnabled,
                    visible = kvp.Value.IsVisible,
                }).ToArray(),
            });
        }
        catch (Exception ex) { Debug.WriteLine($"pane.exposed emit failed: {ex.Message}"); }
    }

    /// <summary>
    /// Expose a WinForms control (typically a Form or UserControl) for inspection and interaction.
    /// Walks the control tree and registers all named controls.
    /// </summary>
    public static void Expose(System.Windows.Forms.Control control)
    {
        if (control == null) throw new ArgumentNullException(nameof(control));
        _controls.Clear();
        WinFormsDiscovery.Walk(control, _controls);
        _lastExposeAt = DateTime.UtcNow;
        _totalExposeCalls++;
        Debug.WriteLine($"XRai: Exposed {_controls.Count} controls from WinForms {control.GetType().Name}");

        try
        {
            _server?.PushEvent("pane.exposed", new
            {
                rootType = control.GetType().Name,
                controlCount = _controls.Count,
                framework = "WinForms",
                controls = _controls.All.Select(kvp => new
                {
                    name = kvp.Key,
                    type = kvp.Value.Type,
                    enabled = kvp.Value.IsEnabled,
                    visible = kvp.Value.IsVisible,
                }).ToArray(),
            });
        }
        catch (Exception ex) { Debug.WriteLine($"pane.exposed emit failed: {ex.Message}"); }
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

        // Snapshot the initial property dictionary and fire model.exposed
        // so Studio can render the ViewModel inspector immediately.
        try
        {
            var adapter = _models.Default;
            var initialProps = adapter?.GetAll() ?? new Dictionary<string, object?>();
            _server?.PushEvent("model.exposed", new
            {
                name = key,
                modelType = viewModel.GetType().Name,
                properties = initialProps,
            });
        }
        catch (Exception ex) { Debug.WriteLine($"model.exposed emit failed: {ex.Message}"); }
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
