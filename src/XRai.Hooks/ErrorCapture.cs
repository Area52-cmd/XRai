namespace XRai.Hooks;

public static class ErrorCapture
{
    private static PipeServer? _server;
    private static bool _installed;

    public static void Install(PipeServer server)
    {
        _server = server;
        if (_installed) return;
        AppDomain.CurrentDomain.UnhandledException += OnUnhandledException;
        _installed = true;
    }

    /// <summary>
    /// Critical for Excel-DNA / .NET 8 hosted scenarios. The
    /// AppDomain.UnhandledException subscription points into THIS assembly;
    /// if it survives Pilot.Stop, the CLR cannot drain the rooted delegate
    /// when Excel-DNA unloads the addin's AssemblyLoadContext, which causes
    /// CoreCLR to FailFast with 0x80131506 (COR_E_EXECUTIONENGINE) on
    /// Excel shutdown. Pilot.Stop now calls Uninstall to clear it.
    /// </summary>
    public static void Uninstall()
    {
        if (!_installed) return;
        try { AppDomain.CurrentDomain.UnhandledException -= OnUnhandledException; } catch { }
        _server = null;
        _installed = false;
    }

    private static void OnUnhandledException(object sender, UnhandledExceptionEventArgs e)
    {
        if (e.ExceptionObject is Exception ex)
        {
            _server?.PushEvent("error", new
            {
                exception = ex.GetType().Name,
                message = ex.Message,
                stack = ex.StackTrace,
                timestamp = DateTime.UtcNow.ToString("o"),
            });
        }
    }
}
