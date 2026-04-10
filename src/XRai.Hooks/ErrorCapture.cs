namespace XRai.Hooks;

public static class ErrorCapture
{
    private static PipeServer? _server;

    public static void Install(PipeServer server)
    {
        _server = server;
        AppDomain.CurrentDomain.UnhandledException += OnUnhandledException;
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
