using System.Diagnostics;

namespace XRai.Hooks;

public static class LogCapture
{
    private static HooksTraceListener? _listener;

    public static void Install(PipeServer server)
    {
        _listener = new HooksTraceListener(server);
        Trace.Listeners.Add(_listener);
    }

    public static void Uninstall()
    {
        if (_listener != null)
        {
            Trace.Listeners.Remove(_listener);
            _listener = null;
        }
    }

    private class HooksTraceListener : TraceListener
    {
        private readonly PipeServer _server;

        public HooksTraceListener(PipeServer server)
        {
            _server = server;
        }

        public override void Write(string? message)
        {
            if (!string.IsNullOrEmpty(message))
            {
                _server.PushEvent("log", new
                {
                    message,
                    source = "Trace",
                    timestamp = DateTime.UtcNow.ToString("o"),
                });
            }
        }

        public override void WriteLine(string? message)
        {
            Write(message);
        }
    }
}
