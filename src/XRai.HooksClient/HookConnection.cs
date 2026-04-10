using System.Diagnostics;
using System.IO.Pipes;
using System.Text.Json;
using System.Text.Json.Nodes;
using XRai.Core;

namespace XRai.HooksClient;

public class HookConnection : IDisposable
{
    private NamedPipeClientStream? _pipe;
    private StreamReader? _reader;
    private StreamWriter? _writer;
    private bool _disposed;
    private bool _lastAutoReconnectAttempted;
    private string? _lastAutoReconnectError;

    public bool IsConnected => _pipe?.IsConnected == true;
    public string? PipeName { get; private set; }

    /// <summary>
    /// When true (default), SendCommand will transparently try to (re)connect to
    /// the Excel hooks pipe if no connection is active. This is the expected
    /// behavior for agent-driven workflows where every XRai.Tool.exe invocation
    /// is a fresh process and can't carry hook pipe state across boundaries.
    /// </summary>
    public bool AutoReconnect { get; set; } = true;

    public void Connect(int excelPid, int timeoutMs = 5000)
    {
        Disconnect();

        PipeName = $"xrai_{excelPid}";
        _pipe = new NamedPipeClientStream(".", PipeName, PipeDirection.InOut);
        _pipe.Connect(timeoutMs);
        _reader = new StreamReader(_pipe);
        _writer = new StreamWriter(_pipe) { AutoFlush = true };

        // Auth handshake — read the token stored for this pipe and present it
        // as the first line. The server closes the connection immediately if the
        // token is missing or wrong (unless XRAI_ALLOW_UNAUTH=1 is set upstream).
        PerformAuthHandshake(PipeName);
    }

    /// <summary>
    /// True if the most recent Connect() skipped the auth handshake because the
    /// target process is running an older XRai.Hooks version without token auth.
    /// Surfaced in diagnostics so callers know they're in legacy mode.
    /// </summary>
    public bool LegacyAuthFallback { get; private set; }

    private void PerformAuthHandshake(string pipeName)
    {
        if (_writer == null || _reader == null || _pipe == null) return;

        LegacyAuthFallback = false;

        var token = PipeAuth.ReadToken(pipeName);
        if (string.IsNullOrEmpty(token))
        {
            // No token file — target process is running an older XRai.Hooks
            // that doesn't use token auth. Auto-fallback to legacy mode: don't
            // send a handshake line, dispatch commands directly. The older
            // server treats the first line as the first command so this is
            // wire-compatible.
            //
            // Security note: the pipe ACL (per-user + SYSTEM only) remains in
            // force regardless. Falling back to no-handshake does NOT expose
            // the pipe to other local users — it only skips the extra
            // application-level token check that old XRai.Hooks builds don't
            // know how to respond to.
            LegacyAuthFallback = true;
            return;
        }

        try
        {
            _writer.WriteLine(PipeAuth.BuildHandshakeLine(token));
            var response = _reader.ReadLine();
            if (response == null)
            {
                ForceDisconnect();
                throw new InvalidOperationException(
                    $"Hooks pipe closed during auth handshake on '{pipeName}'.");
            }

            var node = JsonNode.Parse(response);
            var ok = node?["ok"]?.GetValue<bool>() ?? false;
            if (!ok)
            {
                var err = node?["error"]?.GetValue<string>() ?? "unknown";
                var code = node?["code"]?.GetValue<string>() ?? "";
                ForceDisconnect();
                throw new InvalidOperationException(
                    $"Hooks pipe authentication failed: {err} ({code}).");
            }
        }
        catch (IOException ex)
        {
            ForceDisconnect();
            throw new InvalidOperationException(
                $"Hooks pipe I/O error during auth handshake: {ex.Message}", ex);
        }
    }

    public void Disconnect()
    {
        _writer = null;
        _reader = null;
        if (_pipe != null)
        {
            try { _pipe.Dispose(); } catch { }
            _pipe = null;
        }
        PipeName = null;
    }

    /// <summary>
    /// Try to auto-connect by discovering a running Excel process and connecting
    /// to its xrai_{pid} pipe. Returns true on success. Does not throw on failure.
    /// </summary>
    public bool TryAutoConnect(int timeoutMs = 2000)
    {
        _lastAutoReconnectAttempted = true;
        _lastAutoReconnectError = null;

        try
        {
            var procs = Process.GetProcessesByName("EXCEL");
            if (procs.Length == 0)
            {
                _lastAutoReconnectError = "No Excel process found";
                return false;
            }

            Exception? lastEx = null;
            foreach (var p in procs)
            {
                try
                {
                    Connect(p.Id, timeoutMs);
                    if (IsConnected) return true;
                }
                catch (Exception ex)
                {
                    lastEx = ex;
                }
            }

            _lastAutoReconnectError = lastEx?.Message ?? "No responding xrai_{pid} pipe found";
            return false;
        }
        catch (Exception ex)
        {
            _lastAutoReconnectError = ex.Message;
            return false;
        }
    }

    /// <summary>
    /// Timeout for ReadLine when waiting for a hooks response. Prevents indefinite
    /// hangs when the UI thread is blocked inside a modal dialog (ShowDialog).
    /// Default: 30s. Set to 0 for no timeout (legacy behavior).
    /// </summary>
    public int ReadTimeoutMs { get; set; } = 30_000;

    public JsonNode? SendCommand(string cmd, object? data = null)
    {
        // Auto-reconnect on demand: the CLI is stateless (every invocation is a
        // fresh process) so there's no pipe handle to carry across calls. If
        // hooks aren't connected and auto-reconnect is enabled, try to find
        // the running Excel and connect transparently.
        if (!IsConnected && AutoReconnect)
        {
            TryAutoConnect();
        }

        EnsureConnected();

        var obj = new JsonObject { ["cmd"] = cmd };
        if (data != null)
        {
            var json = JsonSerializer.Serialize(data);
            var parsed = JsonNode.Parse(json);
            if (parsed is JsonObject extra)
            {
                foreach (var kvp in extra)
                    obj[kvp.Key] = kvp.Value?.DeepClone();
            }
        }

        return SendRaw(obj, retryOnFailure: AutoReconnect);
    }

    private JsonNode? SendRaw(JsonObject obj, bool retryOnFailure)
    {
        try
        {
            _writer!.WriteLine(obj.ToJsonString());

            // Read with timeout to prevent indefinite hangs when the Hooks
            // UI thread is blocked inside a modal dialog (ShowDialog).
            string? response;
            if (ReadTimeoutMs > 0)
            {
                var readTask = Task.Run(() => _reader!.ReadLine());
                if (!readTask.Wait(ReadTimeoutMs))
                {
                    // Timeout — pipe is likely stuck behind a modal dialog.
                    // Mark disconnected so subsequent calls auto-reconnect.
                    ForceDisconnect();
                    return null;
                }
                response = readTask.Result;
            }
            else
            {
                response = _reader!.ReadLine();
            }

            if (response == null)
            {
                // Pipe closed (server-side disconnect). Mark as disconnected.
                ForceDisconnect();

                // Auto-reconnect + single retry if allowed.
                if (retryOnFailure && TryAutoConnect())
                {
                    return SendRaw(obj, retryOnFailure: false); // one retry only
                }
                return null;
            }

            return JsonNode.Parse(response);
        }
        catch (Exception ex) when (ex is IOException or ObjectDisposedException or InvalidOperationException)
        {
            // Pipe broken (Excel crashed, process killed, pipe server died).
            ForceDisconnect();

            // Auto-reconnect + single retry if allowed.
            if (retryOnFailure && TryAutoConnect())
            {
                return SendRaw(obj, retryOnFailure: false);
            }
            return null;
        }
    }

    private void ForceDisconnect()
    {
        try { _pipe?.Dispose(); } catch { }
        _pipe = null;
        _writer = null;
        _reader = null;
    }

    public string? ReadLine()
    {
        EnsureConnected();
        return _reader!.ReadLine();
    }

    private void EnsureConnected()
    {
        if (IsConnected) return;

        var baseMsg = "Hooks pipe not connected.";
        if (_lastAutoReconnectAttempted && _lastAutoReconnectError != null)
            throw new InvalidOperationException(
                $"{baseMsg} Auto-reconnect failed: {_lastAutoReconnectError}. " +
                "Ensure Excel is running with an add-in that calls Pilot.Start(). " +
                "Run 'XRai.Tool.exe doctor' to diagnose.");

        throw new InvalidOperationException(
            $"{baseMsg} Call {{\"cmd\":\"connect\"}} or {{\"cmd\":\"attach\"}} first, " +
            "or ensure Excel is running with a XRai.Hooks-enabled add-in loaded.");
    }

    public void Dispose()
    {
        if (_disposed) return;
        _disposed = true;
        Disconnect();
    }
}
