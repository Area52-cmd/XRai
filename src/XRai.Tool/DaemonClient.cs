using System.IO.Pipes;
using System.Text.Json.Nodes;
using XRai.Core;

namespace XRai.Tool;

// PipeAuth lives in XRai.Core — referenced via the ProjectReference below.

/// <summary>
/// Thin stdin→daemon pipe→stdout forwarder. When a daemon is running, a regular
/// XRai.Tool.exe invocation becomes a transparent passthrough: every JSON line
/// from stdin is copied to the daemon pipe, every response line from the daemon
/// is copied to stdout. The daemon serializes actual command execution in a
/// single queue, so rapid successive CLI calls never cause COM races.
///
/// STALE DAEMON DETECTION: before forwarding, the client pings the daemon and
/// compares its build_version against the local binary's version. If they don't
/// match, the daemon is running older code than what's on disk — the client
/// auto-stops the stale daemon and falls back to direct mode (the caller can
/// manually restart the daemon if desired). This prevents the common failure mode
/// where a user ships a fix, the new binary is on disk, but commands keep running
/// through a zombie daemon with old code.
/// </summary>
public static class DaemonClient
{
    /// <summary>
    /// Connect to the running daemon, forward stdin → daemon, daemon → stdout.
    /// Returns the process exit code (0 on normal EOF, 1 on pipe failure).
    /// Returns -1 as a special sentinel when the daemon was detected as stale
    /// and should fall back to direct mode.
    /// </summary>
    public static int Run()
    {
        // Pre-flight: verify the daemon is running the same build as this client.
        // If not, stop it and signal the caller to use direct mode.
        if (!PreflightVersionCheck())
        {
            return -1; // sentinel: stale daemon, caller should use direct mode
        }

        try
        {
            using var pipe = new NamedPipeClientStream(".", DaemonServer.PipeName, PipeDirection.InOut);
            pipe.Connect(2000);

            using var reader = new StreamReader(pipe);
            using var writer = new StreamWriter(pipe) { AutoFlush = true };

            // Auth handshake: present the token before any commands. The daemon
            // rejects unauthenticated clients unless XRAI_ALLOW_UNAUTH=1 is set.
            if (!PerformAuthHandshake(DaemonServer.PipeName, reader, writer))
            {
                Console.Error.WriteLine("[xrai] Daemon authentication failed.");
                return 1;
            }

            // Use the SAME bracket-counting JSON stream parser as Repl.Run so that
            // multi-line pretty-printed JSON works identically with or without the
            // daemon. Previously this used ReadLine which broke multi-line JSON
            // — fixed in Round 9.
            foreach (var doc in JsonStreamReader.ReadDocuments(Console.In))
            {
                // Forward each complete JSON document as a single line to the daemon.
                // The daemon's pipe protocol is still newline-delimited per-document,
                // so we join any internal whitespace/newlines by stripping them during
                // the write. Actually the daemon's pipe reader uses ReadLine so we
                // must send exactly one line. Serialize the parsed document back to
                // compact JSON before forwarding.
                var compactDoc = CompactJson(doc);
                writer.WriteLine(compactDoc);

                var response = reader.ReadLine();
                if (response == null)
                {
                    Console.Error.WriteLine("[xrai] Daemon pipe closed unexpectedly");
                    return 1;
                }
                Console.Out.WriteLine(response);
                Console.Out.Flush();
            }

            return 0;
        }
        catch (TimeoutException)
        {
            Console.Error.WriteLine($"[xrai] Could not connect to daemon at {DaemonServer.PipeName} (timeout).");
            Console.Error.WriteLine("[xrai] Start the daemon with: XRai.Tool.exe --daemon");
            return 1;
        }
        catch (Exception ex)
        {
            Console.Error.WriteLine($"[xrai] Daemon client error: {ex.Message}");
            return 1;
        }
    }

    /// <summary>
    /// Ping the daemon and compare its build_version against the local binary's.
    /// If they don't match, the daemon is stale — stop it and return false so the
    /// caller can fall back to direct mode.
    /// Returns true if the daemon is alive AND matches the local version.
    /// </summary>
    private static bool PreflightVersionCheck()
    {
        try
        {
            var response = PingForResponse();
            if (response == null) return false;

            var node = JsonNode.Parse(response);
            var daemonVersion = node?["build_version"]?.GetValue<string>();
            var localVersion = DaemonServer.BuildVersion;

            if (daemonVersion == null)
            {
                // Old daemon that doesn't report build_version. Stop it.
                Console.Error.WriteLine("[xrai] Daemon is running pre-versioning build — auto-stopping to pick up new binary.");
                SendStop();
                Thread.Sleep(500);
                return false;
            }

            if (daemonVersion != localVersion)
            {
                Console.Error.WriteLine($"[xrai] Daemon build ({daemonVersion}) differs from local binary ({localVersion}). Auto-stopping stale daemon.");
                Console.Error.WriteLine("[xrai] Falling back to direct mode. Restart the daemon manually with: XRai.Tool.exe --daemon");
                SendStop();
                Thread.Sleep(500);
                return false;
            }

            return true;
        }
        catch
        {
            // If we can't even ping, assume no daemon and let normal logic handle it
            return false;
        }
    }

    /// <summary>
    /// Read the daemon auth token from disk, send it as the first line on the pipe,
    /// and validate the server's response. Returns true if the handshake succeeded
    /// (or was waived via XRAI_ALLOW_UNAUTH=1). Returns false if auth failed — the
    /// caller should abort.
    /// </summary>
    private static bool PerformAuthHandshake(string pipeName, StreamReader reader, StreamWriter writer)
    {
        var token = PipeAuth.ReadToken(pipeName);
        if (string.IsNullOrEmpty(token))
        {
            if (PipeAuth.AllowUnauthenticated)
            {
                Console.Error.WriteLine("[xrai] WARNING: XRAI_ALLOW_UNAUTH=1 set and no token file found — connecting without auth.");
                // Don't send anything; the daemon's legacy fallback will treat
                // the first command line as the handshake.
                return true;
            }

            var tokenPath = PipeAuth.GetTokenFilePath(pipeName);
            Console.Error.WriteLine($"[xrai] Auth token file not found at '{tokenPath}'.");
            Console.Error.WriteLine("[xrai] The daemon may not be running, or may be running a pre-auth build.");
            Console.Error.WriteLine("[xrai] Start the daemon with: XRai.Tool.exe --daemon");
            return false;
        }

        try
        {
            writer.WriteLine(PipeAuth.BuildHandshakeLine(token));
            var response = reader.ReadLine();
            if (response == null)
            {
                Console.Error.WriteLine("[xrai] Daemon closed the pipe during auth handshake.");
                return false;
            }

            var node = JsonNode.Parse(response);
            var ok = node?["ok"]?.GetValue<bool>() ?? false;
            if (!ok)
            {
                var err = node?["error"]?.GetValue<string>() ?? "unknown";
                var code = node?["code"]?.GetValue<string>() ?? "";
                Console.Error.WriteLine($"[xrai] Daemon auth rejected: {err} ({code}).");
                return false;
            }
            return true;
        }
        catch (Exception ex)
        {
            Console.Error.WriteLine($"[xrai] Auth handshake I/O error: {ex.Message}");
            return false;
        }
    }

    /// <summary>
    /// Ping the daemon and return the raw response JSON line. Null on failure.
    /// </summary>
    private static string? PingForResponse()
    {
        try
        {
            using var pipe = new NamedPipeClientStream(".", DaemonServer.PipeName, PipeDirection.InOut);
            pipe.Connect(500);
            using var reader = new StreamReader(pipe);
            using var writer = new StreamWriter(pipe) { AutoFlush = true };
            writer.WriteLine("__daemon_ping__");
            return reader.ReadLine();
        }
        catch
        {
            return null;
        }
    }

    /// <summary>
    /// Ping the daemon to check if it's alive. Returns true if a response was received.
    /// </summary>
    public static bool Ping()
    {
        var response = PingForResponse();
        return response != null && response.Contains("alive");
    }

    /// <summary>
    /// Convert a possibly-multi-line JSON document to a single-line compact form
    /// for transmission over the daemon pipe (which uses newline-delimited framing).
    /// Preserves the JSON semantically — just strips whitespace outside string literals.
    /// </summary>
    private static string CompactJson(string json)
    {
        try
        {
            using var doc = System.Text.Json.JsonDocument.Parse(json);
            return System.Text.Json.JsonSerializer.Serialize(doc.RootElement);
        }
        catch
        {
            // If parsing fails, send as-is and let the daemon report the error.
            // Strip newlines as a last resort to avoid framing issues.
            return json.Replace("\r", " ").Replace("\n", " ");
        }
    }

    /// <summary>
    /// Send a stop signal to the running daemon. Returns true if accepted.
    /// </summary>
    public static bool SendStop()
    {
        try
        {
            using var pipe = new NamedPipeClientStream(".", DaemonServer.PipeName, PipeDirection.InOut);
            pipe.Connect(1000);
            using var reader = new StreamReader(pipe);
            using var writer = new StreamWriter(pipe) { AutoFlush = true };
            writer.WriteLine("__daemon_stop__");
            var response = reader.ReadLine();
            return response != null && response.Contains("stopping");
        }
        catch
        {
            return false;
        }
    }
}
