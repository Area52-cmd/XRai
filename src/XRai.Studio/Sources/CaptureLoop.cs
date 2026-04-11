using System.Diagnostics;
using System.Text.Json.Nodes;
using XRai.Vision;

namespace XRai.Studio.Sources;

/// <summary>
/// Streams JPEG screenshots of a target window at a configurable rate to the
/// Studio event bus. Dedicated background thread — never blocks the pipe
/// server, the STA worker, or the Kestrel server.
///
/// Behavior:
///   - Active rate (4 fps) when at least one WebSocket client is subscribed
///   - Idle rate (0.5 fps) when nobody is watching — near-zero CPU
///   - Hash-and-skip: if the captured JPEG bytes are identical to the prior
///     frame, the publish is skipped entirely (no base64 alloc, no WebSocket
///     traffic). Cuts idle-attached bandwidth from ~160 KB/s to near-zero
///     when the target window isn't changing.
///   - Re-probes the window handle every frame so it survives Excel kills /
///     relaunches without re-instantiating the loop.
/// </summary>
public sealed class CaptureLoop : IDisposable
{
    private readonly EventBus _bus;
    private readonly Func<nint?> _hwndProvider;
    private readonly CancellationTokenSource _cts = new();
    private Thread? _thread;
    private volatile bool _disposed;
    private long _lastCaptureTicks;
    private ulong _lastJpegHash;

    /// <summary>Default 4 fps. Can be overridden via constructor for slow machines.</summary>
    public int ActiveIntervalMs { get; set; } = 250;

    /// <summary>Downshift interval when no clients are watching (0.5 fps).
    /// The loop still runs but doesn't actually capture, so the daemon
    /// uses near-zero CPU when nothing is observing.</summary>
    public int IdleIntervalMs { get; set; } = 2000;

    /// <summary>JPEG quality 1-100. 70 ≈ 40 KB per 1280x720 frame.</summary>
    public int JpegQuality { get; set; } = 70;

    public CaptureLoop(EventBus bus, Func<nint?> hwndProvider)
    {
        _bus = bus;
        _hwndProvider = hwndProvider;
    }

    public void Start()
    {
        if (_thread != null) return;
        _thread = new Thread(Loop)
        {
            IsBackground = true,
            Name = "xrai-studio-capture-loop"
        };
        _thread.Start();
    }

    private void Loop()
    {
        while (!_disposed && !_cts.IsCancellationRequested)
        {
            try
            {
                // Active when at least one client is connected, idle otherwise.
                // No grace period — when the last client disconnects we
                // immediately drop to the slow rate so we're not burning
                // CPU on screenshots nobody is watching.
                var hasClients = _bus.SubscriberCount > 0;
                var interval = hasClients ? ActiveIntervalMs : IdleIntervalMs;

                // Sleep until the next tick boundary, accounting for the
                // time the previous capture took.
                var sinceLast = (int)Math.Min(int.MaxValue, Environment.TickCount64 - _lastCaptureTicks);
                var sleep = Math.Max(0, interval - sinceLast);
                if (sleep > 0) Thread.Sleep(sleep);
                if (_disposed) break;

                _lastCaptureTicks = Environment.TickCount64;

                // Skip the actual capture (but not the loop) when there
                // are no subscribers. Cheap idle.
                if (!hasClients) continue;

                var hwnd = _hwndProvider();
                if (!hwnd.HasValue || hwnd.Value == nint.Zero) continue;

                var captureStart = Environment.TickCount64;
                var jpeg = Capture.CaptureHwndToJpeg(hwnd.Value, crop: null, quality: JpegQuality);
                if (jpeg == null || jpeg.Length == 0) continue;

                // Hash-and-skip: identical frames are common when the target
                // window is idle. Compute a fast 64-bit hash and skip the
                // base64 alloc + bus publish if the JPEG bytes are unchanged.
                var hash = FastHash(jpeg);
                if (hash == _lastJpegHash) continue;
                _lastJpegHash = hash;

                var b64 = Convert.ToBase64String(jpeg);
                _bus.Publish(StudioEvent.Now("frame", "capture", new JsonObject
                {
                    ["mime"] = "image/jpeg",
                    ["b64"] = b64,
                    ["bytes"] = jpeg.Length,
                    ["capture_ms"] = Environment.TickCount64 - captureStart,
                }));
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"CaptureLoop iteration failed: {ex.Message}");
                Thread.Sleep(500);
            }
        }
    }

    public void Dispose()
    {
        if (_disposed) return;
        _disposed = true;
        try { _cts.Cancel(); } catch { }
        try { _thread?.Join(1000); } catch { }
        _cts.Dispose();
    }

    /// <summary>
    /// Fast non-cryptographic 64-bit hash for change detection. Samples
    /// the buffer at 32-byte stride so the cost is constant regardless of
    /// JPEG size. Collisions are acceptable — worst case is publishing
    /// two consecutive identical-looking frames once in a blue moon.
    /// FNV-1a-style with a length salt.
    /// </summary>
    private static ulong FastHash(byte[] buf)
    {
        if (buf == null || buf.Length == 0) return 0;
        const ulong fnvOffset = 14695981039346656037;
        const ulong fnvPrime = 1099511628211;
        ulong h = fnvOffset ^ (ulong)buf.Length;
        // Walk the buffer in ~256 strided reads + every byte of the first
        // and last 64 bytes (catches header and trailer changes).
        int n = buf.Length;
        int stride = Math.Max(32, n / 256);
        for (int i = 0; i < n; i += stride)
        {
            h ^= buf[i];
            h *= fnvPrime;
        }
        for (int i = 0; i < Math.Min(64, n); i++)
        {
            h ^= buf[i];
            h *= fnvPrime;
        }
        for (int i = Math.Max(0, n - 64); i < n; i++)
        {
            h ^= buf[i];
            h *= fnvPrime;
        }
        return h;
    }
}
