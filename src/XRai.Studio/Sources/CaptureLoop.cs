using System.Diagnostics;
using System.Text.Json.Nodes;
using XRai.Vision;

namespace XRai.Studio.Sources;

/// <summary>
/// Streams JPEG screenshots of a target window at a configurable rate to the
/// Studio event bus. Dedicated background thread — never blocks the pipe
/// server, the STA worker, or the Kestrel server.
///
/// Adaptive behavior:
///   - Default capture interval is 250ms (4 fps)
///   - If no WebSocket clients are connected for 60 seconds, downshifts
///     to 2000ms (0.5 fps) — zero cost when nobody is watching
///   - Retries window-handle lookup if the previous HWND is no longer valid
///     (e.g. Excel was killed and relaunched by rebuild)
///   - Skips a tick if the prior capture is still inflight (rare)
/// </summary>
public sealed class CaptureLoop : IDisposable
{
    private readonly EventBus _bus;
    private readonly Func<nint?> _hwndProvider;
    private readonly CancellationTokenSource _cts = new();
    private Thread? _thread;
    private volatile bool _disposed;
    private long _lastCaptureTicks;
    private long _lastClientSeenTicks;

    /// <summary>Default 4 fps. Can be overridden via constructor for slow machines.</summary>
    public int ActiveIntervalMs { get; set; } = 250;

    /// <summary>Downshift interval when idle (0.5 fps).</summary>
    public int IdleIntervalMs { get; set; } = 2000;

    /// <summary>After this long with no subscribers, downshift.</summary>
    public int IdleAfterMs { get; set; } = 60_000;

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
        _lastClientSeenTicks = Environment.TickCount64;
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
                var interval = ComputeInterval();
                var sleep = interval - (int)Math.Min(interval, Environment.TickCount64 - _lastCaptureTicks);
                if (sleep > 0) Thread.Sleep(sleep);
                if (_disposed) break;

                // Skip frame entirely when there are no subscribers —
                // publishing into the bus is cheap but capturing a window
                // is not, so we save cycles.
                if (_bus.SubscriberCount == 0)
                {
                    // Mark the idle clock but don't reset it
                    _lastCaptureTicks = Environment.TickCount64;
                    continue;
                }

                _lastClientSeenTicks = Environment.TickCount64;

                var hwnd = _hwndProvider();
                if (!hwnd.HasValue || hwnd.Value == nint.Zero)
                {
                    _lastCaptureTicks = Environment.TickCount64;
                    continue;
                }

                var captureStart = Environment.TickCount64;
                var jpeg = Capture.CaptureHwndToJpeg(hwnd.Value, crop: null, quality: JpegQuality);
                _lastCaptureTicks = Environment.TickCount64;

                if (jpeg == null || jpeg.Length == 0) continue;

                var b64 = Convert.ToBase64String(jpeg);
                var data = new JsonObject
                {
                    ["mime"] = "image/jpeg",
                    ["b64"] = b64,
                    ["bytes"] = jpeg.Length,
                    ["capture_ms"] = Environment.TickCount64 - captureStart,
                };
                _bus.Publish(StudioEvent.Now("frame", "capture", data));
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"CaptureLoop iteration failed: {ex.Message}");
                Thread.Sleep(500);
            }
        }
    }

    private int ComputeInterval()
    {
        // Downshift when idle
        if (_bus.SubscriberCount == 0)
        {
            return IdleIntervalMs;
        }

        var idleFor = Environment.TickCount64 - _lastClientSeenTicks;
        if (idleFor > IdleAfterMs)
            return IdleIntervalMs;

        return ActiveIntervalMs;
    }

    public void Dispose()
    {
        if (_disposed) return;
        _disposed = true;
        try { _cts.Cancel(); } catch { }
        try { _thread?.Join(1000); } catch { }
        _cts.Dispose();
    }
}
