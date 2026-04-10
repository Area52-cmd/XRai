using System.Diagnostics;
using System.Drawing;
using System.Drawing.Imaging;
using System.Runtime.InteropServices;
using System.Text;
using System.Text.Json.Nodes;
using XRai.Core;

namespace XRai.Vision;

public class Capture
{
    /// <summary>
    /// Optional: provides the COM Application.Hwnd so screenshot uses the
    /// correct XLMAIN window (the one with the active workbook) instead of
    /// guessing from Process.MainWindowHandle.
    /// Set by the caller (Program.cs/DaemonServer) after ExcelSession.Attach:
    ///   capture.SetComHwndProvider(() => (nint)session.App.Hwnd);
    /// </summary>
    private Func<nint?>? _comHwndProvider;

    public void SetComHwndProvider(Func<nint?> provider) => _comHwndProvider = provider;

    // ── Win32 P/Invoke ────────────────────────────────────────────────

    [DllImport("user32.dll")]
    private static extern bool PrintWindow(nint hwnd, nint hdcBlt, uint nFlags);

    [DllImport("user32.dll")]
    private static extern bool GetWindowRect(nint hwnd, out RECT lpRect);

    [DllImport("user32.dll")]
    private static extern nint GetForegroundWindow();

    [DllImport("user32.dll")]
    private static extern bool IsWindowVisible(nint hWnd);

    [DllImport("user32.dll")]
    private static extern bool IsWindow(nint hWnd);

    [DllImport("user32.dll", CharSet = CharSet.Unicode)]
    private static extern int GetWindowText(nint hWnd, StringBuilder lpString, int nMaxCount);

    [DllImport("user32.dll", CharSet = CharSet.Unicode)]
    private static extern int GetClassName(nint hWnd, StringBuilder lpClassName, int nMaxCount);

    [DllImport("user32.dll")]
    private static extern uint GetWindowThreadProcessId(nint hWnd, out uint lpdwProcessId);

    private delegate bool EnumWindowsProc(nint hWnd, nint lParam);

    [DllImport("user32.dll")]
    private static extern bool EnumWindows(EnumWindowsProc lpEnumFunc, nint lParam);

    private const uint PW_RENDERFULLCONTENT = 2;

    [StructLayout(LayoutKind.Sequential)]
    private struct RECT
    {
        public int Left, Top, Right, Bottom;
        public int Width => Right - Left;
        public int Height => Bottom - Top;
    }

    public void Register(CommandRouter router)
    {
        router.Register("screenshot", HandleScreenshot);
    }

    /// <summary>
    /// Capture a screenshot. Modes:
    ///   mode:"window" (default)    — capture Excel main window only (today's behavior)
    ///   mode:"main_plus_modal"     — capture Excel main window + any top-level modal dialogs composited (auto)
    ///   mode:"all_windows"         — composite ALL visible top-level windows owned by the Excel process
    ///   mode:"foreground"          — capture whichever window is currently in the foreground
    ///   hwnd:<int>                 — capture a specific window by handle (use win32.dialog.list to get handles)
    ///
    /// Default mode is "main_plus_modal" so the common case (Excel + its open dialog)
    /// is captured without the agent having to know about modes.
    /// </summary>
    private string HandleScreenshot(JsonObject args)
    {
        var mode = args["mode"]?.GetValue<string>() ?? "main_plus_modal";
        var hwndArg = args["hwnd"]?.GetValue<long>();
        var path = args["path"]?.GetValue<string>();

        // Specific HWND mode bypasses everything
        if (hwndArg.HasValue)
        {
            return CaptureSingleWindow((nint)hwndArg.Value, path, "hwnd");
        }

        // Foreground mode
        if (mode == "foreground")
        {
            var fg = GetForegroundWindow();
            if (fg == nint.Zero)
                return Response.Error("No foreground window");
            return CaptureSingleWindow(fg, path, "foreground");
        }

        // All modes below need Excel running
        var procs = Process.GetProcessesByName("EXCEL");
        if (procs.Length == 0)
            return Response.Error("No Excel process found");
        var excelPids = procs.Select(p => (uint)p.Id).ToHashSet();

        // Try to get the COM-provided Hwnd (the correct window for the active workbook)
        nint? comHwnd = null;
        try { comHwnd = _comHwndProvider?.Invoke(); } catch { }

        // Legacy "window" mode — Excel main window only
        if (mode == "window")
        {
            var mainHwnd = comHwnd ?? procs[0].MainWindowHandle;
            if (mainHwnd == nint.Zero)
                return Response.Error("Excel main window handle not found");
            return CaptureSingleWindow(mainHwnd, path, comHwnd.HasValue ? "com_hwnd" : "main_window");
        }

        // Enumerate all Excel-owned top-level windows
        var windows = EnumerateExcelWindows(excelPids);
        if (windows.Count == 0)
            return Response.Error("No visible Excel windows found");

        if (mode == "main_plus_modal")
        {
            // Filter to the correct main window + any dialog-class windows.
            // Use the COM Hwnd to identify the RIGHT XLMAIN window (the one
            // showing the active workbook) when multiple SDI windows exist
            // in the same process.
            var filtered = new List<WindowInfo>();
            foreach (var w in windows)
            {
                if (w.ClassName == "XLMAIN")
                {
                    // If we have the COM Hwnd, only include the matching XLMAIN
                    if (comHwnd.HasValue)
                    {
                        if (w.Handle == comHwnd.Value)
                            filtered.Add(w);
                        // Skip non-matching XLMAIN windows (blank/start screen)
                    }
                    else
                    {
                        filtered.Add(w); // No COM Hwnd — include all
                    }
                }
                else if (w.ClassName == "#32770" ||
                         w.ClassName == "NUIDialog" ||
                         w.ClassName.StartsWith("bosa_sdm_", StringComparison.OrdinalIgnoreCase) ||
                         !string.IsNullOrWhiteSpace(w.Title))
                {
                    filtered.Add(w); // Always include dialogs
                }
            }

            // Fallback: if COM Hwnd filtering excluded ALL XLMAIN windows
            // (e.g., COM Hwnd doesn't match any enumerated window), include all
            if (!filtered.Any(w => w.ClassName == "XLMAIN"))
            {
                filtered = windows.Where(w =>
                    w.ClassName == "XLMAIN" ||
                    w.ClassName == "#32770" ||
                    w.ClassName == "NUIDialog" ||
                    w.ClassName.StartsWith("bosa_sdm_", StringComparison.OrdinalIgnoreCase) ||
                    !string.IsNullOrWhiteSpace(w.Title)).ToList();
            }

            return CaptureComposite(filtered, path, comHwnd.HasValue ? "com_hwnd_plus_modal" : "main_plus_modal");
        }

        if (mode == "all_windows")
        {
            return CaptureComposite(windows, path, "all_windows");
        }

        return Response.Error($"Unknown screenshot mode: '{mode}'. Valid: window | main_plus_modal | all_windows | foreground. Or pass hwnd:<handle>.");
    }

    // ── Public helper for Studio / streaming clients ─────────────────

    /// <summary>
    /// Capture a window by HWND into an in-memory JPEG byte array. Used by
    /// XRai.Studio's CaptureLoop to stream continuous frames to the dashboard.
    /// No file IO — the bitmap is encoded and returned directly.
    /// </summary>
    /// <param name="hwnd">Target window handle. Must be valid.</param>
    /// <param name="crop">Optional crop rectangle in client coordinates (relative to the window's top-left). Null = full window.</param>
    /// <param name="quality">JPEG quality 1-100. Studio default is 70 for ~40 KB per 1280x720 frame.</param>
    /// <returns>JPEG bytes, or null on failure (invalid hwnd, zero-size window, encode error).</returns>
    public static byte[]? CaptureHwndToJpeg(nint hwnd, Rectangle? crop = null, int quality = 70)
    {
        if (!IsWindow(hwnd)) return null;
        if (!GetWindowRect(hwnd, out var rect)) return null;
        if (rect.Width <= 0 || rect.Height <= 0) return null;

        try
        {
            using var bmp = new Bitmap(rect.Width, rect.Height);
            using (var gfx = Graphics.FromImage(bmp))
            {
                var hdc = gfx.GetHdc();
                try { PrintWindow(hwnd, hdc, PW_RENDERFULLCONTENT); }
                finally { gfx.ReleaseHdc(hdc); }
            }

            // Optional crop to the requested subrect (e.g. task pane bounds)
            Bitmap toEncode = bmp;
            Bitmap? cropped = null;
            if (crop.HasValue)
            {
                var c = crop.Value;
                // Clamp crop to bitmap bounds
                int x = Math.Max(0, c.X);
                int y = Math.Max(0, c.Y);
                int w = Math.Min(bmp.Width - x, c.Width);
                int h = Math.Min(bmp.Height - y, c.Height);
                if (w > 0 && h > 0)
                {
                    cropped = bmp.Clone(new Rectangle(x, y, w, h), bmp.PixelFormat);
                    toEncode = cropped;
                }
            }

            using var ms = new MemoryStream();
            var jpegEncoder = GetJpegEncoder();
            if (jpegEncoder != null)
            {
                var encoderParams = new EncoderParameters(1);
                encoderParams.Param[0] = new EncoderParameter(System.Drawing.Imaging.Encoder.Quality, (long)Math.Clamp(quality, 1, 100));
                toEncode.Save(ms, jpegEncoder, encoderParams);
            }
            else
            {
                toEncode.Save(ms, ImageFormat.Jpeg);
            }

            cropped?.Dispose();
            return ms.ToArray();
        }
        catch
        {
            return null;
        }
    }

    private static ImageCodecInfo? _jpegEncoderCache;
    private static ImageCodecInfo? GetJpegEncoder()
    {
        if (_jpegEncoderCache != null) return _jpegEncoderCache;
        foreach (var codec in ImageCodecInfo.GetImageEncoders())
        {
            if (codec.FormatID == ImageFormat.Jpeg.Guid)
            {
                _jpegEncoderCache = codec;
                return codec;
            }
        }
        return null;
    }

    // ── Single window capture ────────────────────────────────────────

    private string CaptureSingleWindow(nint hwnd, string? path, string modeLabel)
    {
        if (!IsWindow(hwnd))
            return Response.Error($"Invalid window handle: 0x{hwnd:X}");

        if (!GetWindowRect(hwnd, out var rect))
            return Response.Error("Failed to get window rectangle");

        if (rect.Width <= 0 || rect.Height <= 0)
            return Response.Error($"Window has invalid dimensions ({rect.Width}x{rect.Height})");

        using var bmp = new Bitmap(rect.Width, rect.Height);
        using (var gfx = Graphics.FromImage(bmp))
        {
            var hdc = gfx.GetHdc();
            try { PrintWindow(hwnd, hdc, PW_RENDERFULLCONTENT); }
            finally { gfx.ReleaseHdc(hdc); }
        }

        path = EnsureOutputPath(path);
        bmp.Save(path, ImageFormat.Png);

        return Response.Ok(new
        {
            path,
            mode = modeLabel,
            hwnd = hwnd.ToInt64(),
            title = GetTitle(hwnd),
            class_name = GetClass(hwnd),
            width = rect.Width,
            height = rect.Height,
            size_bytes = new FileInfo(path).Length,
        });
    }

    // ── Composite capture (multiple windows into one image) ──────────

    private string CaptureComposite(List<WindowInfo> windows, string? path, string modeLabel)
    {
        if (windows.Count == 0)
            return Response.Error("No windows to capture");

        // Compute bounding rectangle that encompasses all windows
        int minLeft = int.MaxValue, minTop = int.MaxValue;
        int maxRight = int.MinValue, maxBottom = int.MinValue;

        var validWindows = new List<(WindowInfo Info, RECT Rect, Bitmap Bmp)>();
        foreach (var w in windows)
        {
            if (!GetWindowRect(w.Handle, out var r)) continue;
            if (r.Width <= 0 || r.Height <= 0) continue;

            var bmp = new Bitmap(r.Width, r.Height);
            using (var gfx = Graphics.FromImage(bmp))
            {
                var hdc = gfx.GetHdc();
                try { PrintWindow(w.Handle, hdc, PW_RENDERFULLCONTENT); }
                finally { gfx.ReleaseHdc(hdc); }
            }

            validWindows.Add((w, r, bmp));

            minLeft = Math.Min(minLeft, r.Left);
            minTop = Math.Min(minTop, r.Top);
            maxRight = Math.Max(maxRight, r.Right);
            maxBottom = Math.Max(maxBottom, r.Bottom);
        }

        if (validWindows.Count == 0)
            return Response.Error("None of the enumerated windows had valid dimensions");

        int compositeWidth = maxRight - minLeft;
        int compositeHeight = maxBottom - minTop;

        using var composite = new Bitmap(compositeWidth, compositeHeight);
        using (var gfx = Graphics.FromImage(composite))
        {
            gfx.Clear(Color.Transparent);
            // Draw main windows first, dialogs on top (so modals overlay correctly)
            foreach (var (_, r, b) in validWindows.OrderBy(v => v.Info.ClassName == "XLMAIN" ? 0 : 1))
            {
                gfx.DrawImage(b, r.Left - minLeft, r.Top - minTop);
                b.Dispose();
            }
        }

        path = EnsureOutputPath(path);
        composite.Save(path, ImageFormat.Png);

        var windowsJson = new JsonArray();
        foreach (var w in windows)
        {
            windowsJson.Add(new JsonObject
            {
                ["hwnd"] = w.Handle.ToInt64(),
                ["title"] = w.Title,
                ["class_name"] = w.ClassName,
            });
        }

        return Response.Ok(new
        {
            path,
            mode = modeLabel,
            width = compositeWidth,
            height = compositeHeight,
            window_count = validWindows.Count,
            windows = windowsJson,
            size_bytes = new FileInfo(path).Length,
        });
    }

    // ── Window enumeration ───────────────────────────────────────────

    private List<WindowInfo> EnumerateExcelWindows(HashSet<uint> excelPids)
    {
        var result = new List<WindowInfo>();
        EnumWindows((hWnd, _) =>
        {
            if (!IsWindowVisible(hWnd)) return true;
            GetWindowThreadProcessId(hWnd, out uint pid);
            if (!excelPids.Contains(pid)) return true;

            var title = GetTitle(hWnd);
            var className = GetClass(hWnd);

            result.Add(new WindowInfo
            {
                Handle = hWnd,
                Title = title,
                ClassName = className,
                Pid = pid,
            });
            return true;
        }, 0);
        return result;
    }

    private static string GetTitle(nint hWnd)
    {
        var sb = new StringBuilder(512);
        GetWindowText(hWnd, sb, sb.Capacity);
        return sb.ToString();
    }

    private static string GetClass(nint hWnd)
    {
        var sb = new StringBuilder(256);
        GetClassName(hWnd, sb, sb.Capacity);
        return sb.ToString();
    }

    private static string EnsureOutputPath(string? path)
    {
        if (!string.IsNullOrEmpty(path)) return path;
        var timestamp = DateTime.Now.ToString("yyyyMMdd_HHmmss_fff");
        return Path.Combine(Path.GetTempPath(), $"xrai_screenshot_{timestamp}.png");
    }

    private struct WindowInfo
    {
        public nint Handle;
        public string Title;
        public string ClassName;
        public uint Pid;
    }
}
