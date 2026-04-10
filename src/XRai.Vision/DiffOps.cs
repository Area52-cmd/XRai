using System.Drawing;
using System.Drawing.Imaging;
using System.Text.Json.Nodes;
using XRai.Core;

namespace XRai.Vision;

/// <summary>
/// Screenshot diffing: save named baselines, compare current state against them.
/// Baselines stored in %TEMP%/xrai_baselines/{name}.png.
/// Diff images highlight changed pixels in red and are saved to temp files.
/// </summary>
public class DiffOps
{
    private readonly Capture _capture;

    private static readonly string BaselineDir =
        Path.Combine(Path.GetTempPath(), "xrai_baselines");

    public DiffOps(Capture capture)
    {
        _capture = capture;
    }

    public void Register(CommandRouter router)
    {
        router.Register("screenshot.baseline", HandleBaseline);
        router.Register("screenshot.compare", HandleCompare);
    }

    // ── screenshot.baseline ─────────────────────────────────────────

    /// <summary>
    /// Take a screenshot and save it as a named baseline.
    /// Args: name (required), mode (optional, default "main_plus_modal")
    /// </summary>
    private string HandleBaseline(JsonObject args)
    {
        var name = args["name"]?.GetValue<string>();
        if (string.IsNullOrWhiteSpace(name))
            return Response.Error("screenshot.baseline requires 'name'");

        // Sanitize the name for filesystem safety
        var safeName = SanitizeName(name);

        Directory.CreateDirectory(BaselineDir);
        var baselinePath = Path.Combine(BaselineDir, $"{safeName}.png");

        // Take a screenshot to a temp path, then move it to the baseline location
        var tempPath = Path.Combine(Path.GetTempPath(), $"xrai_baseline_temp_{Guid.NewGuid():N}.png");
        args["path"] = tempPath;

        // Dispatch screenshot through the Capture instance by directly
        // invoking the router — but we already have the screenshot logic.
        // Instead, take the screenshot by capturing the same way Capture does.
        // We'll use a simpler approach: call the screenshot command through
        // a temporary router registration won't work. Let's just capture directly.
        try
        {
            var screenshotResult = CaptureScreenshot(args);
            if (screenshotResult == null)
                return Response.Error("Screenshot capture failed");

            // Move the captured file to the baseline location
            if (File.Exists(baselinePath))
                File.Delete(baselinePath);
            File.Move(tempPath, baselinePath);

            return Response.Ok(new
            {
                baseline = name,
                path = baselinePath,
                width = screenshotResult.Value.Width,
                height = screenshotResult.Value.Height,
                size_bytes = new FileInfo(baselinePath).Length,
            });
        }
        catch (Exception ex)
        {
            // Clean up temp file on failure
            try { if (File.Exists(tempPath)) File.Delete(tempPath); } catch { }
            return Response.Error($"Failed to capture baseline: {ex.Message}");
        }
    }

    // ── screenshot.compare ──────────────────────────────────────────

    /// <summary>
    /// Take a current screenshot and compare it against a named baseline.
    /// Returns match_percentage, diff_pixel_count, and diff_image_path.
    /// Args: name (required), threshold (optional, default 30)
    /// </summary>
    private string HandleCompare(JsonObject args)
    {
        var name = args["name"]?.GetValue<string>();
        if (string.IsNullOrWhiteSpace(name))
            return Response.Error("screenshot.compare requires 'name'");

        var threshold = args["threshold"]?.GetValue<int>() ?? 30;
        var safeName = SanitizeName(name);
        var baselinePath = Path.Combine(BaselineDir, $"{safeName}.png");

        if (!File.Exists(baselinePath))
            return Response.Error($"Baseline '{name}' not found at {baselinePath}. Use screenshot.baseline first.");

        // Take a current screenshot
        var currentPath = Path.Combine(Path.GetTempPath(), $"xrai_compare_{Guid.NewGuid():N}.png");
        args["path"] = currentPath;

        try
        {
            var screenshotResult = CaptureScreenshot(args);
            if (screenshotResult == null)
                return Response.Error("Screenshot capture failed");

            // Load both images
            using var baselineBmp = new Bitmap(baselinePath);
            using var currentBmp = new Bitmap(currentPath);

            // Compare dimensions
            if (baselineBmp.Width != currentBmp.Width || baselineBmp.Height != currentBmp.Height)
            {
                return Response.Ok(new
                {
                    match_percentage = 0.0,
                    diff_pixel_count = -1,
                    size_mismatch = true,
                    baseline_size = $"{baselineBmp.Width}x{baselineBmp.Height}",
                    current_size = $"{currentBmp.Width}x{currentBmp.Height}",
                    baseline_path = baselinePath,
                    current_path = currentPath,
                });
            }

            // Pixel-by-pixel comparison with diff image generation
            int totalPixels = baselineBmp.Width * baselineBmp.Height;
            int diffCount = 0;

            using var diffBmp = new Bitmap(baselineBmp.Width, baselineBmp.Height);

            for (int y = 0; y < baselineBmp.Height; y++)
            {
                for (int x = 0; x < baselineBmp.Width; x++)
                {
                    var a = baselineBmp.GetPixel(x, y);
                    var b = currentBmp.GetPixel(x, y);

                    if (PixelsMatch(a, b, threshold))
                    {
                        // Dim the matching pixel (semi-transparent original)
                        diffBmp.SetPixel(x, y, Color.FromArgb(128, a.R, a.G, a.B));
                    }
                    else
                    {
                        // Red overlay on changed pixels
                        diffBmp.SetPixel(x, y, Color.FromArgb(255, 255, 0, 0));
                        diffCount++;
                    }
                }
            }

            // Save diff image
            var diffPath = Path.Combine(Path.GetTempPath(),
                $"xrai_diff_{safeName}_{DateTime.Now:yyyyMMdd_HHmmss}.png");
            diffBmp.Save(diffPath, ImageFormat.Png);

            double matchPct = totalPixels > 0
                ? Math.Round((totalPixels - diffCount) * 100.0 / totalPixels, 2)
                : 100.0;

            return Response.Ok(new
            {
                match_percentage = matchPct,
                diff_pixel_count = diffCount,
                total_pixels = totalPixels,
                diff_image_path = diffPath,
                baseline_path = baselinePath,
                current_path = currentPath,
                threshold,
                width = baselineBmp.Width,
                height = baselineBmp.Height,
            });
        }
        catch (Exception ex)
        {
            return Response.Error($"Comparison failed: {ex.Message}");
        }
    }

    // ── Helpers ──────────────────────────────────────────────────────

    /// <summary>
    /// Pixel comparison with threshold to ignore anti-aliasing differences.
    /// Uses sum of absolute RGB channel differences.
    /// </summary>
    private static bool PixelsMatch(Color a, Color b, int threshold = 30)
    {
        return Math.Abs(a.R - b.R) + Math.Abs(a.G - b.G) + Math.Abs(a.B - b.B) < threshold;
    }

    private static string SanitizeName(string name)
    {
        var invalid = Path.GetInvalidFileNameChars();
        var safe = new char[name.Length];
        for (int i = 0; i < name.Length; i++)
            safe[i] = Array.IndexOf(invalid, name[i]) >= 0 ? '_' : name[i];
        return new string(safe);
    }

    /// <summary>
    /// Capture a screenshot using the same mechanism as the Capture class.
    /// Returns image dimensions, or null on failure. The image is saved to
    /// the path specified in args["path"].
    /// </summary>
    private (int Width, int Height)? CaptureScreenshot(JsonObject args)
    {
        // We use Screen.PrimaryScreen bounds as a simple capture approach.
        // The Capture class uses PrintWindow for Excel-specific capture.
        // For diff purposes we use the same approach — capture via the
        // Capture class by routing through a temporary router.
        var tempRouter = new CommandRouter(new EventStream(TextWriter.Null));
        _capture.Register(tempRouter);
        var result = tempRouter.Dispatch(new JsonObject
        {
            ["cmd"] = "screenshot",
            ["path"] = args["path"]?.GetValue<string>(),
            ["mode"] = args["mode"]?.GetValue<string>() ?? "main_plus_modal",
        }.ToJsonString());

        // Parse the result to extract dimensions
        try
        {
            var node = System.Text.Json.JsonDocument.Parse(result);
            if (node.RootElement.TryGetProperty("ok", out var ok) && ok.GetBoolean())
            {
                var w = node.RootElement.GetProperty("width").GetInt32();
                var h = node.RootElement.GetProperty("height").GetInt32();
                return (w, h);
            }
        }
        catch { }

        return null;
    }
}
