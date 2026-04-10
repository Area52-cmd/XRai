using System.Diagnostics;
using System.Drawing;
using System.Drawing.Imaging;
using System.Text;
using System.Text.Json.Nodes;
using XRai.Core;

namespace XRai.Vision;

/// <summary>
/// OCR operations using Windows' built-in OCR via PowerShell (Windows.Media.Ocr).
/// Falls back to Tesseract CLI if PowerShell OCR is unavailable.
///
/// Commands:
///   ocr.screen  — OCR a rectangular region of the screen
///   ocr.element — OCR the bounding rectangle of a UIA element (future extension)
/// </summary>
public class OcrOps
{
    public void Register(CommandRouter router)
    {
        router.Register("ocr.screen", HandleOcrScreen);
        router.Register("ocr.element", HandleOcrElement);
    }

    // ── ocr.screen ──────────────────────────────────────────────────

    /// <summary>
    /// Capture a screen region and run OCR on it.
    /// Args: x, y, width, height (all required — screen coordinates in pixels)
    /// Optional: language (default "en"), path (save captured image to this path)
    /// </summary>
    private string HandleOcrScreen(JsonObject args)
    {
        var x = args["x"]?.GetValue<int>();
        var y = args["y"]?.GetValue<int>();
        var width = args["width"]?.GetValue<int>();
        var height = args["height"]?.GetValue<int>();

        if (x == null || y == null || width == null || height == null)
            return Response.Error("ocr.screen requires 'x', 'y', 'width', 'height' (screen coordinates in pixels)");

        if (width <= 0 || height <= 0)
            return Response.Error("width and height must be positive");

        var language = args["language"]?.GetValue<string>() ?? "en";

        // Capture the screen region
        var imagePath = args["path"]?.GetValue<string>()
            ?? Path.Combine(Path.GetTempPath(), $"xrai_ocr_{DateTime.Now:yyyyMMdd_HHmmss_fff}.png");

        try
        {
            using var bmp = new Bitmap(width.Value, height.Value);
            using (var gfx = Graphics.FromImage(bmp))
            {
                gfx.CopyFromScreen(x.Value, y.Value, 0, 0, new Size(width.Value, height.Value));
            }
            bmp.Save(imagePath, ImageFormat.Png);
        }
        catch (Exception ex)
        {
            return Response.Error($"Screen capture failed: {ex.Message}");
        }

        return RunOcr(imagePath, language);
    }

    // ── ocr.element ─────────────────────────────────────────────────

    /// <summary>
    /// OCR the bounding rectangle of a specific window/element.
    /// Args: hwnd (window handle) or title (window title substring)
    /// Falls back to capturing the full window if no specific element is found.
    /// </summary>
    private string HandleOcrElement(JsonObject args)
    {
        var hwndArg = args["hwnd"]?.GetValue<long>();
        var title = args["title"]?.GetValue<string>();
        var language = args["language"]?.GetValue<string>() ?? "en";

        if (hwndArg == null && title == null)
            return Response.Error("ocr.element requires 'hwnd' (window handle) or 'title' (window title substring)");

        // Find the window handle
        nint hwnd = nint.Zero;
        if (hwndArg.HasValue)
        {
            hwnd = (nint)hwndArg.Value;
        }
        else if (title != null)
        {
            hwnd = FindWindowByTitle(title);
            if (hwnd == nint.Zero)
                return Response.Error($"No window found with title containing '{title}'");
        }

        if (!IsWindow(hwnd))
            return Response.Error($"Invalid window handle: 0x{hwnd:X}");

        if (!GetWindowRect(hwnd, out var rect))
            return Response.Error("Failed to get window rectangle");

        if (rect.Width <= 0 || rect.Height <= 0)
            return Response.Error($"Window has invalid dimensions ({rect.Width}x{rect.Height})");

        // Capture via PrintWindow (same as Capture.cs)
        var imagePath = Path.Combine(Path.GetTempPath(), $"xrai_ocr_element_{DateTime.Now:yyyyMMdd_HHmmss_fff}.png");
        try
        {
            using var bmp = new Bitmap(rect.Width, rect.Height);
            using (var gfx = Graphics.FromImage(bmp))
            {
                var hdc = gfx.GetHdc();
                try { PrintWindow(hwnd, hdc, 2 /* PW_RENDERFULLCONTENT */); }
                finally { gfx.ReleaseHdc(hdc); }
            }
            bmp.Save(imagePath, ImageFormat.Png);
        }
        catch (Exception ex)
        {
            return Response.Error($"Window capture failed: {ex.Message}");
        }

        return RunOcr(imagePath, language);
    }

    // ── OCR engine ──────────────────────────────────────────────────

    /// <summary>
    /// Run OCR on an image file. Tries Windows.Media.Ocr via PowerShell first,
    /// then falls back to Tesseract CLI.
    /// </summary>
    private static string RunOcr(string imagePath, string language)
    {
        // Strategy 1: PowerShell with Windows.Media.Ocr (built into Windows 10+)
        var psResult = TryPowerShellOcr(imagePath, language);
        if (psResult != null)
        {
            return Response.Ok(new
            {
                text = psResult,
                engine = "windows_media_ocr",
                image_path = imagePath,
                language,
            });
        }

        // Strategy 2: Tesseract CLI
        var tessResult = TryTesseractOcr(imagePath, language);
        if (tessResult != null)
        {
            return Response.Ok(new
            {
                text = tessResult,
                engine = "tesseract",
                image_path = imagePath,
                language,
            });
        }

        return Response.ErrorWithData(
            "OCR failed. Neither Windows.Media.Ocr (PowerShell) nor Tesseract CLI produced results.",
            new
            {
                image_path = imagePath,
                install_tesseract = "winget install UB-Mannheim.TesseractOCR",
                hint = "Ensure PowerShell 5.1+ is available, or install Tesseract and add it to PATH.",
            });
    }

    private static string? TryPowerShellOcr(string imagePath, string language)
    {
        try
        {
            // Map simple language codes to BCP-47 tags for Windows.Media.Ocr
            var bcp47 = language switch
            {
                "en" => "en-US",
                "de" => "de-DE",
                "fr" => "fr-FR",
                "es" => "es-ES",
                "ja" => "ja-JP",
                "zh" => "zh-CN",
                _ => language
            };

            // PowerShell script that uses Windows.Media.Ocr
            var script = $@"
Add-Type -AssemblyName System.Runtime.WindowsRuntime
$null = [Windows.Media.Ocr.OcrEngine, Windows.Foundation, ContentType=WindowsRuntime]
$null = [Windows.Graphics.Imaging.BitmapDecoder, Windows.Foundation, ContentType=WindowsRuntime]
$null = [Windows.Storage.StorageFile, Windows.Foundation, ContentType=WindowsRuntime]

function Await($WinRtTask, $ResultType) {{
    $asTaskGeneric = ([System.WindowsRuntimeSystemExtensions].GetMethods() | Where-Object {{ $_.Name -eq 'AsTask' -and $_.GetParameters().Count -eq 1 -and $_.GetParameters()[0].ParameterType.Name -eq 'IAsyncOperation`1' }})[0]
    $asTask = $asTaskGeneric.MakeGenericMethod($ResultType)
    $netTask = $asTask.Invoke($null, @($WinRtTask))
    $netTask.Wait(-1) | Out-Null
    $netTask.Result
}}

$file = Await ([Windows.Storage.StorageFile]::GetFileFromPathAsync('{imagePath.Replace("'", "''")}')) ([Windows.Storage.StorageFile])
$stream = Await ($file.OpenAsync([Windows.Storage.FileAccessMode]::Read)) ([Windows.Storage.Streams.IRandomAccessStream])
$decoder = Await ([Windows.Graphics.Imaging.BitmapDecoder]::CreateAsync($stream)) ([Windows.Graphics.Imaging.BitmapDecoder])
$bitmap = Await ($decoder.GetSoftwareBitmapAsync()) ([Windows.Graphics.Imaging.SoftwareBitmap])

$lang = New-Object Windows.Globalization.Language('{bcp47}')
$engine = [Windows.Media.Ocr.OcrEngine]::TryCreateFromLanguage($lang)
if ($engine -eq $null) {{
    $engine = [Windows.Media.Ocr.OcrEngine]::TryCreateFromUserProfileLanguages()
}}
if ($engine -eq $null) {{ exit 1 }}

$result = Await ($engine.RecognizeAsync($bitmap)) ([Windows.Media.Ocr.OcrResult])
Write-Output $result.Text
$stream.Dispose()
";

            var psi = new ProcessStartInfo
            {
                FileName = "powershell.exe",
                Arguments = "-NoProfile -NonInteractive -ExecutionPolicy Bypass -Command -",
                RedirectStandardInput = true,
                RedirectStandardOutput = true,
                RedirectStandardError = true,
                UseShellExecute = false,
                CreateNoWindow = true,
            };

            using var proc = Process.Start(psi);
            if (proc == null) return null;

            proc.StandardInput.Write(script);
            proc.StandardInput.Close();

            var output = proc.StandardOutput.ReadToEnd();
            proc.WaitForExit(15000);

            if (proc.ExitCode == 0 && !string.IsNullOrWhiteSpace(output))
                return output.Trim();
        }
        catch { }

        return null;
    }

    private static string? TryTesseractOcr(string imagePath, string language)
    {
        try
        {
            var tessLang = language switch
            {
                "en" => "eng",
                "de" => "deu",
                "fr" => "fra",
                "es" => "spa",
                "ja" => "jpn",
                "zh" => "chi_sim",
                _ => language
            };

            var psi = new ProcessStartInfo
            {
                FileName = "tesseract",
                Arguments = $"\"{imagePath}\" stdout -l {tessLang}",
                RedirectStandardOutput = true,
                RedirectStandardError = true,
                UseShellExecute = false,
                CreateNoWindow = true,
            };

            using var proc = Process.Start(psi);
            if (proc == null) return null;

            var output = proc.StandardOutput.ReadToEnd();
            proc.WaitForExit(15000);

            if (proc.ExitCode == 0 && !string.IsNullOrWhiteSpace(output))
                return output.Trim();
        }
        catch { }

        return null;
    }

    // ── Win32 P/Invoke ──────────────────────────────────────────────

    [System.Runtime.InteropServices.DllImport("user32.dll")]
    private static extern bool PrintWindow(nint hwnd, nint hdcBlt, uint nFlags);

    [System.Runtime.InteropServices.DllImport("user32.dll")]
    private static extern bool GetWindowRect(nint hwnd, out RECT lpRect);

    [System.Runtime.InteropServices.DllImport("user32.dll")]
    private static extern bool IsWindow(nint hWnd);

    [System.Runtime.InteropServices.DllImport("user32.dll")]
    private static extern bool IsWindowVisible(nint hWnd);

    [System.Runtime.InteropServices.DllImport("user32.dll", CharSet = System.Runtime.InteropServices.CharSet.Unicode)]
    private static extern int GetWindowText(nint hWnd, StringBuilder lpString, int nMaxCount);

    private delegate bool EnumWindowsProc(nint hWnd, nint lParam);

    [System.Runtime.InteropServices.DllImport("user32.dll")]
    private static extern bool EnumWindows(EnumWindowsProc lpEnumFunc, nint lParam);

    [System.Runtime.InteropServices.StructLayout(System.Runtime.InteropServices.LayoutKind.Sequential)]
    private struct RECT
    {
        public int Left, Top, Right, Bottom;
        public int Width => Right - Left;
        public int Height => Bottom - Top;
    }

    private static nint FindWindowByTitle(string titleSubstring)
    {
        nint found = nint.Zero;
        EnumWindows((hWnd, _) =>
        {
            if (!IsWindowVisible(hWnd)) return true;
            var sb = new StringBuilder(512);
            GetWindowText(hWnd, sb, sb.Capacity);
            var title = sb.ToString();
            if (title.Contains(titleSubstring, StringComparison.OrdinalIgnoreCase))
            {
                found = hWnd;
                return false; // stop enumeration
            }
            return true;
        }, 0);
        return found;
    }
}
