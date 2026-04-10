using System.Diagnostics;
using System.Text;
using System.Text.Json;
using System.Text.Json.Nodes;
using System.Xml.Linq;

namespace XRai.Core;

/// <summary>
/// Lightweight test reporting system. Records structured pass/fail steps
/// with optional screenshots, and generates HTML or JUnit XML reports.
///
/// State is stored in a static singleton so it persists across commands
/// within a session.
///
/// Commands:
///   test.start  — begin a named test session
///   test.step   — record a test step (pass/fail/skip)
///   test.assert — assert a condition, auto-record as step
///   test.end    — end the session, return summary
///   test.report — generate HTML or JUnit XML report
/// </summary>
public class TestReporter
{
    // ── Static singleton state ───────────────────────────────────────

    private static TestSession? _currentSession;
    private static readonly object Lock = new();

    public static TestSession? CurrentSession
    {
        get { lock (Lock) return _currentSession; }
    }

    public class TestSession
    {
        public string Name { get; set; } = "";
        public DateTime StartedAt { get; set; }
        public DateTime? EndedAt { get; set; }
        public List<TestStep> Steps { get; set; } = new();
        public bool IsEnded => EndedAt.HasValue;
    }

    public class TestStep
    {
        public string Name { get; set; } = "";
        public string Status { get; set; } = "pass"; // pass, fail, skip
        public string? Message { get; set; }
        public string? ScreenshotPath { get; set; }
        public DateTime Timestamp { get; set; }
        public double DurationMs { get; set; }
    }

    // ── Registration ────────────────────────────────────────────────

    public void Register(CommandRouter router)
    {
        router.Register("test.start", HandleStart);
        router.Register("test.step", HandleStep);
        router.Register("test.end", HandleEnd);
        router.Register("test.report", HandleReport);
    }

    // ── test.start ──────────────────────────────────────────────────

    private string HandleStart(JsonObject args)
    {
        var name = args["name"]?.GetValue<string>();
        if (string.IsNullOrWhiteSpace(name))
            return Response.Error("test.start requires 'name'");

        lock (Lock)
        {
            _currentSession = new TestSession
            {
                Name = name,
                StartedAt = DateTime.UtcNow,
            };
        }

        return Response.Ok(new
        {
            session = name,
            started_at = _currentSession.StartedAt.ToString("o"),
        });
    }

    // ── test.step ───────────────────────────────────────────────────

    private string HandleStep(JsonObject args)
    {
        var session = CurrentSession;
        if (session == null || session.IsEnded)
            return Response.Error("No active test session. Call test.start first.");

        var name = args["name"]?.GetValue<string>() ?? $"Step {session.Steps.Count + 1}";
        var status = args["status"]?.GetValue<string>() ?? "pass";
        var message = args["message"]?.GetValue<string>();
        var screenshotPath = args["screenshot"]?.GetValue<string>();

        if (status != "pass" && status != "fail" && status != "skip")
            return Response.Error("test.step 'status' must be 'pass', 'fail', or 'skip'");

        var step = new TestStep
        {
            Name = name,
            Status = status,
            Message = message,
            ScreenshotPath = screenshotPath,
            Timestamp = DateTime.UtcNow,
            DurationMs = 0,
        };

        lock (Lock)
        {
            session.Steps.Add(step);
        }

        return Response.Ok(new
        {
            step = name,
            status,
            step_number = session.Steps.Count,
        });
    }

    // ── RecordAssertResult (called from AssertOps) ──────────────────

    /// <summary>
    /// Record an assertion result as a test step. Called from AssertOps
    /// when a test session is active.
    /// </summary>
    public static void RecordAssertResult(string stepName, bool passed, string? message = null)
    {
        var session = CurrentSession;
        if (session == null || session.IsEnded) return;

        var step = new TestStep
        {
            Name = stepName,
            Status = passed ? "pass" : "fail",
            Message = message,
            Timestamp = DateTime.UtcNow,
        };

        lock (Lock)
        {
            session.Steps.Add(step);
        }
    }

    // ── test.end ────────────────────────────────────────────────────

    private string HandleEnd(JsonObject args)
    {
        var session = CurrentSession;
        if (session == null)
            return Response.Error("No active test session. Call test.start first.");

        lock (Lock)
        {
            session.EndedAt = DateTime.UtcNow;
        }

        var total = session.Steps.Count;
        var passed = session.Steps.Count(s => s.Status == "pass");
        var failed = session.Steps.Count(s => s.Status == "fail");
        var skipped = session.Steps.Count(s => s.Status == "skip");
        var duration = (session.EndedAt!.Value - session.StartedAt).TotalSeconds;

        return Response.Ok(new
        {
            session = session.Name,
            total,
            passed,
            failed,
            skipped,
            duration_seconds = Math.Round(duration, 2),
            result = failed > 0 ? "FAIL" : "PASS",
        });
    }

    // ── test.report ─────────────────────────────────────────────────

    /// <summary>
    /// Generate a report file.
    /// Args: path (required), format ("html" or "junit", default "html")
    /// </summary>
    private string HandleReport(JsonObject args)
    {
        var session = CurrentSession;
        if (session == null)
            return Response.Error("No test session. Call test.start first.");

        var path = args["path"]?.GetValue<string>();
        if (string.IsNullOrWhiteSpace(path))
            return Response.Error("test.report requires 'path'");

        var format = args["format"]?.GetValue<string>() ?? "html";

        try
        {
            if (format.Equals("junit", StringComparison.OrdinalIgnoreCase) ||
                format.Equals("xml", StringComparison.OrdinalIgnoreCase))
            {
                GenerateJUnitXml(session, path);
            }
            else
            {
                GenerateHtmlReport(session, path);
            }

            return Response.Ok(new
            {
                report_path = path,
                format,
                steps = session.Steps.Count,
                size_bytes = new FileInfo(path).Length,
            });
        }
        catch (Exception ex)
        {
            return Response.Error($"Failed to generate report: {ex.Message}");
        }
    }

    // ── Report generators ───────────────────────────────────────────

    private static void GenerateJUnitXml(TestSession session, string path)
    {
        var duration = session.EndedAt.HasValue
            ? (session.EndedAt.Value - session.StartedAt).TotalSeconds
            : (DateTime.UtcNow - session.StartedAt).TotalSeconds;

        var failures = session.Steps.Count(s => s.Status == "fail");

        var testsuite = new XElement("testsuite",
            new XAttribute("name", session.Name),
            new XAttribute("tests", session.Steps.Count),
            new XAttribute("failures", failures),
            new XAttribute("skipped", session.Steps.Count(s => s.Status == "skip")),
            new XAttribute("time", Math.Round(duration, 2)),
            new XAttribute("timestamp", session.StartedAt.ToString("o")));

        foreach (var step in session.Steps)
        {
            var testcase = new XElement("testcase",
                new XAttribute("name", step.Name),
                new XAttribute("time", Math.Round(step.DurationMs / 1000.0, 3)));

            if (step.Status == "fail")
            {
                testcase.Add(new XElement("failure",
                    new XAttribute("message", step.Message ?? "Assertion failed")));
            }
            else if (step.Status == "skip")
            {
                testcase.Add(new XElement("skipped",
                    new XAttribute("message", step.Message ?? "Skipped")));
            }

            testsuite.Add(testcase);
        }

        var doc = new XDocument(
            new XDeclaration("1.0", "UTF-8", null),
            new XElement("testsuites", testsuite));

        // Ensure directory exists
        var dir = Path.GetDirectoryName(path);
        if (!string.IsNullOrEmpty(dir)) Directory.CreateDirectory(dir);

        doc.Save(path);
    }

    private static void GenerateHtmlReport(TestSession session, string path)
    {
        var duration = session.EndedAt.HasValue
            ? (session.EndedAt.Value - session.StartedAt).TotalSeconds
            : (DateTime.UtcNow - session.StartedAt).TotalSeconds;

        var total = session.Steps.Count;
        var passed = session.Steps.Count(s => s.Status == "pass");
        var failed = session.Steps.Count(s => s.Status == "fail");
        var skipped = session.Steps.Count(s => s.Status == "skip");
        var overallResult = failed > 0 ? "FAIL" : "PASS";
        var resultColor = failed > 0 ? "#e74c3c" : "#27ae60";

        var sb = new StringBuilder();
        sb.AppendLine("<!DOCTYPE html>");
        sb.AppendLine("<html lang=\"en\"><head><meta charset=\"UTF-8\">");
        sb.AppendLine($"<title>XRai Test Report — {Esc(session.Name)}</title>");
        sb.AppendLine("<style>");
        sb.AppendLine("body { font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, sans-serif; margin: 40px; background: #f5f5f5; }");
        sb.AppendLine(".container { max-width: 900px; margin: 0 auto; background: white; border-radius: 8px; box-shadow: 0 2px 8px rgba(0,0,0,0.1); padding: 32px; }");
        sb.AppendLine("h1 { margin: 0 0 8px 0; color: #2c3e50; }");
        sb.AppendLine(".meta { color: #7f8c8d; margin-bottom: 24px; }");
        sb.AppendLine(".summary { display: flex; gap: 16px; margin-bottom: 24px; }");
        sb.AppendLine(".stat { padding: 12px 20px; border-radius: 6px; text-align: center; min-width: 80px; }");
        sb.AppendLine(".stat .num { font-size: 28px; font-weight: bold; }");
        sb.AppendLine(".stat .label { font-size: 12px; text-transform: uppercase; color: #666; }");
        sb.AppendLine(".pass-bg { background: #d5f5e3; } .fail-bg { background: #fadbd8; } .skip-bg { background: #fdebd0; } .total-bg { background: #d6eaf8; }");
        sb.AppendLine("table { width: 100%; border-collapse: collapse; }");
        sb.AppendLine("th { text-align: left; padding: 10px; border-bottom: 2px solid #ddd; color: #555; }");
        sb.AppendLine("td { padding: 10px; border-bottom: 1px solid #eee; }");
        sb.AppendLine(".badge { display: inline-block; padding: 2px 10px; border-radius: 12px; font-size: 12px; font-weight: bold; color: white; }");
        sb.AppendLine(".badge-pass { background: #27ae60; } .badge-fail { background: #e74c3c; } .badge-skip { background: #f39c12; }");
        sb.AppendLine(".screenshot { max-width: 300px; border: 1px solid #ddd; border-radius: 4px; margin-top: 8px; }");
        sb.AppendLine("</style></head><body><div class=\"container\">");

        sb.AppendLine($"<h1>{Esc(session.Name)}</h1>");
        sb.AppendLine($"<div class=\"meta\">{session.StartedAt:yyyy-MM-dd HH:mm:ss} UTC &middot; Duration: {duration:F1}s &middot; ");
        sb.AppendLine($"Overall: <strong style=\"color:{resultColor}\">{overallResult}</strong></div>");

        sb.AppendLine("<div class=\"summary\">");
        sb.AppendLine($"<div class=\"stat total-bg\"><div class=\"num\">{total}</div><div class=\"label\">Total</div></div>");
        sb.AppendLine($"<div class=\"stat pass-bg\"><div class=\"num\">{passed}</div><div class=\"label\">Passed</div></div>");
        sb.AppendLine($"<div class=\"stat fail-bg\"><div class=\"num\">{failed}</div><div class=\"label\">Failed</div></div>");
        sb.AppendLine($"<div class=\"stat skip-bg\"><div class=\"num\">{skipped}</div><div class=\"label\">Skipped</div></div>");
        sb.AppendLine("</div>");

        sb.AppendLine("<table><thead><tr><th>#</th><th>Step</th><th>Status</th><th>Message</th></tr></thead><tbody>");

        for (int i = 0; i < session.Steps.Count; i++)
        {
            var step = session.Steps[i];
            var badgeClass = step.Status switch
            {
                "pass" => "badge-pass",
                "fail" => "badge-fail",
                _ => "badge-skip"
            };

            sb.AppendLine("<tr>");
            sb.AppendLine($"<td>{i + 1}</td>");
            sb.AppendLine($"<td>{Esc(step.Name)}");

            // Embed screenshot as base64 if available
            if (!string.IsNullOrEmpty(step.ScreenshotPath) && File.Exists(step.ScreenshotPath))
            {
                try
                {
                    var bytes = File.ReadAllBytes(step.ScreenshotPath);
                    var b64 = Convert.ToBase64String(bytes);
                    sb.AppendLine($"<br/><img class=\"screenshot\" src=\"data:image/png;base64,{b64}\" alt=\"screenshot\"/>");
                }
                catch { }
            }

            sb.AppendLine("</td>");
            sb.AppendLine($"<td><span class=\"badge {badgeClass}\">{step.Status.ToUpperInvariant()}</span></td>");
            sb.AppendLine($"<td>{Esc(step.Message ?? "")}</td>");
            sb.AppendLine("</tr>");
        }

        sb.AppendLine("</tbody></table>");
        sb.AppendLine("</div></body></html>");

        var dir = Path.GetDirectoryName(path);
        if (!string.IsNullOrEmpty(dir)) Directory.CreateDirectory(dir);

        File.WriteAllText(path, sb.ToString(), Encoding.UTF8);
    }

    private static string Esc(string text)
    {
        return text
            .Replace("&", "&amp;")
            .Replace("<", "&lt;")
            .Replace(">", "&gt;")
            .Replace("\"", "&quot;");
    }
}
