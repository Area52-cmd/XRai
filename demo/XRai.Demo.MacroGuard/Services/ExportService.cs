using System.Text;
using System.Web;
using XRai.Demo.MacroGuard.Models;

namespace XRai.Demo.MacroGuard.Services;

public static class ExportService
{
    /// <summary>
    /// Export a single module's VBA code as a .bas file content string.
    /// </summary>
    public static string ExportAsBas(VbaModuleInfo module)
    {
        var sb = new StringBuilder();
        sb.AppendLine($"Attribute VB_Name = \"{module.Name}\"");
        sb.AppendLine(module.Code);
        return sb.ToString();
    }

    /// <summary>
    /// Export all modules as a single concatenated text file with module headers.
    /// </summary>
    public static string ExportAllAsText(IEnumerable<VbaModuleInfo> modules)
    {
        var sb = new StringBuilder();
        foreach (var mod in modules)
        {
            sb.AppendLine(new string('=', 60));
            sb.AppendLine($"' Module: {mod.Name}  ({mod.Type}, {mod.LineCount} lines)");
            sb.AppendLine(new string('=', 60));
            sb.AppendLine(mod.Code);
            sb.AppendLine();
        }
        return sb.ToString();
    }

    /// <summary>
    /// Export VBA documentation as HTML with basic syntax highlighting.
    /// </summary>
    public static string ExportDocumentationHtml(IEnumerable<VbaModuleInfo> modules)
    {
        var sb = new StringBuilder();
        sb.AppendLine("<!DOCTYPE html><html><head><meta charset='utf-8'/>");
        sb.AppendLine("<title>VBA Documentation</title>");
        sb.AppendLine("<style>");
        sb.AppendLine("body { font-family: 'Segoe UI', sans-serif; background: #1e1e1e; color: #d4d4d4; padding: 20px; }");
        sb.AppendLine("h1 { color: #569cd6; } h2 { color: #4ec9b0; border-bottom: 1px solid #333; padding-bottom: 4px; }");
        sb.AppendLine("pre { background: #252526; padding: 12px; border-radius: 4px; overflow-x: auto; }");
        sb.AppendLine(".kw { color: #569cd6; } .cm { color: #6a9955; } .str { color: #ce9178; } .fn { color: #dcdcaa; }");
        sb.AppendLine("</style></head><body>");
        sb.AppendLine("<h1>VBA Documentation</h1>");
        sb.AppendLine($"<p>Generated: {DateTime.Now:yyyy-MM-dd HH:mm:ss}</p>");

        foreach (var mod in modules)
        {
            sb.AppendLine($"<h2>{Encode(mod.Name)} <small>({Encode(mod.Type)}, {mod.LineCount} lines)</small></h2>");
            sb.AppendLine("<pre>");
            foreach (var line in mod.Code.Split('\n'))
            {
                sb.AppendLine(HighlightLine(line));
            }
            sb.AppendLine("</pre>");
        }

        sb.AppendLine("</body></html>");
        return sb.ToString();
    }

    /// <summary>
    /// Export health check report as CSV.
    /// </summary>
    public static string ExportHealthCsv(IEnumerable<VbaIssue> issues)
    {
        var sb = new StringBuilder();
        sb.AppendLine("Severity,Module,Line,Rule,Description,SuggestedFix");
        foreach (var issue in issues)
        {
            sb.AppendLine($"{Csv(issue.Severity)},{Csv(issue.ModuleName)},{issue.LineNumber},{Csv(issue.RuleName)},{Csv(issue.Description)},{Csv(issue.SuggestedFix)}");
        }
        return sb.ToString();
    }

    private static string Encode(string s) => HttpUtility.HtmlEncode(s);

    private static string Csv(string s) => s.Contains(',') || s.Contains('"') ? $"\"{s.Replace("\"", "\"\"")}\"" : s;

    private static string HighlightLine(string line)
    {
        var encoded = Encode(line);
        // Comments
        if (encoded.TrimStart().StartsWith("&#39;") || encoded.TrimStart().StartsWith("'"))
            return $"<span class='cm'>{encoded}</span>";
        // Keywords
        var keywords = new[] { "Sub", "Function", "End Sub", "End Function", "If", "Then", "Else", "ElseIf",
            "End If", "For", "Next", "Do", "Loop", "While", "Wend", "Select", "Case", "Dim", "Set",
            "Public", "Private", "Const", "As", "String", "Long", "Integer", "Boolean", "Variant",
            "Object", "Double", "On Error", "GoTo", "Resume", "Exit", "With", "End With", "Each", "In",
            "Option Explicit", "Call", "Nothing", "New", "True", "False" };
        foreach (var kw in keywords.OrderByDescending(k => k.Length))
        {
            encoded = System.Text.RegularExpressions.Regex.Replace(
                encoded, $@"\b{System.Text.RegularExpressions.Regex.Escape(kw)}\b",
                $"<span class='kw'>{kw}</span>");
        }
        return encoded;
    }
}
