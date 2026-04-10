using System.Text.RegularExpressions;
using XRai.Demo.MacroGuard.Models;

namespace XRai.Demo.MacroGuard.Scanner;

public class GoToUsageRule : IScannerRule
{
    public string Name => "GoToUsage";
    public bool RequiresStrictMode => false;

    private static readonly Regex GoToPattern = new(
        @"^\s*GoTo\s+(\w+)",
        RegexOptions.IgnoreCase | RegexOptions.Compiled);

    public List<VbaIssue> Check(VbaModuleInfo module)
    {
        var issues = new List<VbaIssue>();
        var lines = module.Code.Split('\n');

        for (int i = 0; i < lines.Length; i++)
        {
            var trimmed = lines[i].TrimStart();
            if (trimmed.StartsWith("'")) continue; // skip comments

            var match = GoToPattern.Match(lines[i]);
            if (match.Success)
            {
                var label = match.Groups[1].Value;
                // Exclude error handler GoTo patterns
                if (trimmed.StartsWith("On Error GoTo", StringComparison.OrdinalIgnoreCase))
                    continue;
                if (label.Equals("0", StringComparison.OrdinalIgnoreCase))
                    continue; // On Error GoTo 0

                issues.Add(new VbaIssue
                {
                    Severity = "Warning",
                    ModuleName = module.Name,
                    LineNumber = i + 1,
                    RuleName = Name,
                    Description = $"'GoTo {label}' — use structured control flow instead",
                    SuggestedFix = "Replace GoTo with If/Then/Else, Do/Loop, or Select Case"
                });
            }
        }
        return issues;
    }
}
