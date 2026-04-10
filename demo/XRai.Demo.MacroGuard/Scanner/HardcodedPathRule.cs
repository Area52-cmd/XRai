using System.Text.RegularExpressions;
using XRai.Demo.MacroGuard.Models;

namespace XRai.Demo.MacroGuard.Scanner;

public class HardcodedPathRule : IScannerRule
{
    public string Name => "HardcodedPath";
    public bool RequiresStrictMode => false;

    private static readonly Regex PathPattern = new(
        @"""[^""]*(?:[A-Za-z]:\\|\\\\[A-Za-z])[^""]*""",
        RegexOptions.Compiled);

    public List<VbaIssue> Check(VbaModuleInfo module)
    {
        var issues = new List<VbaIssue>();
        var lines = module.Code.Split('\n');

        for (int i = 0; i < lines.Length; i++)
        {
            var trimmed = lines[i].TrimStart();
            if (trimmed.StartsWith("'")) continue; // skip comments

            if (PathPattern.IsMatch(lines[i]))
            {
                issues.Add(new VbaIssue
                {
                    Severity = "Warning",
                    ModuleName = module.Name,
                    LineNumber = i + 1,
                    RuleName = Name,
                    Description = "Hardcoded file path detected — use variables or configuration instead",
                    SuggestedFix = "Store file paths in a named range, worksheet cell, or configuration module"
                });
            }
        }
        return issues;
    }
}
