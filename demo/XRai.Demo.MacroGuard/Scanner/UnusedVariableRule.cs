using System.Text.RegularExpressions;
using XRai.Demo.MacroGuard.Models;

namespace XRai.Demo.MacroGuard.Scanner;

public class UnusedVariableRule : IScannerRule
{
    public string Name => "UnusedVariable";
    public bool RequiresStrictMode => false;

    private static readonly Regex DimPattern = new(
        @"^\s*Dim\s+(\w+)",
        RegexOptions.IgnoreCase | RegexOptions.Compiled);

    public List<VbaIssue> Check(VbaModuleInfo module)
    {
        var issues = new List<VbaIssue>();
        var lines = module.Code.Split('\n');

        for (int i = 0; i < lines.Length; i++)
        {
            var match = DimPattern.Match(lines[i]);
            if (!match.Success) continue;

            var varName = match.Groups[1].Value;

            // Check if variable is used anywhere else in the module
            bool used = false;
            for (int j = 0; j < lines.Length; j++)
            {
                if (j == i) continue;
                // Check for usage (not just in Dim statements)
                if (DimPattern.IsMatch(lines[j]) && DimPattern.Match(lines[j]).Groups[1].Value == varName)
                    continue;
                if (Regex.IsMatch(lines[j], @"\b" + Regex.Escape(varName) + @"\b", RegexOptions.IgnoreCase))
                {
                    used = true;
                    break;
                }
            }

            if (!used)
            {
                issues.Add(new VbaIssue
                {
                    Severity = "Info",
                    ModuleName = module.Name,
                    LineNumber = i + 1,
                    RuleName = Name,
                    Description = $"Variable '{varName}' is declared but never used",
                    SuggestedFix = $"Remove unused variable declaration: Dim {varName}"
                });
            }
        }
        return issues;
    }
}
