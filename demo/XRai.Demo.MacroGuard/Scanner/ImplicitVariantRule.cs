using System.Text.RegularExpressions;
using XRai.Demo.MacroGuard.Models;

namespace XRai.Demo.MacroGuard.Scanner;

public class ImplicitVariantRule : IScannerRule
{
    public string Name => "ImplicitVariant";
    public bool RequiresStrictMode => true;

    private static readonly Regex DimWithoutAs = new(
        @"^\s*Dim\s+(\w+)\s*$",
        RegexOptions.IgnoreCase | RegexOptions.Compiled);

    private static readonly Regex DimWithoutAsInline = new(
        @"^\s*Dim\s+(\w+)\s*[,\n]",
        RegexOptions.IgnoreCase | RegexOptions.Compiled);

    public List<VbaIssue> Check(VbaModuleInfo module)
    {
        var issues = new List<VbaIssue>();
        var lines = module.Code.Split('\n');

        for (int i = 0; i < lines.Length; i++)
        {
            var trimmed = lines[i].TrimStart();
            if (trimmed.StartsWith("'")) continue;

            if (Regex.IsMatch(trimmed, @"^Dim\s+\w+\s*$", RegexOptions.IgnoreCase) ||
                (Regex.IsMatch(trimmed, @"^Dim\s+", RegexOptions.IgnoreCase) &&
                 !trimmed.Contains(" As ", StringComparison.OrdinalIgnoreCase)))
            {
                var varMatch = Regex.Match(trimmed, @"^Dim\s+(\w+)", RegexOptions.IgnoreCase);
                if (varMatch.Success)
                {
                    issues.Add(new VbaIssue
                    {
                        Severity = "Info",
                        ModuleName = module.Name,
                        LineNumber = i + 1,
                        RuleName = Name,
                        Description = $"Variable '{varMatch.Groups[1].Value}' declared without explicit type (implicit Variant)",
                        SuggestedFix = $"Add explicit type: Dim {varMatch.Groups[1].Value} As <Type>"
                    });
                }
            }
        }
        return issues;
    }
}
