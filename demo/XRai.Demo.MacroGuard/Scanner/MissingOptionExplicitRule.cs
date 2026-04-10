using XRai.Demo.MacroGuard.Models;

namespace XRai.Demo.MacroGuard.Scanner;

public class MissingOptionExplicitRule : IScannerRule
{
    public string Name => "MissingOptionExplicit";
    public bool RequiresStrictMode => false;

    public List<VbaIssue> Check(VbaModuleInfo module)
    {
        var issues = new List<VbaIssue>();
        if (!module.Code.TrimStart().StartsWith("Option Explicit", StringComparison.OrdinalIgnoreCase))
        {
            issues.Add(new VbaIssue
            {
                Severity = "Warning",
                ModuleName = module.Name,
                LineNumber = 1,
                RuleName = Name,
                Description = "Module lacks 'Option Explicit' declaration",
                SuggestedFix = "Add 'Option Explicit' at the top of the module"
            });
        }
        return issues;
    }
}
