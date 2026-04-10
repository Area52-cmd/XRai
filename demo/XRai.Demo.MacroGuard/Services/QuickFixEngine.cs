using System.Text.RegularExpressions;
using XRai.Demo.MacroGuard.Models;

namespace XRai.Demo.MacroGuard.Services;

public static class QuickFixEngine
{
    public static List<QuickFix> GenerateFixes(IEnumerable<VbaIssue> issues, IEnumerable<VbaModuleInfo> modules)
    {
        var fixes = new List<QuickFix>();
        var moduleLookup = modules.ToDictionary(m => m.Name, StringComparer.OrdinalIgnoreCase);

        foreach (var issue in issues)
        {
            switch (issue.RuleName)
            {
                case "MissingOptionExplicit":
                    fixes.Add(new QuickFix
                    {
                        Name = "Add Option Explicit",
                        Description = "Prepend 'Option Explicit' to the top of the module",
                        TargetModule = issue.ModuleName,
                        IssueRule = issue.RuleName,
                        PreviewCode = "Option Explicit\n' (prepended to module)"
                    });
                    break;

                case "MissingErrorHandler":
                    if (moduleLookup.TryGetValue(issue.ModuleName, out var mod))
                    {
                        var procName = ExtractProcName(issue.Description);
                        fixes.Add(new QuickFix
                        {
                            Name = $"Add Error Handler to '{procName}'",
                            Description = "Wrap procedure body in On Error GoTo / handler / Resume template",
                            TargetModule = issue.ModuleName,
                            IssueRule = issue.RuleName,
                            PreviewCode = $"Sub {procName}()\n    On Error GoTo ErrHandler\n    ' ... existing code ...\n    Exit Sub\nErrHandler:\n    MsgBox Err.Description, vbCritical\n    Resume Next\nEnd Sub"
                        });
                    }
                    break;

                case "HardcodedPath":
                    fixes.Add(new QuickFix
                    {
                        Name = "Extract Hardcoded Path",
                        Description = "Create a module-level Const and replace the literal string",
                        TargetModule = issue.ModuleName,
                        IssueRule = issue.RuleName,
                        PreviewCode = $"' At module top:\nConst FILE_PATH As String = \"<extracted path>\"\n' In code:\n' Replace hardcoded string with FILE_PATH"
                    });
                    break;
            }
        }

        return fixes;
    }

    public static string ApplyFix(QuickFix fix, string moduleCode)
    {
        switch (fix.IssueRule)
        {
            case "MissingOptionExplicit":
                return "Option Explicit\n" + moduleCode;

            case "NoErrorHandler":
            {
                var procName = ExtractProcNameFromFix(fix.Name);
                var pattern = new Regex(
                    $@"((?:Public\s+|Private\s+)?(?:Sub|Function)\s+{Regex.Escape(procName)}\s*\([^)]*\))\s*\n",
                    RegexOptions.IgnoreCase);
                return pattern.Replace(moduleCode, m =>
                    $"{m.Groups[1].Value}\n    On Error GoTo ErrHandler_{procName}\n", 1) +
                    $"\n' (Error handler label added — complete the handler manually)";
            }

            case "HardcodedPath":
            {
                var pathMatch = Regex.Match(moduleCode, @"""([A-Z]:\\[^""]+)""", RegexOptions.IgnoreCase);
                if (pathMatch.Success)
                {
                    var path = pathMatch.Groups[1].Value;
                    var constLine = $"Const FILE_PATH As String = \"{path}\"\n";
                    var modified = moduleCode.Replace($"\"{path}\"", "FILE_PATH");
                    return constLine + modified;
                }
                return moduleCode;
            }

            default:
                return moduleCode;
        }
    }

    private static string ExtractProcName(string description)
    {
        var m = Regex.Match(description, @"'(\w+)'");
        return m.Success ? m.Groups[1].Value : "Unknown";
    }

    private static string ExtractProcNameFromFix(string fixName)
    {
        var m = Regex.Match(fixName, @"'(\w+)'");
        return m.Success ? m.Groups[1].Value : "Unknown";
    }
}
