using XRai.Demo.MacroGuard.Models;

namespace XRai.Demo.MacroGuard.Scanner;

public class VbaAnalyzer
{
    private static readonly List<IScannerRule> AllRules = new()
    {
        new MissingOptionExplicitRule(),
        new MissingErrorHandlerRule(),
        new UnusedVariableRule(),
        new HardcodedPathRule(),
        new EmptyProcedureRule(),
        new GoToUsageRule(),
        new ImplicitVariantRule(),
        new LongProcedureRule()
    };

    public List<IScannerRule> GetRules(bool strictMode)
    {
        return AllRules.Where(r => !r.RequiresStrictMode || strictMode).ToList();
    }

    public List<VbaIssue> Analyze(VbaModuleInfo module, List<IScannerRule> rules)
    {
        var issues = new List<VbaIssue>();
        foreach (var rule in rules)
        {
            issues.AddRange(rule.Check(module));
        }
        return issues;
    }

    /// <summary>
    /// Reads all VBA modules from the active workbook via COM interop.
    /// Requires "Trust access to the VBA project object model" in Trust Center.
    /// </summary>
    public static List<VbaModuleInfo> ReadVbProject()
    {
        var modules = new List<VbaModuleInfo>();
        try
        {
            dynamic app = ExcelDna.Integration.ExcelDnaUtil.Application;
            var vbProject = app.ActiveWorkbook.VBProject;

            foreach (var component in vbProject.VBComponents)
            {
                var codeModule = component.CodeModule;
                int lineCount = (int)codeModule.CountOfLines;
                string code = lineCount > 0 ? (string)codeModule.Lines[1, lineCount] : "";

                string type = ((int)component.Type) switch
                {
                    1 => "Standard",
                    2 => "Class",
                    3 => "UserForm",
                    100 => code.Contains("Worksheet") ? "Sheet" : "ThisWorkbook",
                    _ => "Standard"
                };

                var macroNames = ExtractMacroNames(code);

                modules.Add(new VbaModuleInfo
                {
                    Name = (string)component.Name,
                    Type = type,
                    LineCount = lineCount,
                    MacroNames = macroNames,
                    Code = code
                });
            }
        }
        catch (Exception ex)
        {
            System.Diagnostics.Debug.WriteLine($"VBA access error: {ex.Message}");
        }
        return modules;
    }

    public static bool CheckVbaAccess()
    {
        try
        {
            dynamic app = ExcelDna.Integration.ExcelDnaUtil.Application;
            var _ = app.ActiveWorkbook.VBProject.VBComponents.Count;
            return true;
        }
        catch
        {
            return false;
        }
    }

    private static List<string> ExtractMacroNames(string code)
    {
        var names = new List<string>();
        foreach (var line in code.Split('\n'))
        {
            var trimmed = line.TrimStart();
            if (trimmed.StartsWith("Sub ", StringComparison.OrdinalIgnoreCase) ||
                trimmed.StartsWith("Public Sub ", StringComparison.OrdinalIgnoreCase) ||
                trimmed.StartsWith("Private Sub ", StringComparison.OrdinalIgnoreCase) ||
                trimmed.StartsWith("Function ", StringComparison.OrdinalIgnoreCase) ||
                trimmed.StartsWith("Public Function ", StringComparison.OrdinalIgnoreCase) ||
                trimmed.StartsWith("Private Function ", StringComparison.OrdinalIgnoreCase))
            {
                var match = System.Text.RegularExpressions.Regex.Match(trimmed, @"(?:Sub|Function)\s+(\w+)", System.Text.RegularExpressions.RegexOptions.IgnoreCase);
                if (match.Success)
                    names.Add(match.Groups[1].Value);
            }
        }
        return names;
    }
}
