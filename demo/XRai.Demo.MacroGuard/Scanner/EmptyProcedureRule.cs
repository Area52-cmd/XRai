using System.Text.RegularExpressions;
using XRai.Demo.MacroGuard.Models;

namespace XRai.Demo.MacroGuard.Scanner;

public class EmptyProcedureRule : IScannerRule
{
    public string Name => "EmptyProcedure";
    public bool RequiresStrictMode => false;

    private static readonly Regex ProcStart = new(
        @"^\s*(Public\s+|Private\s+)?(Sub|Function)\s+(\w+)",
        RegexOptions.IgnoreCase | RegexOptions.Compiled);

    private static readonly Regex ProcEnd = new(
        @"^\s*End\s+(Sub|Function)",
        RegexOptions.IgnoreCase | RegexOptions.Compiled);

    public List<VbaIssue> Check(VbaModuleInfo module)
    {
        var issues = new List<VbaIssue>();
        var lines = module.Code.Split('\n');

        string? currentProc = null;
        int procStartLine = 0;
        bool hasExecutableCode = false;

        for (int i = 0; i < lines.Length; i++)
        {
            var line = lines[i];
            var startMatch = ProcStart.Match(line);
            if (startMatch.Success)
            {
                currentProc = startMatch.Groups[3].Value;
                procStartLine = i + 1;
                hasExecutableCode = false;
                continue;
            }

            if (currentProc != null)
            {
                if (ProcEnd.IsMatch(line))
                {
                    if (!hasExecutableCode)
                    {
                        issues.Add(new VbaIssue
                        {
                            Severity = "Info",
                            ModuleName = module.Name,
                            LineNumber = procStartLine,
                            RuleName = Name,
                            Description = $"Procedure '{currentProc}' contains no executable code",
                            SuggestedFix = "Add implementation or remove the empty procedure"
                        });
                    }
                    currentProc = null;
                }
                else
                {
                    var trimmed = line.Trim();
                    if (!string.IsNullOrWhiteSpace(trimmed) && !trimmed.StartsWith("'") && !trimmed.StartsWith("Dim ", StringComparison.OrdinalIgnoreCase))
                        hasExecutableCode = true;
                }
            }
        }
        return issues;
    }
}
