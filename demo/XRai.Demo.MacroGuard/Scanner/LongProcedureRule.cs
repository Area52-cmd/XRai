using System.Text.RegularExpressions;
using XRai.Demo.MacroGuard.Models;

namespace XRai.Demo.MacroGuard.Scanner;

public class LongProcedureRule : IScannerRule
{
    public string Name => "LongProcedure";
    public bool RequiresStrictMode => false;

    private const int MaxLines = 50;

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

        for (int i = 0; i < lines.Length; i++)
        {
            var startMatch = ProcStart.Match(lines[i]);
            if (startMatch.Success)
            {
                currentProc = startMatch.Groups[3].Value;
                procStartLine = i + 1;
                continue;
            }

            if (currentProc != null && ProcEnd.IsMatch(lines[i]))
            {
                int length = (i + 1) - procStartLine;
                if (length > MaxLines)
                {
                    issues.Add(new VbaIssue
                    {
                        Severity = "Info",
                        ModuleName = module.Name,
                        LineNumber = procStartLine,
                        RuleName = Name,
                        Description = $"Procedure '{currentProc}' is {length} lines long (exceeds {MaxLines} line limit)",
                        SuggestedFix = "Break into smaller, focused procedures"
                    });
                }
                currentProc = null;
            }
        }
        return issues;
    }
}
