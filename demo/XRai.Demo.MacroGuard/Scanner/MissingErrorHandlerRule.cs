using System.Text.RegularExpressions;
using XRai.Demo.MacroGuard.Models;

namespace XRai.Demo.MacroGuard.Scanner;

public class MissingErrorHandlerRule : IScannerRule
{
    public string Name => "MissingErrorHandler";
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
        bool hasErrorHandler = false;

        for (int i = 0; i < lines.Length; i++)
        {
            var line = lines[i];
            var startMatch = ProcStart.Match(line);
            if (startMatch.Success)
            {
                currentProc = startMatch.Groups[3].Value;
                procStartLine = i + 1;
                hasErrorHandler = false;
                continue;
            }

            if (currentProc != null)
            {
                if (line.TrimStart().StartsWith("On Error", StringComparison.OrdinalIgnoreCase))
                    hasErrorHandler = true;

                if (ProcEnd.IsMatch(line))
                {
                    if (!hasErrorHandler)
                    {
                        issues.Add(new VbaIssue
                        {
                            Severity = "Warning",
                            ModuleName = module.Name,
                            LineNumber = procStartLine,
                            RuleName = Name,
                            Description = $"Procedure '{currentProc}' has no error handler (On Error)",
                            SuggestedFix = "Add 'On Error GoTo ErrHandler' and an error handling block"
                        });
                    }
                    currentProc = null;
                }
            }
        }
        return issues;
    }
}
