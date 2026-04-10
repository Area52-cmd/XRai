namespace XRai.Demo.MacroGuard.Models;

public class VbaIssue
{
    public string Severity { get; set; } = "Warning";
    public string ModuleName { get; set; } = "";
    public int LineNumber { get; set; }
    public string RuleName { get; set; } = "";
    public string Description { get; set; } = "";
    public string SuggestedFix { get; set; } = "";

    public override string ToString() => $"[{Severity}] {ModuleName}:{LineNumber} — {Description}";
}
