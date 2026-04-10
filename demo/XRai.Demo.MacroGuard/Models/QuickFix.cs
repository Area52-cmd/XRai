namespace XRai.Demo.MacroGuard.Models;

public class QuickFix
{
    public string Name { get; set; } = "";
    public string Description { get; set; } = "";
    public string TargetModule { get; set; } = "";
    public string PreviewCode { get; set; } = "";
    public string IssueRule { get; set; } = "";

    public override string ToString() => $"{Name} — {TargetModule}";
}
