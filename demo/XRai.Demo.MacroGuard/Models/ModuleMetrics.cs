namespace XRai.Demo.MacroGuard.Models;

public class ModuleMetrics
{
    public string ModuleName { get; set; } = "";
    public int TotalLines { get; set; }
    public int CommentLines { get; set; }
    public double CommentRatio => TotalLines > 0 ? Math.Round(100.0 * CommentLines / TotalLines, 1) : 0;
    public int ProcedureCount { get; set; }
    public double AvgProcedureLength { get; set; }
    public int CyclomaticComplexity { get; set; }
    public int DeadProcedures { get; set; }

    public override string ToString() => $"{ModuleName}: {TotalLines} lines, {CommentRatio}% comments, CC={CyclomaticComplexity}";
}
