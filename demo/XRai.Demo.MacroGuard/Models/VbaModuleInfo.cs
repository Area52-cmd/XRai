namespace XRai.Demo.MacroGuard.Models;

public class VbaModuleInfo
{
    public string Name { get; set; } = "";
    public string Type { get; set; } = "Standard";
    public int LineCount { get; set; }
    public List<string> MacroNames { get; set; } = new();
    public string Code { get; set; } = "";

    public override string ToString() => $"{Name} ({Type}, {LineCount} lines)";
}
