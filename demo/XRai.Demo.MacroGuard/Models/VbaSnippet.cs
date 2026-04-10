namespace XRai.Demo.MacroGuard.Models;

public class VbaSnippet
{
    public string Name { get; set; } = "";
    public string Category { get; set; } = "";
    public string Description { get; set; } = "";
    public string Code { get; set; } = "";

    public override string ToString() => $"[{Category}] {Name}";
}
