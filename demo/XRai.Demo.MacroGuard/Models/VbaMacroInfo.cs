namespace XRai.Demo.MacroGuard.Models;

public class VbaMacroInfo
{
    public string Name { get; set; } = "";
    public string ModuleName { get; set; } = "";
    public int StartLine { get; set; }
    public int EndLine { get; set; }
    public string Parameters { get; set; } = "";
    public bool IsFunction { get; set; }

    public override string ToString() => IsFunction ? $"Function {Name}" : $"Sub {Name}";
}
