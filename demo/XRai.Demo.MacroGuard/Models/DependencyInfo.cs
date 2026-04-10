namespace XRai.Demo.MacroGuard.Models;

public class DependencyInfo
{
    public string CallerModule { get; set; } = "";
    public string CallerProcedure { get; set; } = "";
    public List<string> Callees { get; set; } = new();

    public override string ToString() => $"{CallerModule}.{CallerProcedure} calls: {string.Join(", ", Callees)}";
}
