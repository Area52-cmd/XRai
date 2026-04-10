namespace XRai.Demo.MacroGuard.Models;

public class VbaSnapshot
{
    public DateTime Timestamp { get; set; } = DateTime.Now;
    public string Label { get; set; } = "";
    public Dictionary<string, string> ModuleCode { get; set; } = new();

    public override string ToString() => $"[{Timestamp:yyyy-MM-dd HH:mm:ss}] {Label}";
}
