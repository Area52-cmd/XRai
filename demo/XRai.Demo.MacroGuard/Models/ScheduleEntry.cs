namespace XRai.Demo.MacroGuard.Models;

public class ScheduleEntry
{
    public string MacroName { get; set; } = "";
    public string TriggerType { get; set; } = "Manual";
    public int IntervalSeconds { get; set; } = 60;
    public bool IsEnabled { get; set; } = true;
    public DateTime? LastRun { get; set; }
    public DateTime? NextRun { get; set; }

    public override string ToString() => $"{MacroName} [{TriggerType}] {(IsEnabled ? "Active" : "Disabled")}";
}
