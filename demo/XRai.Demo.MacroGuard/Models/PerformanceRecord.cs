namespace XRai.Demo.MacroGuard.Models;

public class PerformanceRecord
{
    public string MacroName { get; set; } = "";
    public long LastRunMs { get; set; }
    public double AvgRunMs { get; set; }
    public int RunCount { get; set; }

    private long _totalMs;

    public void RecordRun(long elapsedMs)
    {
        LastRunMs = elapsedMs;
        RunCount++;
        _totalMs += elapsedMs;
        AvgRunMs = Math.Round((double)_totalMs / RunCount, 2);
    }

    public override string ToString() => $"{MacroName}: {AvgRunMs:F1}ms avg ({RunCount} runs)";
}
