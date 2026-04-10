namespace XRai.Core;

/// <summary>
/// Abstraction for reporting rebuild/reload progress. The default implementation
/// just appends to an in-memory list (current behavior). XRai.Studio provides
/// an alternative that ALSO publishes each step to the event bus so the
/// dashboard can render live rebuild progress instead of waiting for the
/// final response.
///
/// Status values:
///   "start"    — step has begun
///   "ok"       — step completed successfully
///   "error"    — step failed
///   "skip"     — step was skipped (e.g. idempotent nuget source add)
///   "warning"  — step completed with warnings
/// </summary>
public interface IStepReporter
{
    /// <summary>
    /// Record a step outcome. Called synchronously from the orchestrator as
    /// each step finishes. Implementations must not block — fire events and
    /// return immediately.
    /// </summary>
    void Report(string step, string status, long elapsedMs, string? detail = null);

    /// <summary>
    /// Record that a step has started (optional; useful for "running" animations
    /// in the dashboard). Default implementations can no-op this.
    /// </summary>
    void Starting(string step) { }

    /// <summary>
    /// All step lines recorded so far, in the exact format the current batched
    /// rebuild response returns. Used to keep backward compatibility with the
    /// existing `steps` field in the rebuild response.
    /// </summary>
    IReadOnlyList<string> Lines { get; }
}

/// <summary>
/// Default in-memory implementation. Matches the existing ReloadOrchestrator
/// behavior exactly — appends to a list and returns it at the end.
/// </summary>
public class ListStepReporter : IStepReporter
{
    private readonly List<string> _lines = new();

    public void Report(string step, string status, long elapsedMs, string? detail = null)
    {
        var line = detail != null
            ? $"{step}: {status} ({elapsedMs} ms) — {detail}"
            : $"{step}: {status} ({elapsedMs} ms)";
        _lines.Add(line);
    }

    public IReadOnlyList<string> Lines => _lines;
}

/// <summary>
/// Step reporter that fans out to TWO sinks simultaneously: an in-memory list
/// (for the batched response) AND a callback the Studio uses to publish a
/// rebuild.step event. Used when the daemon is running with --studio enabled.
/// </summary>
public class TeeStepReporter : IStepReporter
{
    private readonly List<string> _lines = new();
    private readonly Action<string, string, long, string?> _onStep;

    public TeeStepReporter(Action<string, string, long, string?> onStep)
    {
        _onStep = onStep;
    }

    public void Report(string step, string status, long elapsedMs, string? detail = null)
    {
        var line = detail != null
            ? $"{step}: {status} ({elapsedMs} ms) — {detail}"
            : $"{step}: {status} ({elapsedMs} ms)";
        _lines.Add(line);
        try { _onStep(step, status, elapsedMs, detail); }
        catch { /* sink errors must not break the rebuild */ }
    }

    public void Starting(string step)
    {
        try { _onStep(step, "start", 0, null); }
        catch { }
    }

    public IReadOnlyList<string> Lines => _lines;
}
