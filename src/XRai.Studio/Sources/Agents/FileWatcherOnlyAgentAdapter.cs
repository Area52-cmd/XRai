namespace XRai.Studio.Sources.Agents;

/// <summary>
/// Fallback adapter used when no recognized agent transcript is available.
/// This adapter does not produce any agent.* events on its own — it's a
/// placeholder so the dashboard reports "no agent connected" cleanly. The
/// FileWatcherSource still provides file.changed events, so users building
/// with an unrecognized agent (hand-editing files, running their own
/// scripts) still see real-time file change visibility.
///
/// Use this adapter when:
///   - Claude Code is not installed
///   - The user explicitly wants to disable agent tailing
///   - A custom workflow edits files directly without a transcript
/// </summary>
public sealed class FileWatcherOnlyAgentAdapter : IAgentAdapter
{
    public string AgentName => "(no agent)";
    public bool IsConnected => false;

    public FileWatcherOnlyAgentAdapter(EventBus bus) { /* nothing to do */ }

    public void Start() { /* no-op */ }
    public void Dispose() { /* no-op */ }
}
