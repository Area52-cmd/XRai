namespace XRai.Studio.Sources.Agents;

/// <summary>
/// Chooses the best available agent adapter at Studio startup. Order of
/// preference:
///   1. Claude Code — if %USERPROFILE%\.claude\projects\ exists and contains
///      at least one JSONL file with a recent last-write time
///   2. Future: Codex, Cursor, Aider — stubs for when those adapters land
///   3. File-watcher-only fallback — no agent events, just file change events
///
/// The factory returns the first adapter whose probe succeeds. Callers can
/// override with an explicit adapter instance if auto-detection is not
/// desired.
/// </summary>
public static class AgentAdapterFactory
{
    /// <summary>
    /// Auto-detect the active coding agent on this machine and return a
    /// ready-to-Start adapter. Never throws — falls back to the no-op
    /// adapter if nothing matches.
    /// </summary>
    public static IAgentAdapter Detect(EventBus bus)
    {
        // Probe Claude Code first — the current reference implementation.
        if (ClaudeCodeAvailable())
            return new ClaudeCodeAgentAdapter(bus);

        // Future probes would go here:
        //   if (CodexAvailable()) return new CodexAgentAdapter(bus);
        //   if (CursorAvailable()) return new CursorAgentAdapter(bus);
        //   if (AiderAvailable()) return new AiderAgentAdapter(bus);

        return new FileWatcherOnlyAgentAdapter(bus);
    }

    private static bool ClaudeCodeAvailable()
    {
        try
        {
            var root = ClaudeCodeAgentAdapter.DefaultProjectsRoot();
            if (!Directory.Exists(root)) return false;

            // Any JSONL in any project subfolder counts
            foreach (var project in Directory.EnumerateDirectories(root))
            {
                try
                {
                    if (Directory.EnumerateFiles(project, "*.jsonl").Any())
                        return true;
                }
                catch { }
            }
            return false;
        }
        catch
        {
            return false;
        }
    }
}
