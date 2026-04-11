namespace XRai.Studio.Sources;

/// <summary>
/// Abstraction over the source of "agent activity" (transcripts, tool calls,
/// edits) so Studio is not tied to any one AI coding agent. Today's concrete
/// implementations:
///
///   <see cref="ClaudeCodeAgentAdapter"/>
///     Tails ~/.claude/projects/**/*.jsonl (Claude Code)
///
///   Future (stubs shipped as interfaces that just need a parser):
///     CodexAgentAdapter        — OpenAI Codex CLI / Codex SDK transcripts
///     CursorAgentAdapter       — Cursor IDE chat history
///     AiderAgentAdapter        — Aider chat/history files
///     FileWatcherOnlyAdapter   — generic fallback for any agent that only
///                                edits files (no reasoning stream). Still
///                                useful: every edit is visible via the
///                                FileWatcherSource, just no thinking text.
///
/// Every adapter normalizes its native events into the same set of bus
/// event kinds so the dashboard can render them uniformly regardless of
/// which agent produced them:
///
///   agent.session         — a new session was detected / attached
///   agent.message.user    — a user turn (prompt)
///   agent.message.text    — an assistant text response block
///   agent.message.think   — assistant thinking / reasoning block
///   agent.tool.use        — the agent invoked a tool (Edit, Write, Bash, etc.)
///   agent.tool.result     — the tool returned
///
/// This keeps the bus shape stable and the dashboard fully agent-agnostic.
/// </summary>
public interface IAgentAdapter : IDisposable
{
    /// <summary>Human-readable agent identifier — shown in the dashboard UI.</summary>
    string AgentName { get; }

    /// <summary>True if the adapter found an active session and is actively reading.</summary>
    bool IsConnected { get; }

    /// <summary>Start background reading. Publishes events onto the supplied bus.</summary>
    void Start();
}
