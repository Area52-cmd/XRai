using System.Diagnostics;
using System.Text.Json.Nodes;

namespace XRai.Studio.Sources.Agents;

/// <summary>
/// Tails the live Claude Code session transcript JSONL and re-emits every
/// event onto the Studio bus as normalized <c>agent.*</c> events. This is
/// the headline Studio feature — it lets the dashboard render a live view
/// of the AI coding agent's work without any integration on Claude Code's
/// side and without consuming any extra API tokens. The transcript is
/// something Claude Code already writes to disk on every turn; we just
/// tail it, parse it, and normalize it.
///
/// Transcript location (Claude Code default):
///   %USERPROFILE%\.claude\projects\{project-hash}\{session-id}.jsonl
///
/// The project-hash is generated from the cwd (e.g. "D:\Code\XRai" becomes
/// "D--Code-Xrai"). We auto-discover the active session by finding the most
/// recently-modified JSONL inside the Claude Code projects directory.
///
/// Transcript line shapes (confirmed empirically via live session profiling):
///   - { type:"user",      message:{ role, content:[{type:"text", text}] } }
///   - { type:"user",      message:{ role, content:[{type:"tool_result", ...}] } }
///   - { type:"assistant", message:{ role, content:[{type:"text", text}] } }
///   - { type:"assistant", message:{ role, content:[{type:"thinking", text}] } }
///   - { type:"assistant", message:{ role, content:[{type:"tool_use", id, name, input}] } }
///
/// Normalized output events (agent-agnostic — the dashboard renders these
/// the same way regardless of which agent adapter produced them):
///   agent.session          — a new active session was detected
///   agent.message.user     — user turn
///   agent.message.text     — assistant text block
///   agent.message.think    — assistant thinking block
///   agent.tool.use         — tool invocation
///   agent.tool.result      — tool result
/// </summary>
public sealed class ClaudeCodeAgentAdapter : IAgentAdapter
{
    public string AgentName => "Claude Code";
    public bool IsConnected => _activeFile != null && !_disposed;

    private readonly EventBus _bus;
    private readonly string? _projectsRoot;
    private readonly CancellationTokenSource _cts = new();
    private Thread? _thread;
    private volatile bool _disposed;
    private string? _activeFile;
    private long _position;

    public int PollIntervalMs { get; set; } = 200;
    public string? ActiveFile => _activeFile;

    public ClaudeCodeAgentAdapter(EventBus bus, string? projectsRoot = null)
    {
        _bus = bus;
        _projectsRoot = projectsRoot ?? DefaultProjectsRoot();
    }

    public static string DefaultProjectsRoot()
    {
        var home = Environment.GetFolderPath(Environment.SpecialFolder.UserProfile);
        return Path.Combine(home, ".claude", "projects");
    }

    public void Start()
    {
        if (_thread != null) return;
        _thread = new Thread(Loop)
        {
            IsBackground = true,
            Name = "xrai-studio-claude-code-tail"
        };
        _thread.Start();
    }

    private void Loop()
    {
        while (!_disposed && !_cts.IsCancellationRequested)
        {
            try
            {
                var file = DiscoverActiveTranscript();
                if (file == null)
                {
                    Thread.Sleep(1000);
                    continue;
                }

                if (_activeFile != file)
                {
                    _activeFile = file;
                    try { _position = new FileInfo(file).Length; }
                    catch { _position = 0; }

                    _bus.Publish(StudioEvent.Now("agent.session", "agent", new JsonObject
                    {
                        ["agent"] = AgentName,
                        ["file"] = file,
                        ["startPosition"] = _position,
                    }));
                }

                TailOnce(file);
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"ClaudeCode tail loop error: {ex.Message}");
            }

            try { Thread.Sleep(PollIntervalMs); }
            catch { break; }
        }
    }

    private void TailOnce(string file)
    {
        try
        {
            using var fs = new FileStream(file, FileMode.Open, FileAccess.Read, FileShare.ReadWrite | FileShare.Delete);
            if (fs.Length < _position)
            {
                _position = fs.Length;
                return;
            }

            fs.Seek(_position, SeekOrigin.Begin);
            using var reader = new StreamReader(fs);
            string? line;
            while ((line = reader.ReadLine()) != null)
            {
                if (!string.IsNullOrWhiteSpace(line))
                    PublishLine(line);
            }
            _position = fs.Position;
        }
        catch (IOException)
        {
            // File locked by Claude Code mid-write — try again next tick
        }
        catch (Exception ex)
        {
            Debug.WriteLine($"ClaudeCode tail read error: {ex.Message}");
        }
    }

    private void PublishLine(string rawLine)
    {
        try
        {
            var node = JsonNode.Parse(rawLine);
            if (node is not JsonObject obj) return;

            var msg = obj["message"] as JsonObject;
            if (msg == null) return;

            var role = msg["role"]?.GetValue<string>();
            var content = msg["content"];
            var ts = obj["timestamp"]?.GetValue<string>();
            var uuid = obj["uuid"]?.GetValue<string>();
            var gitBranch = obj["gitBranch"]?.GetValue<string>();
            var cwd = obj["cwd"]?.GetValue<string>();

            if (content is JsonArray arr)
            {
                foreach (var block in arr)
                {
                    if (block is not JsonObject b) continue;
                    var blockType = b["type"]?.GetValue<string>();

                    switch (blockType)
                    {
                        case "text":
                            EmitText(role, b, ts, uuid, gitBranch, cwd);
                            break;
                        case "thinking":
                            EmitThinking(b, ts, uuid);
                            break;
                        case "tool_use":
                            EmitToolUse(b, ts, uuid, cwd);
                            break;
                        case "tool_result":
                            EmitToolResult(b, ts, uuid);
                            break;
                    }
                }
            }
            else if (content is JsonValue v)
            {
                string? text = null;
                try { text = v.GetValue<string>(); } catch { }
                if (text != null)
                {
                    _bus.Publish(StudioEvent.Now("agent.message.user", "agent", new JsonObject
                    {
                        ["agent"] = AgentName,
                        ["text"] = text,
                        ["uuid"] = uuid,
                        ["sourceTs"] = ts,
                    }));
                }
            }
        }
        catch (Exception ex)
        {
            Debug.WriteLine($"ClaudeCode tail publish error: {ex.Message}");
        }
    }

    private void EmitText(string? role, JsonObject block, string? ts, string? uuid, string? gitBranch, string? cwd)
    {
        var text = block["text"]?.GetValue<string>() ?? "";
        var kind = role == "user" ? "agent.message.user" : "agent.message.text";
        _bus.Publish(StudioEvent.Now(kind, "agent", new JsonObject
        {
            ["agent"] = AgentName,
            ["text"] = text,
            ["uuid"] = uuid,
            ["sourceTs"] = ts,
            ["gitBranch"] = gitBranch,
            ["cwd"] = cwd,
        }));
    }

    private void EmitThinking(JsonObject block, string? ts, string? uuid)
    {
        var text = block["text"]?.GetValue<string>() ?? "";
        _bus.Publish(StudioEvent.Now("agent.message.think", "agent", new JsonObject
        {
            ["agent"] = AgentName,
            ["text"] = text,
            ["uuid"] = uuid,
            ["sourceTs"] = ts,
        }));
    }

    private void EmitToolUse(JsonObject block, string? ts, string? uuid, string? cwd)
    {
        var name = block["name"]?.GetValue<string>() ?? "?";
        var id = block["id"]?.GetValue<string>();
        var input = block["input"]?.DeepClone();

        var data = new JsonObject
        {
            ["agent"] = AgentName,
            ["toolName"] = name,
            ["toolUseId"] = id,
            ["uuid"] = uuid,
            ["sourceTs"] = ts,
            ["cwd"] = cwd,
            ["input"] = input,
        };

        if (input is JsonObject inputObj)
        {
            CopyStringField(inputObj, data, "file_path", "filePath");
            CopyStringField(inputObj, data, "old_string", "oldString");
            CopyStringField(inputObj, data, "new_string", "newString");
            CopyStringField(inputObj, data, "content", "fullContent");
            CopyStringField(inputObj, data, "command", "command");
            CopyStringField(inputObj, data, "description", "description");
            CopyStringField(inputObj, data, "pattern", "pattern");
            CopyStringField(inputObj, data, "path", "path");
            CopyStringField(inputObj, data, "prompt", "prompt");
            CopyStringField(inputObj, data, "url", "url");
            CopyStringField(inputObj, data, "query", "query");
        }

        _bus.Publish(StudioEvent.Now("agent.tool.use", "agent", data));
    }

    private static void CopyStringField(JsonObject from, JsonObject to, string fromKey, string toKey)
    {
        try
        {
            if (from[fromKey] is JsonValue v)
            {
                var s = v.GetValue<string>();
                if (!string.IsNullOrEmpty(s)) to[toKey] = s;
            }
        }
        catch { }
    }

    private void EmitToolResult(JsonObject block, string? ts, string? uuid)
    {
        var toolUseId = block["tool_use_id"]?.GetValue<string>();
        var isError = block["is_error"]?.GetValue<bool>() ?? false;
        var content = block["content"];

        string? summary = null;
        if (content is JsonValue v)
        {
            try { summary = v.GetValue<string>(); } catch { }
        }
        else if (content is JsonArray arr)
        {
            var sb = new System.Text.StringBuilder();
            foreach (var c in arr)
            {
                if (c is JsonObject co && co["type"]?.GetValue<string>() == "text")
                {
                    try { sb.AppendLine(co["text"]?.GetValue<string>()); } catch { }
                }
            }
            summary = sb.ToString().TrimEnd();
        }

        const int MaxSummaryLen = 4000;
        if (summary != null && summary.Length > MaxSummaryLen)
        {
            summary = summary.Substring(0, MaxSummaryLen) + $"\n... [{summary.Length - MaxSummaryLen} more bytes truncated]";
        }

        _bus.Publish(StudioEvent.Now("agent.tool.result", "agent", new JsonObject
        {
            ["agent"] = AgentName,
            ["toolUseId"] = toolUseId,
            ["isError"] = isError,
            ["summary"] = summary,
            ["uuid"] = uuid,
            ["sourceTs"] = ts,
        }));
    }

    private string? DiscoverActiveTranscript()
    {
        if (_projectsRoot == null || !Directory.Exists(_projectsRoot)) return null;

        try
        {
            string? best = null;
            DateTime bestWrite = DateTime.MinValue;

            foreach (var projectDir in Directory.EnumerateDirectories(_projectsRoot))
            {
                try
                {
                    foreach (var file in Directory.EnumerateFiles(projectDir, "*.jsonl", SearchOption.TopDirectoryOnly))
                    {
                        try
                        {
                            var fi = new FileInfo(file);
                            if (fi.LastWriteTimeUtc > bestWrite)
                            {
                                bestWrite = fi.LastWriteTimeUtc;
                                best = file;
                            }
                        }
                        catch { }
                    }
                }
                catch { }
            }

            return best;
        }
        catch
        {
            return null;
        }
    }

    public void Dispose()
    {
        if (_disposed) return;
        _disposed = true;
        try { _cts.Cancel(); } catch { }
        try { _thread?.Join(1000); } catch { }
        _cts.Dispose();
    }
}
