using System.Diagnostics;
using System.IO;
using System.Text.Json;
using System.Text.Json.Nodes;

namespace XRai.Mcp;

/// <summary>
/// Auto-config command that writes MCP config to all supported AI agents:
/// Claude Code, Cursor, Codex, Windsurf, VS Code.
/// Usage: XRai.Mcp.exe setup
/// </summary>
public static class SetupCommand
{
    public static void Run()
    {
        var exePath = Environment.ProcessPath
            ?? Path.Combine(AppContext.BaseDirectory, "XRai.Mcp.exe");

        // Normalize to forward slashes for JSON compatibility
        var exePathJson = exePath.Replace("\\", "/");

        Console.Error.WriteLine();
        Console.Error.WriteLine("  XRai MCP Server — Setup");
        Console.Error.WriteLine($"  Binary: {exePath}");
        Console.Error.WriteLine();

        int success = 0;
        int failed = 0;

        // 1. Claude Code — use CLI if available
        if (TryClaudeCode(exePath)) success++; else failed++;

        // 2. Cursor — ~/.cursor/mcp.json
        if (TryWriteJsonConfig(
            Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.UserProfile), ".cursor", "mcp.json"),
            "Cursor", exePathJson)) success++; else failed++;

        // 3. Codex — ~/.codex/config.toml (not JSON — append MCP entry)
        if (TryWriteCodexConfig(exePathJson)) success++; else failed++;

        // 4. Windsurf — ~/.codeium/windsurf/mcp_config.json
        if (TryWriteJsonConfig(
            Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.UserProfile), ".codeium", "windsurf", "mcp_config.json"),
            "Windsurf", exePathJson)) success++; else failed++;

        // 5. VS Code — .vscode/mcp.json (project-local)
        if (TryWriteVsCodeConfig(exePathJson)) success++; else failed++;

        Console.Error.WriteLine();
        Console.Error.WriteLine($"  Done. {success} configured, {failed} skipped/failed.");
        Console.Error.WriteLine();
    }

    private static bool TryClaudeCode(string exePath)
    {
        try
        {
            var psi = new ProcessStartInfo("claude", $"mcp add xrai -s user -- \"{exePath}\"")
            {
                UseShellExecute = false,
                RedirectStandardOutput = true,
                RedirectStandardError = true,
                CreateNoWindow = true
            };
            var proc = Process.Start(psi);
            if (proc == null) throw new Exception("Failed to start claude CLI");
            proc.WaitForExit(10000);

            if (proc.ExitCode == 0)
            {
                Console.Error.WriteLine("  [OK]   Claude Code — registered via `claude mcp add`");
                return true;
            }
            else
            {
                var err = proc.StandardError.ReadToEnd().Trim();
                Console.Error.WriteLine($"  [SKIP] Claude Code — claude CLI returned exit code {proc.ExitCode}: {err}");
                return false;
            }
        }
        catch (Exception ex)
        {
            Console.Error.WriteLine($"  [SKIP] Claude Code — claude CLI not found ({ex.Message})");
            return false;
        }
    }

    private static bool TryWriteJsonConfig(string configPath, string agentName, string exePath)
    {
        try
        {
            var dir = Path.GetDirectoryName(configPath)!;
            Directory.CreateDirectory(dir);

            JsonObject root;
            if (File.Exists(configPath))
            {
                var existing = File.ReadAllText(configPath);
                root = JsonNode.Parse(existing) as JsonObject ?? new JsonObject();
            }
            else
            {
                root = new JsonObject();
            }

            // Ensure mcpServers object exists
            if (root["mcpServers"] is not JsonObject servers)
            {
                servers = new JsonObject();
                root["mcpServers"] = servers;
            }

            // Add/overwrite xrai entry
            servers["xrai"] = new JsonObject
            {
                ["command"] = exePath,
                ["args"] = new JsonArray()
            };

            var options = new JsonSerializerOptions { WriteIndented = true };
            File.WriteAllText(configPath, root.ToJsonString(options));

            Console.Error.WriteLine($"  [OK]   {agentName} — wrote {configPath}");
            return true;
        }
        catch (Exception ex)
        {
            Console.Error.WriteLine($"  [SKIP] {agentName} — {ex.Message}");
            return false;
        }
    }

    private static bool TryWriteCodexConfig(string exePath)
    {
        try
        {
            var configPath = Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.UserProfile),
                ".codex", "config.toml");

            Directory.CreateDirectory(Path.GetDirectoryName(configPath)!);

            // Read existing or create new
            var lines = File.Exists(configPath)
                ? new List<string>(File.ReadAllLines(configPath))
                : new List<string>();

            // Check if xrai is already configured
            bool hasXrai = lines.Any(l => l.Contains("[mcp_servers.xrai]"));

            if (!hasXrai)
            {
                if (lines.Count > 0 && !string.IsNullOrWhiteSpace(lines[^1]))
                    lines.Add("");

                lines.Add("[mcp_servers.xrai]");
                lines.Add($"command = \"{exePath}\"");
                lines.Add("args = []");

                File.WriteAllLines(configPath, lines);
            }

            Console.Error.WriteLine($"  [OK]   Codex — wrote {configPath}");
            return true;
        }
        catch (Exception ex)
        {
            Console.Error.WriteLine($"  [SKIP] Codex — {ex.Message}");
            return false;
        }
    }

    private static bool TryWriteVsCodeConfig(string exePath)
    {
        try
        {
            // Write to current directory's .vscode/mcp.json
            var vsCodeDir = Path.Combine(Directory.GetCurrentDirectory(), ".vscode");
            var configPath = Path.Combine(vsCodeDir, "mcp.json");

            Directory.CreateDirectory(vsCodeDir);

            JsonObject root;
            if (File.Exists(configPath))
            {
                var existing = File.ReadAllText(configPath);
                root = JsonNode.Parse(existing) as JsonObject ?? new JsonObject();
            }
            else
            {
                root = new JsonObject();
            }

            if (root["servers"] is not JsonObject servers)
            {
                servers = new JsonObject();
                root["servers"] = servers;
            }

            servers["xrai"] = new JsonObject
            {
                ["type"] = "stdio",
                ["command"] = exePath,
                ["args"] = new JsonArray()
            };

            var options = new JsonSerializerOptions { WriteIndented = true };
            File.WriteAllText(configPath, root.ToJsonString(options));

            Console.Error.WriteLine($"  [OK]   VS Code — wrote {configPath}");
            return true;
        }
        catch (Exception ex)
        {
            Console.Error.WriteLine($"  [SKIP] VS Code — {ex.Message}");
            return false;
        }
    }
}
