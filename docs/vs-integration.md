# Visual Studio Integration

XRai works with Visual Studio 2022 and 2026 through the Claude Code community extension and MCP support.

## Visual Studio 2022

### Install the Claude Code Extension

1. Open Visual Studio 2022
2. Go to **Extensions > Manage Extensions**
3. Search for "Claude Code" by dliedke
4. Install and restart Visual Studio

The extension adds a Claude Code terminal pane inside Visual Studio. XRai works identically to standalone Claude Code -- the skill is auto-discovered from `~/.claude/skills/xrai-excel/`.

### Recommended Settings

Enable auto-reload on external file changes so that when Claude edits files, Visual Studio picks them up immediately:

1. Go to **Tools > Options > Environment > Documents**
2. Check **Detect when file is changed outside the environment**
3. Check **Auto-load changes, if saved**

This prevents Visual Studio from prompting you every time Claude modifies a file during the build-verify loop.

### Workflow

1. Open your Excel-DNA add-in solution in Visual Studio
2. Open the Claude Code pane
3. Ask Claude to make changes, build, and test
4. Claude edits `.cs`/`.xaml` files, runs `dotnet build`, reloads the `.xll`, and verifies via XRai
5. Visual Studio auto-reloads the changed files

You can also use Visual Studio's built-in build (Ctrl+Shift+B) alongside Claude. Claude detects the build output and uses it.

## Visual Studio 2026

Visual Studio 2026 has native MCP (Model Context Protocol) support.

### Using XRai via MCP

1. Install the XRai skill as normal (extract to `~/.claude/skills/xrai-excel/`)
2. In Visual Studio 2026, go to **Tools > Options > AI > MCP Servers**
3. Add a new MCP server pointing to:

```
Path: %USERPROFILE%\.claude\skills\xrai-excel\bin\XRai.Tool.exe
Args: --mcp
```

4. Visual Studio's AI assistant (Copilot) can now use XRai commands to interact with Excel

### MCP vs Skill

| Feature | Skill (Claude Code) | MCP (VS 2026 / Copilot) |
|---------|---------------------|--------------------------|
| Auto-discovery | Yes (SKILL.md) | Manual configuration |
| Command access | All 283 commands | All 283 commands |
| Agent guidance | SKILL.md provides patterns | Agent must know XRai protocol |
| Setup | Extract zip | Extract zip + configure MCP |

Both approaches provide access to the same 283 commands. The skill system adds higher-level guidance (when to batch, how to verify, common patterns) that MCP alone does not provide.

## File Watcher Considerations

When Claude is editing files through XRai/Claude Code:

- **Visual Studio file watcher** -- enable auto-reload (see settings above)
- **ReSharper / Rider users** -- ensure external file monitoring is enabled in settings
- **Build output** -- Claude uses `dotnet build` from the command line, which writes to `bin/Debug/`. Visual Studio also writes there. No conflict as long as you are not building simultaneously from both.

## Debugging with XRai

You can debug your add-in with Visual Studio's debugger while XRai is connected:

1. Set breakpoints in your add-in code
2. Start debugging (F5) -- this launches Excel with your add-in
3. In a separate Claude Code session, connect XRai: `{"cmd":"wait"}` (waits for Excel to appear)
4. Send commands via XRai -- breakpoints will hit in Visual Studio

This is useful for stepping through button click handlers or ViewModel property setters while XRai drives the interaction.
