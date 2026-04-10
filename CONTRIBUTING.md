# Contributing to XRai

## Prerequisites

- **Windows 10/11** (x64, desktop session required)
- **.NET 8 SDK** or later -- [download](https://dotnet.microsoft.com/download)
- **Microsoft Excel 2016+** or Microsoft 365
- **Visual Studio 2022/2026** (optional, for IDE development)

Verify your environment:

```powershell
dotnet --version    # Should be 8.x or later
where excel         # Should find Excel
```

## Building from Source

```powershell
git clone <repo-url> XRai
cd XRai
dotnet build XRai.sln
```

This builds all 9 projects: 7 source projects, 1 test project, and 1 demo add-in.

## Running Tests

### Unit Tests

Unit tests run without Excel installed:

```powershell
dotnet test tests/XRai.Tests.Unit
```

### Live Integration Test

The torture test runs 148+ commands against a live Excel instance:

```powershell
powershell -ExecutionPolicy Bypass -File test-harness.ps1
```

This requires Excel to be installed and will launch/close it automatically.

## Building the Skill Package

The skill package is the distributable ZIP that end users install:

```powershell
powershell -ExecutionPolicy Bypass -File build-skill.ps1
```

This:

1. Publishes `XRai.Tool` as a self-contained executable
2. Packs `XRai.Hooks` as a NuGet package
3. Assembles the skill directory from `skill-template/`
4. Generates `docs/commands.md` from the command registry
5. Produces `dist/xrai-excel-skill/` and `dist/xrai-excel-skill.zip`

## Adding New Commands

All commands follow the same pattern:

1. **Define the handler** in the appropriate Ops file under `src/XRai.Com/` (for COM commands) or `src/XRai.HooksClient/` (for hooks commands)

2. **Register the command** in `CommandRouter.cs`:

```csharp
Register("mycommand", async (args, session) =>
{
    var value = args.GetString("param");
    // ... do work ...
    return Response.Ok(new { result = value });
});
```

3. **Return a Response** -- always use `Response.Ok(...)` or `Response.Fail(...)`:

```csharp
return Response.Ok(new { value = 42 });
return Response.Fail("Something went wrong");
```

4. **Add to the help registry** so `{"cmd":"help"}` lists it

5. **Add a unit test** in `tests/XRai.Tests.Unit/`

6. **Update the command catalog** in `skill-template/docs/commands.md`

### COM Commands

COM commands interact with Excel through `Microsoft.Office.Interop.Excel`. Every COM object must be tracked with `ComGuard`:

```csharp
using var guard = new ComGuard();
var sheet = guard.Track(session.GetActiveSheet());
var range = guard.Track(sheet.Range["A1"]);
var value = range.Value2;
```

Never write two-dot expressions -- they create intermediate COM objects that leak and cause zombie Excel processes.

### Hooks Commands

Hooks commands send JSON over the named pipe to `XRai.Hooks` running inside the add-in. The hooks library handles marshaling to the WPF UI thread.

## Project Structure

```
src/
  XRai.Core/            Command router, Repl, Response, EventStream
  XRai.Com/             COM interop (17 Ops files, 130+ commands)
  XRai.Hooks/           NuGet library for add-ins (Pilot, PipeServer, ControlAdapter)
  XRai.HooksClient/     Named pipe client
  XRai.UI/              FlaUI ribbon/dialog driver
  XRai.Vision/          Win32 PrintWindow screenshot capture
  XRai.Tool/            CLI entry point, doctor, init
  XRai.Mcp/             MCP server
tests/
  XRai.Tests.Unit/      xUnit tests
demo/
  XRai.Demo.PortfolioAddin/   Stock portfolio tracker demo
  XRai.Demo.MacroGuard/       VBA management demo
skill-template/         Source of truth for the skill distribution
  SKILL.md              Agent instructions
  docs/                 Command catalog, control matrix, patterns
  reference/            Deep-dive reference docs
  templates/            Project scaffolding templates
```

## Key Design Constraints

- **Target `net8.0-windows`** for all projects -- Windows APIs are required for COM, WPF, and Win32
- **System.Text.Json only** -- no Newtonsoft dependency anywhere
- **Every COM object must use ComGuard** -- deterministic release prevents zombie processes
- **All hooks operations marshal to UI thread** -- WPF controls require it
- **Responses always have an `ok` field** -- `true` or `false`, always present
- **Commands are flat and terse** -- `{"cmd":"read","ref":"A1"}`, not deeply nested
- **XRai.Hooks has zero heavyweight dependencies** -- just System.IO.Pipes and System.Text.Json
