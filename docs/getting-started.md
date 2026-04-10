# Getting Started with XRai

This guide takes you from zero to a working AI-driven Excel add-in in under 10 minutes.

## Prerequisites

- **Windows 10 or 11** (x64)
- **Excel 2016+** or Microsoft 365
- **.NET 8 SDK** or later -- only needed if you are building add-ins ([download](https://dotnet.microsoft.com/download))

No .NET SDK is required to run XRai itself -- the CLI ships self-contained.

## Step 1: Install the Skill

Download the latest `xrai-excel-skill.zip` from the [Releases](https://github.com/user/xrai/releases) page and extract it to your Claude Code skills directory:

```powershell
Expand-Archive xrai-excel-skill.zip -DestinationPath "$env:USERPROFILE\.claude\skills\"
```

Verify the install:

```powershell
Test-Path "$env:USERPROFILE\.claude\skills\xrai-excel\SKILL.md"
# Should return True
```

This places the XRai skill at `~/.claude/skills/xrai-excel/`. Claude Code automatically discovers it in every project from now on.

### What gets installed

```
~/.claude/skills/xrai-excel/
  SKILL.md                      Auto-loaded by Claude Code
  bin/XRai.Tool.exe             Self-contained CLI (~68 MB)
  packages/XRai.Hooks.*.nupkg   NuGet package for add-in hooks
  templates/                    Project templates and wiring samples
  docs/                         Command catalog, control matrix, patterns
```

## Step 2: Create Your First Add-In

Open a terminal in any empty folder and start Claude Code:

```bash
claude
```

Then say:

> Create a new Excel add-in called StockTracker with a dark-themed task pane that shows a portfolio grid, a refresh button, and a status label.

Claude will:

1. Run `XRai.Tool.exe init StockTracker` to scaffold the project
2. Generate the WPF task pane XAML and ViewModel
3. Wire `Pilot.Start()`, `Pilot.Expose()`, and `Pilot.ExposeModel()` into `AutoOpen()`
4. Build the project with `dotnet build`
5. Launch Excel and load the `.xll`
6. Attach via XRai and verify the pane controls are visible
7. Take a screenshot to confirm everything is working

## Step 3: The Build-Verify Loop

Once the add-in is loaded, the development loop is:

1. **Ask Claude to make a change** -- add a feature, fix a bug, restyle the UI
2. **Claude edits the code** -- `.cs`, `.xaml`, `.csproj` files
3. **Claude rebuilds** -- `dotnet build`
4. **Claude reloads** -- sends `{"cmd":"reload"}` to hot-reload the `.xll`
5. **Claude verifies** -- reads pane controls, checks ViewModel state, takes a screenshot

You never leave the conversation. Claude handles the entire edit-build-test cycle autonomously.

## For Existing Add-Ins

If you already have an Excel-DNA add-in and want to make it XRai-enabled:

### 1. Add the XRai.Hooks NuGet package

```bash
dotnet add package XRai.Hooks --version "1.0.*"
```

### 2. Wire three calls in AutoOpen

```csharp
using XRai.Hooks;

public class MyAddIn : IExcelAddIn
{
    public void AutoOpen()
    {
        Pilot.Start();                    // Start the hooks pipe server
        // After your task pane is created:
        Pilot.Expose(myTaskPane);         // Expose WPF controls to XRai
        Pilot.ExposeModel(myViewModel);   // Expose ViewModel properties
    }

    public void AutoClose()
    {
        Pilot.Stop();
    }
}
```

### 3. Name your controls

Every interactive control in your XAML needs an `x:Name`:

```xml
<TextBox x:Name="SpotInput" Text="{Binding Spot}" />
<Button x:Name="CalcButton" Content="Calculate" />
<DataGrid x:Name="TradesGrid" ItemsSource="{Binding Trades}" />
```

Controls without `x:Name` are invisible to XRai.

### 4. Build, load, and go

```bash
dotnet build
```

Load the `.xll` in Excel, then open Claude Code in the project directory. Claude will detect the XRai hooks and start interacting with your pane.

## Next Steps

- [Command Reference (283 commands)](commands.md) -- full catalog of everything XRai can do
- [Task Pane & ViewModel Guide](pane-controls.md) -- deep dive into WPF/WinForms control automation
- [Dialog & Folder Picker Guide](dialogs.md) -- driving native dialogs and file pickers
- [Testing & Assertions Guide](testing.md) -- automated testing with assertions and screenshot diff
- [Troubleshooting](troubleshooting.md) -- common errors and recovery flows
- [Architecture](architecture.md) -- how XRai works under the hood
