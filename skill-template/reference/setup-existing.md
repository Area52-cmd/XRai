# Flow A — Migrating an Existing Excel-DNA Add-in

## One-time machine prep

Add the skill's packages folder as a NuGet source (once per machine):

```bash
dotnet nuget add source "$HOME/.claude/skills/xrai-excel/packages" --name XRai-Skill-Local
```

Verify: `dotnet nuget list source` — look for `XRai-Skill-Local`. If it already exists, skip.

## A1. Detect the project

```bash
grep -r "ExcelDna.AddIn" --include="*.csproj" .
grep -r "IExcelAddIn"    --include="*.cs"    .
grep -rE "\[ExcelFunction|\[ExcelCommand" --include="*.cs" .
```

Any match confirms Excel-DNA. If nothing matches, stop and ask the user.

## A2. Check if XRai.Hooks is already wired

```bash
grep -r "XRai.Hooks" --include="*.csproj" .
grep -r "Pilot\."   --include="*.cs"    .
```

If both return hits, skip to Verification (step A7).

## A3. Add XRai.Hooks PackageReference

Edit the `.csproj`:

```xml
<PackageReference Include="XRai.Hooks" Version="1.0.0-*" />
```

**Critical**: use `Version="1.0.0-*"` (wildcard), NOT pinned `1.0.0`. The wildcard resolves to the newest matching package.

Alternative CLI form:
```bash
dotnet add <AddinProject>.csproj package XRai.Hooks --version "1.0.0-*" --source XRai-Skill-Local
dotnet restore
```

## A4. Wire Pilot into AddInEntry

Find the class implementing `IExcelAddIn`. Add:

```csharp
using XRai.Hooks;

public class AddInEntry : IExcelAddIn
{
    public void AutoOpen()
    {
        Pilot.Start();
        Pilot.ExposeModel(App.Container.Resolve<MainViewModel>());
        // Pilot.Expose(pane) — see pane creation patterns below
    }

    public void AutoClose()
    {
        Pilot.Stop();
    }
}
```

### Three pane creation patterns for `Pilot.Expose(pane)`

1. **Pane created at startup** — call `Pilot.Expose(pane)` right after `new MyPane()` in `AutoOpen`.
2. **Pane created on ribbon click** — call `Pilot.Expose(pane)` inside the button handler, after constructing the UserControl.
3. **Pane inside `ExcelAsyncUtil.QueueAsMacro`** — call `Pilot.Expose(pane)` inside the same callback, after `var pane = new MyPane();`.

See `~/.claude/skills/xrai-excel/templates/pilot-wiring.cs` for copy-pasteable snippets.

## A5. Enforce x:Name hygiene

Every interactive control Claude should target needs `x:Name` in XAML:

```xml
<!-- Before (invisible to XRai) -->
<Button Content="Refresh" Click="OnRefreshClicked" />

<!-- After (discoverable as "RefreshButton") -->
<Button x:Name="RefreshButton" Content="Refresh" Click="OnRefreshClicked" />
```

Find task pane XAML files:
```bash
find . -name "*.xaml" -path "*TaskPane*" -o -name "*.xaml" -path "*Pane*"
```

Unnamed controls get synthetic names like `_unnamed_Button_Refresh_0` but explicit `x:Name` is always preferred.

## A6. Drop project-local CLAUDE.md

```bash
cp ~/.claude/skills/xrai-excel/templates/consumer-CLAUDE.md ./CLAUDE.md
```

If a `CLAUDE.md` already exists, MERGE — append an `## XRai` section.

## A7. Build and verify

```bash
dotnet build
```

Then verify end-to-end:

```bash
cat <<'CMDS' | "$HOME/.claude/skills/xrai-excel/bin/XRai.Tool.exe"
{"cmd":"connect"}
{"cmd":"pane"}
{"cmd":"model"}
CMDS
```

Expected:
- `connect` → `hooks: true` (if false: add-in not loaded or Pilot.Start missing)
- `pane` → non-empty `controls` array (if empty: Pilot.Expose not called)
- `model` → ViewModel properties (if empty: Pilot.ExposeModel not called)

## Quick reference: file locations

| Concern | File |
|---|---|
| PackageReference for XRai.Hooks | `*.csproj` — must use `Version="1.0.0-*"` |
| `Pilot.Start()` / `Pilot.Stop()` | `AddInEntry.cs` (class implementing `IExcelAddIn`) |
| `Pilot.Expose(pane)` | Wherever the task pane UserControl is constructed |
| `Pilot.ExposeModel(vm)` | Same file as `Pilot.Start()`, or DI container setup |
| `x:Name` on controls | All `.xaml` files in the task pane |
| Project CLAUDE.md | Project root (merge with existing) |
