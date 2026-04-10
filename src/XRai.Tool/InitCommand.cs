using System.Text;

namespace XRai.Tool;

/// <summary>
/// Scaffolds a new XRai-enabled Excel-DNA add-in project.
///
/// Default template ("full") generates a production-grade add-in with:
///   - Excel-DNA 1.9.0 + XRai.Hooks wired up
///   - WPF task pane hosted inside a real Custom Task Pane via ElementHost
///     (the canonical CTP + WinForms-host + WPF-child bridge pattern)
///   - Ribbon tab with a "Show Pane" button
///   - MVVM scaffolding (ViewModelBase, RelayCommand)
///   - Dark theme ResourceDictionary (Colors.xaml + Styles.xaml)
///   - Per-workbook CTP lifecycle management
///   - Solution file, .gitignore, and a runtimeconfig copy target
///     that makes ExcelDna .NET 8 loading work out of the box
///
/// Use <c>--template minimal</c> to get the older automation-only single-file
/// scaffold (kept for backward compat).
/// </summary>
public static class InitCommand
{
    public static int Run(string[] args)
    {
        string? projectNameArg = null;
        string template = "full";
        bool inPlaceFlag = false;

        // Skip the leading "init" token and parse remaining args.
        for (int i = 1; i < args.Length; i++)
        {
            var a = args[i];
            if (a == "--template" && i + 1 < args.Length)
            {
                template = args[++i].ToLowerInvariant();
            }
            else if (a.StartsWith("--template="))
            {
                template = a.Substring("--template=".Length).ToLowerInvariant();
            }
            else if (a == "--in-place" || a == "--here")
            {
                inPlaceFlag = true;
            }
            else if (!a.StartsWith("-") && projectNameArg == null)
            {
                projectNameArg = a;
            }
        }

        if (template != "full" && template != "minimal")
        {
            Console.WriteLine($"Error: unknown template '{template}'. Valid values: full, minimal.");
            return 1;
        }

        var cwd = Directory.GetCurrentDirectory();

        // In-place detection: if the user is inside an empty directory and either
        // passed --in-place OR didn't pass a project name, scaffold directly into
        // cwd using the directory's own name as the project name. This is the
        // common "I made a folder and cd'd into it" workflow — creating a nested
        // subdirectory in that case is surprising and produces broken paths.
        var cwdName = new DirectoryInfo(cwd).Name;
        var cwdIsEmpty = !Directory.EnumerateFileSystemEntries(cwd).Any();
        var scaffoldInPlace = inPlaceFlag || (projectNameArg == null && cwdIsEmpty);

        string projectName;
        string outputDir;

        if (scaffoldInPlace)
        {
            if (!cwdIsEmpty)
            {
                Console.WriteLine($"Error: --in-place requires an empty current directory ({cwd}).");
                Console.WriteLine("       Either remove existing files or pass a project name to scaffold into a subdirectory.");
                return 1;
            }
            projectName = projectNameArg ?? cwdName;
            outputDir = cwd;
            Console.WriteLine($"Creating XRai-enabled Excel add-in IN PLACE: {projectName}");
            Console.WriteLine($"Location: {outputDir}");
        }
        else
        {
            projectName = projectNameArg ?? "MyExcelAddin";
            outputDir = Path.Combine(cwd, projectName);

            if (Directory.Exists(outputDir))
            {
                Console.WriteLine($"Error: Directory '{projectName}' already exists.");
                Console.WriteLine($"       Pass --in-place to scaffold into the current directory instead.");
                return 1;
            }
            Console.WriteLine($"Creating XRai-enabled Excel add-in: {projectName}");
            Console.WriteLine($"Location: {outputDir}");
        }

        Console.WriteLine($"Template: {template}");
        Console.WriteLine();

        Directory.CreateDirectory(outputDir);

        if (template == "minimal")
            return ScaffoldMinimal(projectName, outputDir);

        return ScaffoldFull(projectName, outputDir);
    }

    // ─────────────────────────────────────────────────────────────────────────
    //  FULL TEMPLATE — production-grade add-in scaffold
    // ─────────────────────────────────────────────────────────────────────────

    private static int ScaffoldFull(string rawProjectName, string outputDir)
    {
        // Normalize names:
        //   rawProjectName = "My.Great.Addin"    (what the user typed)
        //   projectNs       = "My.Great.Addin"   (root namespace — dots OK)
        //   addinName       = "My.Great.Addin"   (user-facing label, .sln / .dna name)
        //   safeIdent       = "MyGreatAddin"     (bare C# identifier, no dots)
        //
        // The CTP project always ends in ".AddIn" (e.g. My.Great.Addin.AddIn) so
        // XLL names are predictable. The solution groups everything under
        // <rawProjectName>.sln.
        var projectNs = rawProjectName;
        var addinName = rawProjectName;
        var safeIdent = new string(rawProjectName.Where(c => char.IsLetterOrDigit(c) || c == '_').ToArray());
        if (string.IsNullOrEmpty(safeIdent) || char.IsDigit(safeIdent[0]))
            safeIdent = "Addin" + safeIdent;

        var addinProjectName = $"{rawProjectName}.AddIn";
        var addinNs = $"{projectNs}.AddIn";

        // Layout:
        //   <outputDir>/
        //     .gitignore
        //     <rawProjectName>.sln
        //     src/<addinProjectName>/
        //       <addinProjectName>.csproj
        //       <addinProjectName>-AddIn.dna
        //       AddIn/AddInEntry.cs
        //       AddIn/Ribbon.cs
        //       UI/TaskPaneHost.cs
        //       UI/TaskPaneManager.cs
        //       UI/Views/MainPane.xaml
        //       UI/Views/MainPane.xaml.cs
        //       UI/ViewModels/MainViewModel.cs
        //       UI/ViewModels/Base/ViewModelBase.cs
        //       UI/ViewModels/Base/RelayCommand.cs
        //       UI/Themes/Colors.xaml
        //       UI/Themes/Styles.xaml
        var srcDir = Path.Combine(outputDir, "src");
        var projDir = Path.Combine(srcDir, addinProjectName);
        var addinSubDir = Path.Combine(projDir, "AddIn");
        var uiDir = Path.Combine(projDir, "UI");
        var viewsDir = Path.Combine(uiDir, "Views");
        var vmDir = Path.Combine(uiDir, "ViewModels");
        var vmBaseDir = Path.Combine(vmDir, "Base");
        var themesDir = Path.Combine(uiDir, "Themes");

        foreach (var d in new[] { srcDir, projDir, addinSubDir, uiDir, viewsDir, vmDir, vmBaseDir, themesDir })
            Directory.CreateDirectory(d);

        // ── .gitignore ─────────────────────────────────────────────────
        File.WriteAllText(Path.Combine(outputDir, ".gitignore"),
            "bin/\n" +
            "obj/\n" +
            ".vs/\n" +
            "*.user\n" +
            "*.suo\n" +
            "*.userosscache\n" +
            "*.sln.docstates\n" +
            "[Dd]ebug/\n" +
            "[Rr]elease/\n" +
            "x64/\n" +
            "x86/\n" +
            "[Tt]est[Rr]esult*/\n" +
            ".idea/\n");

        // ── <rawProjectName>.sln ───────────────────────────────────────
        File.WriteAllText(Path.Combine(outputDir, $"{rawProjectName}.sln"),
            BuildSolutionFile(rawProjectName, addinProjectName));

        // The .dna file's base name drives the xll output name. We use
        // "<rawProjectName>-AddIn.dna" so the output becomes
        // "<rawProjectName>-AddIn64.xll" — matches the CellVault convention
        // and gives a clean path for the Next Steps instructions.
        var dnaBaseName = $"{rawProjectName}-AddIn";

        // ── <addinProjectName>.csproj ─────────────────────────────────
        File.WriteAllText(Path.Combine(projDir, $"{addinProjectName}.csproj"),
$@"<Project Sdk=""Microsoft.NET.Sdk"">

  <PropertyGroup>
    <TargetFramework>net8.0-windows</TargetFramework>
    <Nullable>enable</Nullable>
    <ImplicitUsings>enable</ImplicitUsings>
    <UseWPF>true</UseWPF>
    <UseWindowsForms>true</UseWindowsForms>
    <RootNamespace>{addinNs}</RootNamespace>
    <AssemblyName>{addinProjectName}</AssemblyName>
    <LangVersion>latest</LangVersion>
    <RollForward>LatestMajor</RollForward>
    <ExcelDnaCreate32BitAddIn>false</ExcelDnaCreate32BitAddIn>
    <ExcelDnaCreate64BitAddIn>true</ExcelDnaCreate64BitAddIn>
  </PropertyGroup>

  <ItemGroup>
    <PackageReference Include=""ExcelDna.AddIn"" Version=""1.9.0"" />
    <PackageReference Include=""ExcelDna.Interop"" Version=""15.0.1"" />
    <PackageReference Include=""XRai.Hooks"" Version=""1.0.0-*"" />
  </ItemGroup>

  <!--
    ExcelDna .NET 8 runtimeconfig fix.
    ExcelDna's host loader looks for <xllname>.runtimeconfig.json alongside the xll.
    .NET SDK emits <assemblyname>.runtimeconfig.json — we copy it to the xll-suffixed name.
  -->
  <Target Name=""CopyRuntimeConfigForExcelDnaHost"" AfterTargets=""Build"">
    <Copy SourceFiles=""$(OutputPath)$(AssemblyName).runtimeconfig.json""
          DestinationFiles=""$(OutputPath){dnaBaseName}64.runtimeconfig.json""
          SkipUnchangedFiles=""true"" />
  </Target>

</Project>
");

        // ── {rawProjectName}-AddIn.dna ────────────────────────────────
        File.WriteAllText(Path.Combine(projDir, $"{dnaBaseName}.dna"),
$@"<?xml version=""1.0"" encoding=""utf-8""?>
<DnaLibrary Name=""{addinName}"" RuntimeVersion=""v8.0""
            xmlns=""http://schemas.excel-dna.net/addin/2020/07/dnalibrary"">
  <ExternalLibrary Path=""{addinProjectName}.dll"" ExplicitExports=""false""
                   LoadFromBytes=""true"" Pack=""true"" IncludePdb=""false"" />
</DnaLibrary>
");

        // ── AddIn/AddInEntry.cs ───────────────────────────────────────
        File.WriteAllText(Path.Combine(addinSubDir, "AddInEntry.cs"),
$@"using ExcelDna.Integration;
using XRai.Hooks;
using XL = Microsoft.Office.Interop.Excel;
using {addinNs}.UI;

namespace {addinNs}.AddIn;

/// <summary>
/// Excel-DNA add-in entry point. Starts XRai.Hooks and creates the task pane
/// on the first workbook that becomes active.
/// </summary>
public class AddInEntry : IExcelAddIn
{{
    private static XL.Application? _app;

    public static XL.Application? Application => _app;

    // Last-resort file logger: writes to %TEMP%\{safeIdent}-startup.log no matter
    // what, even if Pilot.Start() crashed before Pilot.Log was usable. This is the
    // ONLY trace of startup failures that don't reach Pilot.Log.
    private static readonly string _startupLogPath = System.IO.Path.Combine(
        System.IO.Path.GetTempPath(), ""{safeIdent}-startup.log"");

    private static void StartupLog(string message)
    {{
        try
        {{
            System.IO.File.AppendAllText(_startupLogPath,
                $""[{{DateTime.UtcNow:o}}] {{message}}{{System.Environment.NewLine}}"");
        }}
        catch {{ /* never throw from a logger */ }}
    }}

    public void AutoOpen()
    {{
        StartupLog(""AutoOpen entered"");
        try
        {{
            try
            {{
                Pilot.Start();
                StartupLog(""Pilot.Start() OK"");
            }}
            catch (Exception pilotEx)
            {{
                // Pilot.Start failed — log to our file fallback, don't rethrow,
                // the add-in should still load and run.
                StartupLog($""Pilot.Start() FAILED: {{pilotEx}}"");
            }}

            try {{ Pilot.Log(""AutoOpen starting"", ""{safeIdent}""); }} catch {{ }}

            _app = (XL.Application)ExcelDnaUtil.Application;

            // CTPs are per-workbook in Excel — create one whenever a workbook activates.
            _app.WorkbookActivate += OnWorkbookActivate;

            // If Excel started with a workbook already open, create the pane now.
            ExcelAsyncUtil.QueueAsMacro(() =>
            {{
                try
                {{
                    var wb = _app.ActiveWorkbook;
                    if (wb != null)
                        TaskPaneManager.Instance.CreateOrShowPane(wb.FullName);
                }}
                catch (Exception ex)
                {{
                    StartupLog($""Initial CTP create failed: {{ex}}"");
                    try {{ Pilot.Log($""Initial CTP create failed: {{ex.Message}}"", ""{safeIdent}""); }} catch {{ }}
                }}
            }});

            StartupLog(""AutoOpen completed"");
            try {{ Pilot.Log(""AutoOpen completed"", ""{safeIdent}""); }} catch {{ }}
        }}
        catch (Exception ex)
        {{
            StartupLog($""AutoOpen FAILED: {{ex}}"");
            try {{ Pilot.Log($""AutoOpen failed: {{ex}}"", ""{safeIdent}""); }} catch {{ }}
        }}
    }}

    public void AutoClose()
    {{
        try {{ TaskPaneManager.Instance.DisposeAll(); }} catch {{ }}
        try
        {{
            if (_app != null)
                _app.WorkbookActivate -= OnWorkbookActivate;
        }}
        catch {{ }}
        try {{ Pilot.Stop(); }} catch {{ }}
    }}

    private static void OnWorkbookActivate(XL.Workbook wb)
    {{
        try
        {{
            ExcelAsyncUtil.QueueAsMacro(() =>
            {{
                try {{ TaskPaneManager.Instance.CreateOrShowPane(wb.FullName); }}
                catch (Exception ex) {{ Pilot.Log($""WorkbookActivate CTP create failed: {{ex.Message}}"", ""{safeIdent}""); }}
            }});
        }}
        catch {{ }}
    }}
}}
");

        // ── AddIn/Ribbon.cs ───────────────────────────────────────────
        File.WriteAllText(Path.Combine(addinSubDir, "Ribbon.cs"),
$@"using System.Runtime.InteropServices;
using ExcelDna.Integration;
using ExcelDna.Integration.CustomUI;
using {addinNs}.UI;
using XL = Microsoft.Office.Interop.Excel;

namespace {addinNs}.AddIn;

/// <summary>
/// Ribbon tab with a single ""Show Pane"" button that opens (or re-opens)
/// the task pane for the active workbook.
/// </summary>
[ComVisible(true)]
public class Ribbon : ExcelRibbon
{{
    public override string GetCustomUI(string ribbonId) => @""
<customUI xmlns='http://schemas.microsoft.com/office/2009/07/customui'>
  <ribbon>
    <tabs>
      <tab id='tab{safeIdent}' label='{addinName}'>
        <group id='grpMain' label='Main'>
          <button id='btnShowPane'
                  label='Show Pane'
                  size='large'
                  imageMso='TaskPaneControlToggle'
                  onAction='OnShowPanePressed'
                  screentip='Show the {addinName} pane'
                  supertip='Opens (or re-opens) the {addinName} task pane for the current workbook.' />
        </group>
      </tab>
    </tabs>
  </ribbon>
</customUI>"";

    public void OnShowPanePressed(IRibbonControl control)
    {{
        try
        {{
            var app = (XL.Application)ExcelDnaUtil.Application;
            var wb = app.ActiveWorkbook;
            if (wb == null)
            {{
                System.Windows.Forms.MessageBox.Show(
                    ""Open or create a workbook first."",
                    ""{addinName}"",
                    System.Windows.Forms.MessageBoxButtons.OK,
                    System.Windows.Forms.MessageBoxIcon.Information);
                return;
            }}
            TaskPaneManager.Instance.CreateOrShowPane(wb.FullName);
        }}
        catch (Exception ex)
        {{
            System.Windows.Forms.MessageBox.Show(
                $""Failed to open pane: {{ex.Message}}"",
                ""{addinName}"",
                System.Windows.Forms.MessageBoxButtons.OK,
                System.Windows.Forms.MessageBoxIcon.Error);
        }}
    }}
}}
");

        // ── UI/TaskPaneHost.cs ────────────────────────────────────────
        // Unique GUID for this scaffold — can be regenerated by the user.
        var hostGuid = Guid.NewGuid().ToString().ToUpperInvariant();

        File.WriteAllText(Path.Combine(uiDir, "TaskPaneHost.cs"),
$@"using System.Runtime.InteropServices;
using System.Windows.Forms;
using System.Windows.Forms.Integration;
using {addinNs}.UI.Views;

namespace {addinNs}.UI;

/// <summary>
/// Marker interface required by .NET 6+ COM interop. Office CTP hosting requires
/// an IDispatch-compatible default interface. .NET 6+ no longer auto-generates one,
/// so we declare it explicitly.
/// </summary>
public interface ITaskPaneHost {{ }}

/// <summary>
/// WinForms UserControl that bridges the Office Custom Task Pane to WPF via
/// <see cref=""ElementHost""/>. Office requires an ActiveX-compatible control;
/// WPF cannot expose itself as ActiveX, so this thin WinForms shell wraps an
/// ElementHost whose Child is the WPF <see cref=""MainPane""/>.
/// </summary>
[ComVisible(true)]
[ComDefaultInterface(typeof(ITaskPaneHost))]
[Guid(""{hostGuid}"")]
public sealed class TaskPaneHost : UserControl, ITaskPaneHost
{{
    private readonly ElementHost _elementHost;
    private readonly MainPane _mainPane;

    public TaskPaneHost()
    {{
        _mainPane = new MainPane();
        _elementHost = new ElementHost
        {{
            Dock = DockStyle.Fill,
            Child = _mainPane,
        }};
        Controls.Add(_elementHost);
    }}

    /// <summary>Gets the hosted WPF pane.</summary>
    public MainPane MainPane => _mainPane;

    protected override void Dispose(bool disposing)
    {{
        if (disposing) _elementHost.Dispose();
        base.Dispose(disposing);
    }}
}}
");

        // ── UI/TaskPaneManager.cs ─────────────────────────────────────
        File.WriteAllText(Path.Combine(uiDir, "TaskPaneManager.cs"),
$@"using ExcelDna.Integration.CustomUI;
using XRai.Hooks;

namespace {addinNs}.UI;

/// <summary>
/// Manages Custom Task Pane (CTP) lifecycle. One CTP per workbook, keyed by
/// <c>Workbook.FullName</c>. Exposes the pane and its ViewModel to XRai.Hooks
/// once the WPF visual tree is fully realized.
/// </summary>
public sealed class TaskPaneManager
{{
    private static readonly Lazy<TaskPaneManager> _instance = new(() => new TaskPaneManager());
    public static TaskPaneManager Instance => _instance.Value;

    private readonly Dictionary<string, (CustomTaskPane Ctp, TaskPaneHost Host)> _panes =
        new(StringComparer.OrdinalIgnoreCase);

    private TaskPaneManager() {{ }}

    /// <summary>
    /// Creates a new CTP for the workbook or shows the existing one.
    /// </summary>
    public TaskPaneHost CreateOrShowPane(string workbookKey)
    {{
        if (string.IsNullOrWhiteSpace(workbookKey))
            workbookKey = ""default"";

        if (_panes.TryGetValue(workbookKey, out var existing))
        {{
            existing.Ctp.Visible = true;
            return existing.Host;
        }}

        var host = new TaskPaneHost();
        var ctp = CustomTaskPaneFactory.CreateCustomTaskPane(host, ""{addinName}"");
        ctp.DockPosition = MsoCTPDockPosition.msoCTPDockPositionRight;
        ctp.Width = 420;
        ctp.Visible = true;

        // Log visibility changes (user clicking the CTP X button, Show Pane ribbon, etc.).
        // Wrapped in try/catch because the event surface differs across ExcelDna versions.
        try
        {{
            ctp.VisibleStateChange += _ =>
            {{
                try {{ Pilot.Log($""CTP visible={{ctp.Visible}}"", ""{safeIdent}""); }} catch {{ }}
            }};
        }}
        catch {{ }}

        _panes[workbookKey] = (ctp, host);

        // Expose pane controls + ViewModel to XRai once the visual tree is fully rendered.
        // Calling Pilot.Expose synchronously here (before Loaded fires) walks an incomplete
        // tree — collapsed/unrealized subtrees aren't visible yet and never get registered.
        host.MainPane.Loaded += (_, _) =>
        {{
            try
            {{
                Pilot.Expose(host.MainPane);
                Pilot.ExposeModel(host.MainPane.ViewModel);
                Pilot.Log(""Exposed MainPane and MainViewModel"", ""{safeIdent}"");
            }}
            catch (Exception ex)
            {{
                try {{ Pilot.Log($""Expose failed: {{ex.Message}}"", ""{safeIdent}""); }} catch {{ }}
            }}
        }};

        return host;
    }}

    public bool HasPane(string workbookKey) => _panes.ContainsKey(workbookKey);

    public void TogglePane(string workbookKey)
    {{
        if (_panes.TryGetValue(workbookKey, out var existing))
            existing.Ctp.Visible = !existing.Ctp.Visible;
    }}

    public void DisposePane(string workbookKey)
    {{
        if (!_panes.TryGetValue(workbookKey, out var entry)) return;
        _panes.Remove(workbookKey);
        try {{ entry.Ctp.Visible = false; }} catch {{ }}
        try {{ entry.Host.Dispose(); }} catch {{ }}
        try {{ entry.Ctp.Delete(); }} catch {{ }}
    }}

    public void DisposeAll()
    {{
        foreach (var key in _panes.Keys.ToList())
            DisposePane(key);
    }}
}}
");

        // ── UI/Views/MainPane.xaml ────────────────────────────────────
        // Note: NO Width attribute — UserControl stretches to fill the CTP.
        File.WriteAllText(Path.Combine(viewsDir, "MainPane.xaml"),
$@"<UserControl x:Class=""{addinNs}.UI.Views.MainPane""
             xmlns=""http://schemas.microsoft.com/winfx/2006/xaml/presentation""
             xmlns:x=""http://schemas.microsoft.com/winfx/2006/xaml""
             xmlns:d=""http://schemas.microsoft.com/expression/blend/2008""
             xmlns:mc=""http://schemas.openxmlformats.org/markup-compatibility/2006""
             mc:Ignorable=""d""
             d:DesignWidth=""420"" d:DesignHeight=""600""
             Background=""{{DynamicResource BackgroundBrush}}"">
    <UserControl.Resources>
        <ResourceDictionary>
            <ResourceDictionary.MergedDictionaries>
                <ResourceDictionary Source=""/{addinProjectName};component/UI/Themes/Colors.xaml"" />
                <ResourceDictionary Source=""/{addinProjectName};component/UI/Themes/Styles.xaml"" />
            </ResourceDictionary.MergedDictionaries>
        </ResourceDictionary>
    </UserControl.Resources>
    <Grid Margin=""16"">
        <StackPanel>
            <TextBlock Text=""{addinName}""
                       FontSize=""20"" FontWeight=""Bold""
                       Foreground=""{{DynamicResource TextBrush}}"" />

            <TextBlock Text=""Welcome to your new XRai-enabled add-in.""
                       Margin=""0,4,0,0""
                       Foreground=""{{DynamicResource SubtleTextBrush}}"" />

            <TextBlock x:Name=""StatusLabel""
                       Text=""{{Binding Status}}""
                       Margin=""0,20,0,0""
                       Foreground=""{{DynamicResource TextBrush}}"" />

            <Button x:Name=""DemoButton""
                    Content=""Click Me""
                    Margin=""0,12,0,0""
                    Command=""{{Binding DemoCommand}}"" />
        </StackPanel>
    </Grid>
</UserControl>
");

        // ── UI/Views/MainPane.xaml.cs ─────────────────────────────────
        // Note: fully qualify UserControl because UseWindowsForms=true brings
        // System.Windows.Forms.UserControl into scope — without the qualifier
        // `UserControl` is ambiguous.
        File.WriteAllText(Path.Combine(viewsDir, "MainPane.xaml.cs"),
$@"using {addinNs}.UI.ViewModels;

namespace {addinNs}.UI.Views;

public partial class MainPane : System.Windows.Controls.UserControl
{{
    public MainViewModel ViewModel {{ get; }}

    public MainPane()
    {{
        InitializeComponent();
        ViewModel = new MainViewModel();
        DataContext = ViewModel;
    }}
}}
");

        // ── UI/ViewModels/MainViewModel.cs ────────────────────────────
        File.WriteAllText(Path.Combine(vmDir, "MainViewModel.cs"),
$@"using System.Windows.Input;
using {addinNs}.UI.ViewModels.Base;

namespace {addinNs}.UI.ViewModels;

public class MainViewModel : ViewModelBase
{{
    private string _status = ""Ready."";
    public string Status
    {{
        get => _status;
        set => SetField(ref _status, value);
    }}

    private int _clickCount;
    public int ClickCount
    {{
        get => _clickCount;
        set => SetField(ref _clickCount, value);
    }}

    public ICommand DemoCommand {{ get; }}

    public MainViewModel()
    {{
        DemoCommand = new RelayCommand(_ =>
        {{
            ClickCount++;
            Status = $""Clicked {{ClickCount}} time(s)."";
        }});
    }}
}}
");

        // ── UI/ViewModels/Base/ViewModelBase.cs ──────────────────────
        File.WriteAllText(Path.Combine(vmBaseDir, "ViewModelBase.cs"),
$@"using System.ComponentModel;
using System.Runtime.CompilerServices;

namespace {addinNs}.UI.ViewModels.Base;

/// <summary>Minimal INPC base class for view models.</summary>
public abstract class ViewModelBase : INotifyPropertyChanged
{{
    public event PropertyChangedEventHandler? PropertyChanged;

    protected void OnPropertyChanged([CallerMemberName] string? propertyName = null)
        => PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));

    protected bool SetField<T>(ref T field, T value, [CallerMemberName] string? propertyName = null)
    {{
        if (EqualityComparer<T>.Default.Equals(field, value)) return false;
        field = value;
        OnPropertyChanged(propertyName);
        return true;
    }}
}}
");

        // ── UI/ViewModels/Base/RelayCommand.cs ────────────────────────
        File.WriteAllText(Path.Combine(vmBaseDir, "RelayCommand.cs"),
$@"using System.Windows.Input;

namespace {addinNs}.UI.ViewModels.Base;

/// <summary>Simple ICommand implementation for MVVM bindings.</summary>
public sealed class RelayCommand : ICommand
{{
    private readonly Action<object?> _execute;
    private readonly Predicate<object?>? _canExecute;

    public RelayCommand(Action<object?> execute, Predicate<object?>? canExecute = null)
    {{
        _execute = execute ?? throw new ArgumentNullException(nameof(execute));
        _canExecute = canExecute;
    }}

    public bool CanExecute(object? parameter) => _canExecute?.Invoke(parameter) ?? true;
    public void Execute(object? parameter) => _execute(parameter);

    public event EventHandler? CanExecuteChanged
    {{
        add    {{ CommandManager.RequerySuggested += value; }}
        remove {{ CommandManager.RequerySuggested -= value; }}
    }}
}}
");

        // ── UI/Themes/Colors.xaml ─────────────────────────────────────
        File.WriteAllText(Path.Combine(themesDir, "Colors.xaml"),
@"<ResourceDictionary xmlns=""http://schemas.microsoft.com/winfx/2006/xaml/presentation""
                    xmlns:x=""http://schemas.microsoft.com/winfx/2006/xaml"">

    <!-- Neutral dark palette -->
    <SolidColorBrush x:Key=""BackgroundBrush""    Color=""#1A1F2E"" />
    <SolidColorBrush x:Key=""SurfaceBrush""       Color=""#252A3A"" />
    <SolidColorBrush x:Key=""SurfaceHoverBrush""  Color=""#2E3549"" />
    <SolidColorBrush x:Key=""BorderBrush""        Color=""#2E3549"" />

    <SolidColorBrush x:Key=""AccentBrush""        Color=""#4A90E2"" />
    <SolidColorBrush x:Key=""AccentHoverBrush""   Color=""#5AA3F0"" />

    <SolidColorBrush x:Key=""TextBrush""          Color=""#E0E0E0"" />
    <SolidColorBrush x:Key=""SubtleTextBrush""    Color=""#8A8F9E"" />

</ResourceDictionary>
");

        // ── UI/Themes/Styles.xaml ─────────────────────────────────────
        File.WriteAllText(Path.Combine(themesDir, "Styles.xaml"),
@"<ResourceDictionary xmlns=""http://schemas.microsoft.com/winfx/2006/xaml/presentation""
                    xmlns:x=""http://schemas.microsoft.com/winfx/2006/xaml"">

    <ResourceDictionary.MergedDictionaries>
        <ResourceDictionary Source=""Colors.xaml"" />
    </ResourceDictionary.MergedDictionaries>

    <!-- TextBlock default -->
    <Style TargetType=""TextBlock"">
        <Setter Property=""FontFamily"" Value=""Segoe UI"" />
        <Setter Property=""FontSize""   Value=""12"" />
        <Setter Property=""Foreground"" Value=""{StaticResource TextBrush}"" />
    </Style>

    <!-- Button: rounded, accent hover -->
    <Style TargetType=""Button"">
        <Setter Property=""FontFamily"" Value=""Segoe UI"" />
        <Setter Property=""FontSize""   Value=""12"" />
        <Setter Property=""Foreground"" Value=""{StaticResource TextBrush}"" />
        <Setter Property=""Background"" Value=""{StaticResource AccentBrush}"" />
        <Setter Property=""BorderBrush"" Value=""{StaticResource AccentBrush}"" />
        <Setter Property=""BorderThickness"" Value=""1"" />
        <Setter Property=""Padding"" Value=""12,6"" />
        <Setter Property=""Cursor"" Value=""Hand"" />
        <Setter Property=""Template"">
            <Setter.Value>
                <ControlTemplate TargetType=""Button"">
                    <Border x:Name=""Bd""
                            Background=""{TemplateBinding Background}""
                            BorderBrush=""{TemplateBinding BorderBrush}""
                            BorderThickness=""{TemplateBinding BorderThickness}""
                            CornerRadius=""4"">
                        <ContentPresenter HorizontalAlignment=""Center""
                                          VerticalAlignment=""Center""
                                          Margin=""{TemplateBinding Padding}"" />
                    </Border>
                    <ControlTemplate.Triggers>
                        <Trigger Property=""IsMouseOver"" Value=""True"">
                            <Setter TargetName=""Bd"" Property=""Background"" Value=""{StaticResource AccentHoverBrush}"" />
                            <Setter TargetName=""Bd"" Property=""BorderBrush"" Value=""{StaticResource AccentHoverBrush}"" />
                        </Trigger>
                        <Trigger Property=""IsEnabled"" Value=""False"">
                            <Setter TargetName=""Bd"" Property=""Opacity"" Value=""0.5"" />
                        </Trigger>
                    </ControlTemplate.Triggers>
                </ControlTemplate>
            </Setter.Value>
        </Setter>
    </Style>

    <!-- TextBox: dark surface -->
    <Style TargetType=""TextBox"">
        <Setter Property=""FontFamily"" Value=""Segoe UI"" />
        <Setter Property=""FontSize"" Value=""12"" />
        <Setter Property=""Foreground"" Value=""{StaticResource TextBrush}"" />
        <Setter Property=""Background"" Value=""{StaticResource SurfaceBrush}"" />
        <Setter Property=""BorderBrush"" Value=""{StaticResource BorderBrush}"" />
        <Setter Property=""BorderThickness"" Value=""1"" />
        <Setter Property=""Padding"" Value=""6,4"" />
        <Setter Property=""CaretBrush"" Value=""{StaticResource TextBrush}"" />
    </Style>

    <!-- DataGrid: dark -->
    <Style TargetType=""DataGrid"">
        <Setter Property=""Background"" Value=""{StaticResource BackgroundBrush}"" />
        <Setter Property=""Foreground"" Value=""{StaticResource TextBrush}"" />
        <Setter Property=""BorderBrush"" Value=""{StaticResource BorderBrush}"" />
        <Setter Property=""RowBackground"" Value=""{StaticResource BackgroundBrush}"" />
        <Setter Property=""AlternatingRowBackground"" Value=""{StaticResource SurfaceBrush}"" />
        <Setter Property=""GridLinesVisibility"" Value=""Horizontal"" />
        <Setter Property=""HorizontalGridLinesBrush"" Value=""{StaticResource BorderBrush}"" />
        <Setter Property=""HeadersVisibility"" Value=""Column"" />
    </Style>

</ResourceDictionary>
");

        // ── Post-init summary ─────────────────────────────────────────
        Console.WriteLine("  Created .gitignore");
        Console.WriteLine($"  Created {rawProjectName}.sln");
        Console.WriteLine($"  Created src/{addinProjectName}/{addinProjectName}.csproj");
        Console.WriteLine($"  Created src/{addinProjectName}/{dnaBaseName}.dna");
        Console.WriteLine($"  Created src/{addinProjectName}/AddIn/AddInEntry.cs");
        Console.WriteLine($"  Created src/{addinProjectName}/AddIn/Ribbon.cs");
        Console.WriteLine($"  Created src/{addinProjectName}/UI/TaskPaneHost.cs");
        Console.WriteLine($"  Created src/{addinProjectName}/UI/TaskPaneManager.cs");
        Console.WriteLine($"  Created src/{addinProjectName}/UI/Views/MainPane.xaml (+.cs)");
        Console.WriteLine($"  Created src/{addinProjectName}/UI/ViewModels/MainViewModel.cs");
        Console.WriteLine($"  Created src/{addinProjectName}/UI/ViewModels/Base/ViewModelBase.cs");
        Console.WriteLine($"  Created src/{addinProjectName}/UI/ViewModels/Base/RelayCommand.cs");
        Console.WriteLine($"  Created src/{addinProjectName}/UI/Themes/Colors.xaml");
        Console.WriteLine($"  Created src/{addinProjectName}/UI/Themes/Styles.xaml");
        Console.WriteLine();
        Console.WriteLine("Next steps:");
        Console.WriteLine($"  1. cd {rawProjectName}");
        Console.WriteLine($"  2. dotnet build");
        Console.WriteLine($"  3. Load src/{addinProjectName}/bin/Debug/net8.0-windows/{dnaBaseName}64.xll in Excel");
        Console.WriteLine($"     (or the self-contained build at .../publish/{dnaBaseName}64-packed.xll)");
        Console.WriteLine($"  4. Run: echo '{{\"cmd\":\"connect\"}}' | xrai --no-daemon");
        Console.WriteLine();
        return 0;
    }

    private static string BuildSolutionFile(string slnName, string addinProjectName)
    {
        // Minimal but valid Visual Studio solution file targeting VS 2022.
        var projectGuid = "{" + Guid.NewGuid().ToString().ToUpperInvariant() + "}";
        const string csharpSdkGuid = "{9A19103F-16F7-4668-BE54-9A1E7A4F7556}"; // SDK-style C#

        var sb = new StringBuilder();
        sb.AppendLine();
        sb.AppendLine("Microsoft Visual Studio Solution File, Format Version 12.00");
        sb.AppendLine("# Visual Studio Version 17");
        sb.AppendLine("VisualStudioVersion = 17.0.31903.59");
        sb.AppendLine("MinimumVisualStudioVersion = 10.0.40219.1");
        sb.AppendLine(
            $"Project(\"{csharpSdkGuid}\") = \"{addinProjectName}\", " +
            $"\"src\\{addinProjectName}\\{addinProjectName}.csproj\", \"{projectGuid}\"");
        sb.AppendLine("EndProject");
        sb.AppendLine("Global");
        sb.AppendLine("\tGlobalSection(SolutionConfigurationPlatforms) = preSolution");
        sb.AppendLine("\t\tDebug|Any CPU = Debug|Any CPU");
        sb.AppendLine("\t\tRelease|Any CPU = Release|Any CPU");
        sb.AppendLine("\tEndGlobalSection");
        sb.AppendLine("\tGlobalSection(ProjectConfigurationPlatforms) = postSolution");
        sb.AppendLine($"\t\t{projectGuid}.Debug|Any CPU.ActiveCfg = Debug|Any CPU");
        sb.AppendLine($"\t\t{projectGuid}.Debug|Any CPU.Build.0 = Debug|Any CPU");
        sb.AppendLine($"\t\t{projectGuid}.Release|Any CPU.ActiveCfg = Release|Any CPU");
        sb.AppendLine($"\t\t{projectGuid}.Release|Any CPU.Build.0 = Release|Any CPU");
        sb.AppendLine("\tEndGlobalSection");
        sb.AppendLine("\tGlobalSection(SolutionProperties) = preSolution");
        sb.AppendLine("\t\tHideSolutionNode = FALSE");
        sb.AppendLine("\tEndGlobalSection");
        sb.AppendLine("EndGlobal");
        return sb.ToString();
    }

    // ─────────────────────────────────────────────────────────────────────────
    //  MINIMAL TEMPLATE — legacy single-file automation-only scaffold
    // ─────────────────────────────────────────────────────────────────────────

    private static int ScaffoldMinimal(string projectName, string outputDir)
    {
        File.WriteAllText(Path.Combine(outputDir, $"{projectName}.csproj"),
$@"<Project Sdk=""Microsoft.NET.Sdk"">

  <PropertyGroup>
    <TargetFramework>net8.0-windows</TargetFramework>
    <UseWPF>true</UseWPF>
    <Nullable>enable</Nullable>
    <ImplicitUsings>enable</ImplicitUsings>
  </PropertyGroup>

  <ItemGroup>
    <PackageReference Include=""ExcelDna.AddIn"" Version=""1.9.0"" />
    <PackageReference Include=""XRai.Hooks"" Version=""1.0.0-*"" />
  </ItemGroup>

</Project>
");

        File.WriteAllText(Path.Combine(outputDir, "AddInEntry.cs"),
$@"using ExcelDna.Integration;
using XRai.Hooks;

namespace {projectName};

public class AddInEntry : IExcelAddIn
{{
    private static MainPane? _pane;

    public void AutoOpen()
    {{
        Pilot.Start();

        ExcelAsyncUtil.QueueAsMacro(() =>
        {{
            _pane = new MainPane();
            Pilot.Expose(_pane);
            Pilot.ExposeModel(_pane.ViewModel);
            Pilot.Log(""Add-in loaded with XRai hooks"");
        }});
    }}

    public void AutoClose() => Pilot.Stop();
}}
");

        File.WriteAllText(Path.Combine(outputDir, "MainPane.xaml"),
$@"<UserControl x:Class=""{projectName}.MainPane""
             xmlns=""http://schemas.microsoft.com/winfx/2006/xaml/presentation""
             xmlns:x=""http://schemas.microsoft.com/winfx/2006/xaml"">
    <StackPanel Margin=""10"">
        <TextBlock Text=""My Add-in"" FontSize=""18"" FontWeight=""Bold"" Margin=""0,0,0,10""/>
        <TextBox x:Name=""InputBox"" Text=""{{Binding InputText, UpdateSourceTrigger=PropertyChanged}}"" Margin=""0,0,0,8""/>
        <Button x:Name=""GoButton"" Content=""Go"" Click=""GoButton_Click""/>
        <Label x:Name=""ResultLabel"" Content=""{{Binding ResultText}}"" Margin=""0,8,0,0""/>
    </StackPanel>
</UserControl>
");

        File.WriteAllText(Path.Combine(outputDir, "MainPane.xaml.cs"),
$@"using System.Windows;
using System.Windows.Controls;

namespace {projectName};

public partial class MainPane : UserControl
{{
    public MainViewModel ViewModel {{ get; }}

    public MainPane()
    {{
        ViewModel = new MainViewModel();
        DataContext = ViewModel;
        InitializeComponent();
    }}

    private void GoButton_Click(object sender, RoutedEventArgs e)
    {{
        ViewModel.ResultText = $""Processed: {{ViewModel.InputText}}"";
    }}
}}
");

        File.WriteAllText(Path.Combine(outputDir, "MainViewModel.cs"),
$@"using System.ComponentModel;
using System.Runtime.CompilerServices;

namespace {projectName};

public class MainViewModel : INotifyPropertyChanged
{{
    private string _inputText = """";
    private string _resultText = ""Ready"";

    public string InputText {{ get => _inputText; set => Set(ref _inputText, value); }}
    public string ResultText {{ get => _resultText; set => Set(ref _resultText, value); }}

    public event PropertyChangedEventHandler? PropertyChanged;
    private void Set<T>(ref T field, T value, [CallerMemberName] string? name = null)
    {{
        field = value;
        PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(name));
    }}
}}
");

        Console.WriteLine($"  Created {projectName}.csproj");
        Console.WriteLine($"  Created AddInEntry.cs");
        Console.WriteLine($"  Created MainPane.xaml + MainPane.xaml.cs");
        Console.WriteLine($"  Created MainViewModel.cs");
        Console.WriteLine();
        Console.WriteLine("Next steps:");
        Console.WriteLine($"  1. cd {projectName}");
        Console.WriteLine($"  2. dotnet build");
        Console.WriteLine($"  3. Load the .xll in Excel");
        Console.WriteLine($"  4. Run: echo '{{\"cmd\":\"connect\"}}' | xrai --no-daemon");
        Console.WriteLine();
        return 0;
    }
}
