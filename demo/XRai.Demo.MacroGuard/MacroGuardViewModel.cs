using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Diagnostics;
using System.Runtime.CompilerServices;
using System.Text.RegularExpressions;
using System.Windows;
using System.Windows.Input;
using XRai.Demo.MacroGuard.Models;
using XRai.Demo.MacroGuard.Scanner;
using XRai.Demo.MacroGuard.Services;
using Clipboard = System.Windows.Clipboard;
using MessageBox = System.Windows.MessageBox;
using Timer = System.Timers.Timer;

namespace XRai.Demo.MacroGuard;

public class MacroGuardViewModel : INotifyPropertyChanged
{
    // ── Existing backing fields ─────────────────────────────────
    private string _selectedModule = "";
    private string _selectedMacro = "";
    private string _searchText = "";
    private int _macroCount;
    private int _moduleCount;
    private bool _isScanning;
    private double _scanProgress;
    private int _issueCount;
    private int _criticalCount;
    private int _warningCount;
    private int _infoCount;
    private string _apiKey = "";
    private bool _isConnected;
    private bool _isAiWorking;
    private int _tokensUsed;
    private string _selectedPrompt = "Explain this code";
    private string _customPrompt = "";
    private double _temperature = 0.3;
    private int _maxTokens = 4096;
    private int _maxLinesToScan = 10000;
    private bool _autoScan;
    private bool _scanOnSave;
    private bool _strictMode;
    private bool _isDeveloperMode;
    private string _statusMessage = "Ready";
    private string _selectedSeverity = "All";
    private string _selectedTheme = "Dark";
    private string _selectedModuleType = "All";
    private string _selectedAiModel = "claude-sonnet-4-6";
    private string _codePreviewText = "";
    private string _codeInputText = "";
    private string _responseText = "";
    private string _fixDetail = "";
    private int _apiTimeout = 30;
    private int _maxResponseTokens = 4096;
    private bool _ignoreComments;
    private bool _vbaAccessEnabled;

    // ── Premium feature backing fields ──────────────────────────
    // Performance Profiler
    private bool _isProfilingEnabled;
    // Snippets
    private VbaSnippet? _selectedSnippet;
    // Scheduler
    private string _triggerType = "Manual";
    private int _intervalSeconds = 60;
    private string _scheduledMacroName = "";
    private bool _isScheduleEnabled = true;
    // Diff / Version
    private VbaSnapshot? _selectedSnapshot;
    private string _diffText = "";
    // Code Metrics
    private int _totalLinesOfCode;
    private double _overallCommentRatio;
    private double _overallAvgProcLength;
    private int _overallComplexity;
    private int _deadProcedureCount;
    // Quick-Fix
    private QuickFix? _selectedQuickFix;
    private string _quickFixPreview = "";
    // Export
    private string _exportPath = "";

    // Timer tracking for scheduler
    private readonly Dictionary<string, Timer> _activeTimers = new();

    // ── Existing scalar properties ──────────────────────────────
    public string SelectedModule { get => _selectedModule; set => Set(ref _selectedModule, value); }
    public string SelectedMacro { get => _selectedMacro; set => Set(ref _selectedMacro, value); }
    public string SearchText { get => _searchText; set { Set(ref _searchText, value); FilterModules(); } }
    public int MacroCount { get => _macroCount; set => Set(ref _macroCount, value); }
    public int ModuleCount { get => _moduleCount; set => Set(ref _moduleCount, value); }
    public bool IsScanning { get => _isScanning; set => Set(ref _isScanning, value); }
    public double ScanProgress { get => _scanProgress; set => Set(ref _scanProgress, value); }
    public int IssueCount { get => _issueCount; set => Set(ref _issueCount, value); }
    public int CriticalCount { get => _criticalCount; set => Set(ref _criticalCount, value); }
    public int WarningCount { get => _warningCount; set => Set(ref _warningCount, value); }
    public int InfoCount { get => _infoCount; set => Set(ref _infoCount, value); }
    public string ApiKey { get => _apiKey; set { Set(ref _apiKey, value); IsConnected = !string.IsNullOrWhiteSpace(value); } }
    public bool IsConnected { get => _isConnected; set => Set(ref _isConnected, value); }
    public bool IsAiWorking { get => _isAiWorking; set => Set(ref _isAiWorking, value); }
    public int TokensUsed { get => _tokensUsed; set => Set(ref _tokensUsed, value); }
    public string SelectedPrompt { get => _selectedPrompt; set => Set(ref _selectedPrompt, value); }
    public string CustomPrompt { get => _customPrompt; set => Set(ref _customPrompt, value); }
    public double Temperature { get => _temperature; set => Set(ref _temperature, value); }
    public int MaxTokens { get => _maxTokens; set => Set(ref _maxTokens, value); }
    public int MaxLinesToScan { get => _maxLinesToScan; set => Set(ref _maxLinesToScan, value); }
    public bool AutoScan { get => _autoScan; set => Set(ref _autoScan, value); }
    public bool ScanOnSave { get => _scanOnSave; set => Set(ref _scanOnSave, value); }
    public bool StrictMode { get => _strictMode; set => Set(ref _strictMode, value); }
    public bool IsDeveloperMode { get => _isDeveloperMode; set => Set(ref _isDeveloperMode, value); }
    public string StatusMessage { get => _statusMessage; set => Set(ref _statusMessage, value); }
    public string SelectedSeverity { get => _selectedSeverity; set { Set(ref _selectedSeverity, value); FilterIssues(); } }
    public string SelectedTheme { get => _selectedTheme; set => Set(ref _selectedTheme, value); }
    public string SelectedModuleType { get => _selectedModuleType; set { Set(ref _selectedModuleType, value); FilterModules(); } }
    public string SelectedAiModel { get => _selectedAiModel; set => Set(ref _selectedAiModel, value); }
    public string CodePreviewText { get => _codePreviewText; set => Set(ref _codePreviewText, value); }
    public string CodeInputText { get => _codeInputText; set => Set(ref _codeInputText, value); }
    public string ResponseText { get => _responseText; set => Set(ref _responseText, value); }
    public string FixDetail { get => _fixDetail; set => Set(ref _fixDetail, value); }
    public int ApiTimeout { get => _apiTimeout; set => Set(ref _apiTimeout, value); }
    public int MaxResponseTokens { get => _maxResponseTokens; set => Set(ref _maxResponseTokens, value); }
    public bool IgnoreComments { get => _ignoreComments; set => Set(ref _ignoreComments, value); }
    public bool VbaAccessEnabled { get => _vbaAccessEnabled; set => Set(ref _vbaAccessEnabled, value); }

    // ── Premium scalar properties ───────────────────────────────
    public bool IsProfilingEnabled { get => _isProfilingEnabled; set => Set(ref _isProfilingEnabled, value); }
    public VbaSnippet? SelectedSnippet { get => _selectedSnippet; set => Set(ref _selectedSnippet, value); }
    public string TriggerType { get => _triggerType; set => Set(ref _triggerType, value); }
    public int IntervalSeconds { get => _intervalSeconds; set => Set(ref _intervalSeconds, value); }
    public string ScheduledMacroName { get => _scheduledMacroName; set => Set(ref _scheduledMacroName, value); }
    public bool IsScheduleEnabled { get => _isScheduleEnabled; set => Set(ref _isScheduleEnabled, value); }
    public VbaSnapshot? SelectedSnapshot { get => _selectedSnapshot; set => Set(ref _selectedSnapshot, value); }
    public string DiffText { get => _diffText; set => Set(ref _diffText, value); }
    public int TotalLinesOfCode { get => _totalLinesOfCode; set => Set(ref _totalLinesOfCode, value); }
    public double OverallCommentRatio { get => _overallCommentRatio; set => Set(ref _overallCommentRatio, value); }
    public double OverallAvgProcLength { get => _overallAvgProcLength; set => Set(ref _overallAvgProcLength, value); }
    public int OverallComplexity { get => _overallComplexity; set => Set(ref _overallComplexity, value); }
    public int DeadProcedureCount { get => _deadProcedureCount; set => Set(ref _deadProcedureCount, value); }
    public QuickFix? SelectedQuickFix { get => _selectedQuickFix; set { Set(ref _selectedQuickFix, value); QuickFixPreview = value?.PreviewCode ?? ""; } }
    public string QuickFixPreview { get => _quickFixPreview; set => Set(ref _quickFixPreview, value); }
    public string ExportPath { get => _exportPath; set => Set(ref _exportPath, value); }

    // ── Existing collections ────────────────────────────────────
    public ObservableCollection<VbaModuleInfo> Modules { get; } = new();
    public ObservableCollection<VbaModuleInfo> FilteredModules { get; } = new();
    public ObservableCollection<VbaIssue> Issues { get; } = new();
    public ObservableCollection<VbaIssue> FilteredIssues { get; } = new();
    public ObservableCollection<ActionLogEntry> RecentActions { get; } = new();
    public ObservableCollection<string> Prompts { get; } = new()
    {
        "Explain this code",
        "Refactor for clarity",
        "Add error handling",
        "Convert to modern patterns",
        "Find bugs",
        "Write unit test",
        "Generate macro from description"
    };
    public ObservableCollection<string> Severities { get; } = new() { "All", "Error", "Warning", "Info" };
    public ObservableCollection<string> Themes { get; } = new() { "Dark", "Light" };
    public ObservableCollection<string> ModuleTypes { get; } = new()
    {
        "All", "Standard", "Class", "UserForm", "ThisWorkbook", "Sheet"
    };
    public ObservableCollection<string> AiModels { get; } = new()
    {
        "claude-sonnet-4-6", "claude-haiku-4-5"
    };

    // ── Premium collections ─────────────────────────────────────
    public ObservableCollection<PerformanceRecord> PerformanceData { get; } = new();
    public ObservableCollection<VbaSnippet> Snippets { get; }
    public ObservableCollection<ScheduleEntry> Schedules { get; } = new();
    public ObservableCollection<string> TriggerTypes { get; } = new() { "On Timer", "On Workbook Open", "On Sheet Change", "On Save", "Manual" };
    public ObservableCollection<VbaSnapshot> Snapshots { get; } = new();
    public ObservableCollection<ModuleMetrics> MetricsData { get; } = new();
    public ObservableCollection<DependencyInfo> Dependencies { get; } = new();
    public ObservableCollection<QuickFix> QuickFixes { get; } = new();

    // ── Existing commands ───────────────────────────────────────
    public ICommand RefreshCommand { get; }
    public ICommand RunMacroCommand { get; }
    public ICommand ScanCommand { get; }
    public ICommand NavigateCommand { get; }
    public ICommand AutoFixCommand { get; }
    public ICommand AskAiCommand { get; }
    public ICommand CopyResponseCommand { get; }
    public ICommand InsertCodeCommand { get; }
    public ICommand NewModuleCommand { get; }
    public ICommand DeleteModuleCommand { get; }
    public ICommand ExportAllCommand { get; }
    public ICommand ImportCommand { get; }
    public ICommand BackupCommand { get; }
    public ICommand ResetCommand { get; }

    // ── Premium commands ────────────────────────────────────────
    public ICommand InsertSnippetCommand { get; }
    public ICommand AddScheduleCommand { get; }
    public ICommand TakeSnapshotCommand { get; }
    public ICommand CompareSnapshotCommand { get; }
    public ICommand RefreshMetricsCommand { get; }
    public ICommand RefreshDependenciesCommand { get; }
    public ICommand ApplyQuickFixCommand { get; }
    public ICommand ExportBasCommand { get; }
    public ICommand ExportHtmlDocCommand { get; }
    public ICommand ExportHealthCsvCommand { get; }
    public ICommand ExportAllTextCommand { get; }

    private readonly VbaAnalyzer _analyzer = new();
    private readonly AiAssistant _aiAssistant = new();

    public MacroGuardViewModel()
    {
        // Existing commands
        RefreshCommand = new RelayCommand(ExecuteRefresh);
        RunMacroCommand = new RelayCommand(ExecuteRunMacro, () => !string.IsNullOrEmpty(SelectedMacro));
        ScanCommand = new RelayCommand(ExecuteScan, () => !IsScanning);
        NavigateCommand = new RelayCommand(ExecuteNavigate);
        AutoFixCommand = new RelayCommand(ExecuteAutoFix, () => Issues.Count > 0);
        AskAiCommand = new RelayCommand(ExecuteAskAi, () => !IsAiWorking && IsConnected);
        CopyResponseCommand = new RelayCommand(ExecuteCopyResponse, () => !string.IsNullOrEmpty(ResponseText));
        InsertCodeCommand = new RelayCommand(ExecuteInsertCode, () => !string.IsNullOrEmpty(ResponseText));
        NewModuleCommand = new RelayCommand(ExecuteNewModule);
        DeleteModuleCommand = new RelayCommand(ExecuteDeleteModule, () => !string.IsNullOrEmpty(SelectedModule));
        ExportAllCommand = new RelayCommand(ExecuteExportAll, () => Modules.Count > 0);
        ImportCommand = new RelayCommand(ExecuteImport);
        BackupCommand = new RelayCommand(ExecuteBackup, () => Modules.Count > 0);
        ResetCommand = new RelayCommand(ExecuteReset);

        // Premium commands
        InsertSnippetCommand = new RelayCommand(ExecuteInsertSnippet);
        AddScheduleCommand = new RelayCommand(ExecuteAddSchedule);
        TakeSnapshotCommand = new RelayCommand(ExecuteTakeSnapshot);
        CompareSnapshotCommand = new RelayCommand(ExecuteCompareSnapshot);
        RefreshMetricsCommand = new RelayCommand(ExecuteRefreshMetrics);
        RefreshDependenciesCommand = new RelayCommand(ExecuteRefreshDependencies);
        ApplyQuickFixCommand = new RelayCommand(ExecuteApplyQuickFix);
        ExportBasCommand = new RelayCommand(ExecuteExportBas);
        ExportHtmlDocCommand = new RelayCommand(ExecuteExportHtmlDoc);
        ExportHealthCsvCommand = new RelayCommand(ExecuteExportHealthCsv);
        ExportAllTextCommand = new RelayCommand(ExecuteExportAllText);

        // Initialize snippets
        Snippets = new ObservableCollection<VbaSnippet>(SnippetLibrary.GetAllSnippets());

        LoadSampleData();
    }

    // ══════════════════════════════════════════════════════════════
    //  EXISTING METHODS
    // ══════════════════════════════════════════════════════════════

    private void LoadSampleData()
    {
        var mod1 = new VbaModuleInfo
        {
            Name = "Module1",
            Type = "Standard",
            LineCount = 45,
            MacroNames = new List<string> { "CalculateTotal", "FormatReport" },
            Code = "Option Explicit\n\nPublic Sub CalculateTotal()\n    Dim total As Double\n    Dim i As Integer\n    For i = 1 To 100\n        total = total + Cells(i, 1).Value\n    Next i\n    Range(\"B1\").Value = total\nEnd Sub\n\nPublic Sub FormatReport()\n    ' Format the active sheet\n    Dim ws As Worksheet\n    Set ws = ActiveSheet\n    ws.Range(\"A1:Z1\").Font.Bold = True\n    ws.Range(\"A1:Z1\").Interior.Color = RGB(0, 112, 192)\nEnd Sub"
        };

        var mod2 = new VbaModuleInfo
        {
            Name = "Module2",
            Type = "Standard",
            LineCount = 30,
            MacroNames = new List<string> { "ImportData", "CleanData" },
            Code = "Sub ImportData()\n    Dim filePath\n    filePath = \"C:\\Data\\input.csv\"\n    Workbooks.Open filePath\nEnd Sub\n\nSub CleanData()\n    GoTo StartClean\nStartClean:\n    ' Clean the data\nEnd Sub"
        };

        var mod3 = new VbaModuleInfo
        {
            Name = "ThisWorkbook",
            Type = "ThisWorkbook",
            LineCount = 12,
            MacroNames = new List<string> { "Workbook_Open", "Workbook_BeforeClose" },
            Code = "Private Sub Workbook_Open()\n    MsgBox \"Welcome!\"\nEnd Sub\n\nPrivate Sub Workbook_BeforeClose(Cancel As Boolean)\n    ' Save settings\nEnd Sub"
        };

        var mod4 = new VbaModuleInfo
        {
            Name = "clsLogger",
            Type = "Class",
            LineCount = 55,
            MacroNames = new List<string> { "LogMessage", "GetLog", "ClearLog" },
            Code = "Option Explicit\n\nPrivate mLog As String\n\nPublic Sub LogMessage(msg As String)\n    On Error GoTo ErrHandler\n    mLog = mLog & vbCrLf & Now & \": \" & msg\n    Exit Sub\nErrHandler:\n    Debug.Print \"Error logging: \" & Err.Description\nEnd Sub\n\nPublic Function GetLog() As String\n    GetLog = mLog\nEnd Function\n\nPublic Sub ClearLog()\nEnd Sub"
        };

        Modules.Add(mod1);
        Modules.Add(mod2);
        Modules.Add(mod3);
        Modules.Add(mod4);

        FilterModules();
        UpdateCounts();
    }

    private void UpdateCounts()
    {
        ModuleCount = Modules.Count;
        MacroCount = Modules.Sum(m => m.MacroNames.Count);
    }

    private void FilterModules()
    {
        FilteredModules.Clear();
        foreach (var mod in Modules)
        {
            if (SelectedModuleType != "All" && mod.Type != SelectedModuleType) continue;
            if (!string.IsNullOrEmpty(SearchText) &&
                !mod.Name.Contains(SearchText, StringComparison.OrdinalIgnoreCase) &&
                !mod.MacroNames.Any(m => m.Contains(SearchText, StringComparison.OrdinalIgnoreCase)))
                continue;
            FilteredModules.Add(mod);
        }
    }

    private void FilterIssues()
    {
        FilteredIssues.Clear();
        foreach (var issue in Issues)
        {
            if (SelectedSeverity != "All" && issue.Severity != SelectedSeverity) continue;
            FilteredIssues.Add(issue);
        }
    }

    private void LogAction(string action, string details)
    {
        RecentActions.Insert(0, new ActionLogEntry { Action = action, Details = details });
        StatusMessage = $"{action}: {details}";
    }

    // ── Existing command implementations ────────────────────────
    private void ExecuteRefresh()
    {
        FilterModules();
        UpdateCounts();
        LogAction("Refresh", $"Found {ModuleCount} modules, {MacroCount} macros");
    }

    private void ExecuteRunMacro()
    {
        if (string.IsNullOrEmpty(SelectedMacro)) return;

        Action runAction = () =>
        {
            dynamic app = ExcelDna.Integration.ExcelDnaUtil.Application;
            app.Run(SelectedMacro);
        };

        try
        {
            if (IsProfilingEnabled)
            {
                var elapsed = ProfileMacroExecution(SelectedMacro, runAction);
                LogAction("Run Macro (Profiled)", $"{SelectedMacro} — {elapsed}ms");
            }
            else
            {
                runAction();
                LogAction("Run Macro", SelectedMacro);
            }
        }
        catch (Exception ex)
        {
            LogAction("Run Macro Failed", $"{SelectedMacro}: {ex.Message}");
        }
    }

    private void ExecuteScan()
    {
        IsScanning = true;
        ScanProgress = 0;
        Issues.Clear();
        QuickFixes.Clear();

        var rules = _analyzer.GetRules(StrictMode);
        int total = Modules.Count;
        int done = 0;

        foreach (var mod in Modules)
        {
            var moduleIssues = _analyzer.Analyze(mod, rules);
            foreach (var issue in moduleIssues)
                Issues.Add(issue);
            done++;
            ScanProgress = (double)done / total * 100;
        }

        IssueCount = Issues.Count;
        CriticalCount = Issues.Count(i => i.Severity == "Error");
        WarningCount = Issues.Count(i => i.Severity == "Warning");
        InfoCount = Issues.Count(i => i.Severity == "Info");

        // Generate quick-fixes for found issues
        var fixes = QuickFixEngine.GenerateFixes(Issues, Modules);
        foreach (var fix in fixes) QuickFixes.Add(fix);

        FilterIssues();
        IsScanning = false;
        LogAction("Scan", $"Found {IssueCount} issues ({CriticalCount} errors, {WarningCount} warnings, {InfoCount} info), {QuickFixes.Count} quick-fixes");
    }

    private void ExecuteNavigate()
    {
        try
        {
            System.Windows.Forms.SendKeys.Send("%{F11}");
            LogAction("Navigate", "Opened VBA Editor");
        }
        catch (Exception ex)
        {
            LogAction("Navigate Failed", ex.Message);
        }
    }

    private void ExecuteAutoFix()
    {
        int fixCount = 0;
        var fixDetails = new System.Text.StringBuilder();

        foreach (var mod in Modules)
        {
            if (!mod.Code.Contains("Option Explicit", StringComparison.OrdinalIgnoreCase))
            {
                mod.Code = "Option Explicit\n" + mod.Code;
                fixDetails.AppendLine($"Added Option Explicit to {mod.Name}");
                fixCount++;
            }

            var lines = mod.Code.Split('\n').ToList();
            for (int i = 0; i < lines.Count; i++)
            {
                var trimmed = lines[i].TrimStart();
                if ((trimmed.StartsWith("Sub ", StringComparison.OrdinalIgnoreCase) ||
                     trimmed.StartsWith("Public Sub ", StringComparison.OrdinalIgnoreCase) ||
                     trimmed.StartsWith("Private Sub ", StringComparison.OrdinalIgnoreCase)) &&
                    !trimmed.Contains("()_"))
                {
                    bool hasErrorHandler = false;
                    for (int j = i + 1; j < lines.Count; j++)
                    {
                        var t = lines[j].TrimStart();
                        if (t.StartsWith("End Sub", StringComparison.OrdinalIgnoreCase)) break;
                        if (t.StartsWith("On Error", StringComparison.OrdinalIgnoreCase))
                        {
                            hasErrorHandler = true;
                            break;
                        }
                    }

                    if (!hasErrorHandler)
                    {
                        lines.Insert(i + 1, "    On Error GoTo ErrHandler");
                        fixCount++;
                    }
                }
            }
            mod.Code = string.Join('\n', lines);
        }

        FixDetail = fixDetails.ToString();
        LogAction("Auto-Fix", $"Applied {fixCount} fixes");
        ExecuteScan();
    }

    private async void ExecuteAskAi()
    {
        if (string.IsNullOrWhiteSpace(ApiKey) || string.IsNullOrWhiteSpace(CodeInputText)) return;

        IsAiWorking = true;
        ResponseText = "";

        try
        {
            var prompt = SelectedPrompt == "Generate macro from description"
                ? CustomPrompt
                : $"{SelectedPrompt}:\n\n```vba\n{CodeInputText}\n```";

            if (!string.IsNullOrWhiteSpace(CustomPrompt) && SelectedPrompt != "Generate macro from description")
                prompt += $"\n\nAdditional context: {CustomPrompt}";

            var result = await _aiAssistant.AskAsync(
                ApiKey, prompt, SelectedAiModel, Temperature, MaxResponseTokens, ApiTimeout);

            ResponseText = result.Response;
            TokensUsed = result.TokensUsed;
            LogAction("AI Query", $"Used {TokensUsed} tokens with {SelectedAiModel}");
        }
        catch (Exception ex)
        {
            ResponseText = $"Error: {ex.Message}";
            LogAction("AI Error", ex.Message);
        }
        finally
        {
            IsAiWorking = false;
        }
    }

    private void ExecuteCopyResponse()
    {
        if (!string.IsNullOrEmpty(ResponseText))
        {
            Clipboard.SetText(ResponseText);
            LogAction("Copy", "Response copied to clipboard");
        }
    }

    private void ExecuteInsertCode()
    {
        if (string.IsNullOrEmpty(ResponseText)) return;
        try
        {
            dynamic app = ExcelDna.Integration.ExcelDnaUtil.Application;
            var vbProject = app.ActiveWorkbook.VBProject;
            var newModule = vbProject.VBComponents.Add(1);
            newModule.CodeModule.AddFromString(ResponseText);
            LogAction("Insert Code", $"Created module {newModule.Name}");
            ExecuteRefresh();
        }
        catch (Exception ex)
        {
            LogAction("Insert Failed", ex.Message);
        }
    }

    private void ExecuteNewModule()
    {
        try
        {
            dynamic app = ExcelDna.Integration.ExcelDnaUtil.Application;
            var vbProject = app.ActiveWorkbook.VBProject;
            var newModule = vbProject.VBComponents.Add(1);
            LogAction("New Module", $"Created {newModule.Name}");
            ExecuteRefresh();
        }
        catch (Exception ex)
        {
            LogAction("New Module Failed", ex.Message);
        }
    }

    private void ExecuteDeleteModule()
    {
        if (string.IsNullOrEmpty(SelectedModule)) return;

        var result = MessageBox.Show(
            $"Are you sure you want to delete module '{SelectedModule}'?\nThis cannot be undone.",
            "Confirm Delete",
            MessageBoxButton.YesNo,
            MessageBoxImage.Warning);

        if (result != MessageBoxResult.Yes) return;

        try
        {
            dynamic app = ExcelDna.Integration.ExcelDnaUtil.Application;
            var vbProject = app.ActiveWorkbook.VBProject;
            var component = vbProject.VBComponents.Item(SelectedModule);
            vbProject.VBComponents.Remove(component);
            LogAction("Delete Module", SelectedModule);
            ExecuteRefresh();
        }
        catch (Exception ex)
        {
            LogAction("Delete Failed", ex.Message);
        }
    }

    private void ExecuteExportAll()
    {
        using var dialog = new System.Windows.Forms.FolderBrowserDialog
        {
            Description = "Select folder to export VBA modules",
            ShowNewFolderButton = true
        };

        if (dialog.ShowDialog() != System.Windows.Forms.DialogResult.OK) return;

        int exported = 0;
        try
        {
            dynamic app = ExcelDna.Integration.ExcelDnaUtil.Application;
            foreach (var component in app.ActiveWorkbook.VBProject.VBComponents)
            {
                string ext = ((int)component.Type) switch
                {
                    1 => ".bas",
                    2 => ".cls",
                    3 => ".frm",
                    _ => ".bas"
                };
                string path = System.IO.Path.Combine(dialog.SelectedPath, component.Name + ext);
                component.Export(path);
                exported++;
            }
            LogAction("Export All", $"Exported {exported} modules to {dialog.SelectedPath}");
        }
        catch (Exception ex)
        {
            LogAction("Export Failed", ex.Message);
        }
    }

    private void ExecuteImport()
    {
        using var dialog = new System.Windows.Forms.OpenFileDialog
        {
            Filter = "VBA Files|*.bas;*.cls;*.frm|All Files|*.*",
            Title = "Import VBA Module",
            Multiselect = true
        };

        if (dialog.ShowDialog() != System.Windows.Forms.DialogResult.OK) return;

        try
        {
            dynamic app = ExcelDna.Integration.ExcelDnaUtil.Application;
            foreach (var file in dialog.FileNames)
            {
                app.ActiveWorkbook.VBProject.VBComponents.Import(file);
            }
            LogAction("Import", $"Imported {dialog.FileNames.Length} modules");
            ExecuteRefresh();
        }
        catch (Exception ex)
        {
            LogAction("Import Failed", ex.Message);
        }
    }

    private void ExecuteBackup()
    {
        using var dialog = new System.Windows.Forms.FolderBrowserDialog
        {
            Description = "Select backup folder"
        };

        if (dialog.ShowDialog() != System.Windows.Forms.DialogResult.OK) return;

        try
        {
            string backupDir = System.IO.Path.Combine(dialog.SelectedPath,
                $"VBA_Backup_{DateTime.Now:yyyyMMdd_HHmmss}");
            System.IO.Directory.CreateDirectory(backupDir);

            dynamic app = ExcelDna.Integration.ExcelDnaUtil.Application;
            foreach (var component in app.ActiveWorkbook.VBProject.VBComponents)
            {
                string ext = ((int)component.Type) switch
                {
                    1 => ".bas",
                    2 => ".cls",
                    3 => ".frm",
                    _ => ".bas"
                };
                component.Export(System.IO.Path.Combine(backupDir, component.Name + ext));
            }
            LogAction("Backup", $"Backup saved to {backupDir}");
        }
        catch (Exception ex)
        {
            LogAction("Backup Failed", ex.Message);
        }
    }

    private void ExecuteReset()
    {
        AutoScan = false;
        ScanOnSave = false;
        StrictMode = false;
        MaxLinesToScan = 10000;
        Temperature = 0.3;
        MaxTokens = 4096;
        ApiTimeout = 30;
        MaxResponseTokens = 4096;
        IgnoreComments = false;
        SelectedTheme = "Dark";
        IsDeveloperMode = false;
        IsProfilingEnabled = false;
        LogAction("Reset", "Settings restored to defaults");
    }

    // ══════════════════════════════════════════════════════════════
    //  PREMIUM FEATURE 1 — PERFORMANCE PROFILER
    // ══════════════════════════════════════════════════════════════
    public long ProfileMacroExecution(string macroName, Action macroAction)
    {
        var sw = Stopwatch.StartNew();
        macroAction();
        sw.Stop();

        var record = PerformanceData.FirstOrDefault(r => r.MacroName == macroName);
        if (record == null)
        {
            record = new PerformanceRecord { MacroName = macroName };
            PerformanceData.Add(record);
        }
        record.RecordRun(sw.ElapsedMilliseconds);
        OnPropertyChanged(nameof(PerformanceData));

        return sw.ElapsedMilliseconds;
    }

    // ══════════════════════════════════════════════════════════════
    //  PREMIUM FEATURE 2 — CODE SNIPPETS
    // ══════════════════════════════════════════════════════════════
    private void ExecuteInsertSnippet()
    {
        if (SelectedSnippet == null) { StatusMessage = "Select a snippet first"; return; }

        try
        {
            dynamic app = ExcelDna.Integration.ExcelDnaUtil.Application;
            var vbProject = app.ActiveWorkbook.VBProject;
            var newModule = vbProject.VBComponents.Add(1);
            newModule.CodeModule.AddFromString(SelectedSnippet.Code);
            LogAction("Snippet", $"Inserted '{SelectedSnippet.Name}' as {newModule.Name}");
            ExecuteRefresh();
        }
        catch (Exception ex)
        {
            LogAction("Snippet Insert Failed", ex.Message);
        }
    }

    public string GetSelectedSnippetCode() => SelectedSnippet?.Code ?? "";

    // ══════════════════════════════════════════════════════════════
    //  PREMIUM FEATURE 3 — MACRO SCHEDULER
    // ══════════════════════════════════════════════════════════════
    private void ExecuteAddSchedule()
    {
        if (string.IsNullOrWhiteSpace(ScheduledMacroName)) { StatusMessage = "Select a macro to schedule"; return; }

        var entry = new ScheduleEntry
        {
            MacroName = ScheduledMacroName,
            TriggerType = TriggerType,
            IntervalSeconds = IntervalSeconds,
            IsEnabled = IsScheduleEnabled,
            NextRun = TriggerType == "On Timer" ? DateTime.Now.AddSeconds(IntervalSeconds) : null
        };
        Schedules.Add(entry);

        if (entry.TriggerType == "On Timer" && entry.IsEnabled)
            StartScheduleTimer(entry);

        LogAction("Scheduler", $"Added: {entry}");
    }

    private void StartScheduleTimer(ScheduleEntry entry)
    {
        var key = $"{entry.MacroName}_{entry.GetHashCode()}";
        if (_activeTimers.ContainsKey(key)) return;

        var timer = new Timer(entry.IntervalSeconds * 1000);
        timer.Elapsed += (s, e) =>
        {
            entry.LastRun = DateTime.Now;
            entry.NextRun = DateTime.Now.AddSeconds(entry.IntervalSeconds);
            try
            {
                dynamic app = ExcelDna.Integration.ExcelDnaUtil.Application;
                app.Run(entry.MacroName);
            }
            catch { /* Timer callback — suppress */ }
        };
        timer.AutoReset = true;
        timer.Start();
        _activeTimers[key] = timer;
    }

    public void StopAllTimers()
    {
        foreach (var t in _activeTimers.Values) { t.Stop(); t.Dispose(); }
        _activeTimers.Clear();
    }

    // ══════════════════════════════════════════════════════════════
    //  PREMIUM FEATURE 4 — VBA DIFF / VERSION COMPARISON
    // ══════════════════════════════════════════════════════════════
    private void ExecuteTakeSnapshot()
    {
        var snap = new VbaSnapshot
        {
            Label = $"Snapshot #{Snapshots.Count + 1}",
            Timestamp = DateTime.Now,
            ModuleCode = Modules.ToDictionary(m => m.Name, m => m.Code)
        };
        Snapshots.Add(snap);
        LogAction("Snapshot", $"Saved {snap.Label} ({snap.ModuleCode.Count} modules)");
    }

    private void ExecuteCompareSnapshot()
    {
        if (SelectedSnapshot == null) { StatusMessage = "Select a snapshot to compare"; return; }

        var diffLines = new List<string>();
        foreach (var mod in Modules)
        {
            var oldCode = SelectedSnapshot.ModuleCode.GetValueOrDefault(mod.Name, "");
            var diff = DiffEngine.ComputeDiff(oldCode, mod.Code);
            bool hasDiff = diff.Any(d => d.Kind != DiffEngine.DiffKind.Unchanged);
            if (!hasDiff) continue;

            diffLines.Add($"=== {mod.Name} ===");
            foreach (var d in diff)
            {
                var prefix = d.Kind switch
                {
                    DiffEngine.DiffKind.Added => "+ ",
                    DiffEngine.DiffKind.Removed => "- ",
                    _ => "  "
                };
                diffLines.Add(prefix + d.Text);
            }
            diffLines.Add("");
        }

        DiffText = diffLines.Count > 0 ? string.Join("\n", diffLines) : "(No differences found)";
        LogAction("Diff", $"Compared against {SelectedSnapshot.Label}");
    }

    // ══════════════════════════════════════════════════════════════
    //  PREMIUM FEATURE 5 — CODE METRICS
    // ══════════════════════════════════════════════════════════════
    private static readonly Regex ProcStartRx = new(@"^\s*(Public\s+|Private\s+)?(Sub|Function)\s+(\w+)", RegexOptions.IgnoreCase | RegexOptions.Compiled);
    private static readonly Regex ProcEndRx = new(@"^\s*End\s+(Sub|Function)", RegexOptions.IgnoreCase | RegexOptions.Compiled);
    private static readonly Regex BranchRx = new(@"\b(If|ElseIf|Select\s+Case|For\s+Each|For\s+|Do\s+While|Do\s+Until|While\s+)\b", RegexOptions.IgnoreCase | RegexOptions.Compiled);
    private static readonly Regex CommentRx = new(@"^\s*'", RegexOptions.Compiled);
    private static readonly Regex CallRx = new(@"\b(Call\s+)?(\w+)\s*\(", RegexOptions.IgnoreCase | RegexOptions.Compiled);

    private void ExecuteRefreshMetrics()
    {
        MetricsData.Clear();
        int totalLines = 0, totalComments = 0, totalProcs = 0;
        double totalProcLen = 0;
        int totalComplexity = 0;

        foreach (var mod in Modules)
        {
            var m = AnalyzeModuleMetrics(mod.Name, mod.Code);
            MetricsData.Add(m);
            totalLines += m.TotalLines;
            totalComments += m.CommentLines;
            totalProcs += m.ProcedureCount;
            totalProcLen += m.AvgProcedureLength * m.ProcedureCount;
            totalComplexity += m.CyclomaticComplexity;
        }

        TotalLinesOfCode = totalLines;
        OverallCommentRatio = totalLines > 0 ? Math.Round(100.0 * totalComments / totalLines, 1) : 0;
        OverallAvgProcLength = totalProcs > 0 ? Math.Round(totalProcLen / totalProcs, 1) : 0;
        OverallComplexity = totalComplexity;
        DeadProcedureCount = CountDeadProcedures();

        LogAction("Metrics", $"Analyzed {Modules.Count} modules, {totalLines} total lines");
    }

    private static ModuleMetrics AnalyzeModuleMetrics(string moduleName, string code)
    {
        var lines = code.Split('\n');
        int commentLines = lines.Count(l => CommentRx.IsMatch(l));
        int procCount = 0, totalProcLines = 0, complexity = 0;
        bool inProc = false;
        int procStartLine = 0;

        for (int i = 0; i < lines.Length; i++)
        {
            if (ProcStartRx.IsMatch(lines[i])) { inProc = true; procStartLine = i; procCount++; }
            if (inProc) complexity += BranchRx.Matches(lines[i]).Count;
            if (ProcEndRx.IsMatch(lines[i]) && inProc) { totalProcLines += i - procStartLine + 1; inProc = false; }
        }

        return new ModuleMetrics
        {
            ModuleName = moduleName,
            TotalLines = lines.Length,
            CommentLines = commentLines,
            ProcedureCount = procCount,
            AvgProcedureLength = procCount > 0 ? Math.Round((double)totalProcLines / procCount, 1) : 0,
            CyclomaticComplexity = complexity
        };
    }

    private int CountDeadProcedures()
    {
        var allCode = string.Join("\n", Modules.Select(m => m.Code));
        var allProcs = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        var calledProcs = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

        foreach (var line in allCode.Split('\n'))
        {
            var sm = ProcStartRx.Match(line);
            if (sm.Success) allProcs.Add(sm.Groups[3].Value);
            foreach (Match cm in CallRx.Matches(line))
                calledProcs.Add(cm.Groups[2].Value);
        }

        return allProcs.Count(p => !calledProcs.Contains(p));
    }

    // ══════════════════════════════════════════════════════════════
    //  PREMIUM FEATURE 6 — DEPENDENCY GRAPH
    // ══════════════════════════════════════════════════════════════
    private void ExecuteRefreshDependencies()
    {
        Dependencies.Clear();

        var allProcNames = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        foreach (var mod in Modules)
            foreach (var line in mod.Code.Split('\n'))
            {
                var m = ProcStartRx.Match(line);
                if (m.Success) allProcNames.Add(m.Groups[3].Value);
            }

        foreach (var mod in Modules)
        {
            var lines = mod.Code.Split('\n');
            string? currentProc = null;
            var callees = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

            foreach (var line in lines)
            {
                var startMatch = ProcStartRx.Match(line);
                if (startMatch.Success)
                {
                    if (currentProc != null && callees.Count > 0)
                        Dependencies.Add(new DependencyInfo { CallerModule = mod.Name, CallerProcedure = currentProc, Callees = callees.ToList() });
                    currentProc = startMatch.Groups[3].Value;
                    callees = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
                    continue;
                }
                if (ProcEndRx.IsMatch(line))
                {
                    if (currentProc != null && callees.Count > 0)
                        Dependencies.Add(new DependencyInfo { CallerModule = mod.Name, CallerProcedure = currentProc, Callees = callees.ToList() });
                    currentProc = null;
                    continue;
                }
                if (currentProc == null) continue;
                foreach (Match cm in CallRx.Matches(line))
                {
                    var callee = cm.Groups[2].Value;
                    if (allProcNames.Contains(callee) && !string.Equals(callee, currentProc, StringComparison.OrdinalIgnoreCase))
                        callees.Add(callee);
                }
            }
            if (currentProc != null && callees.Count > 0)
                Dependencies.Add(new DependencyInfo { CallerModule = mod.Name, CallerProcedure = currentProc, Callees = callees.ToList() });
        }

        LogAction("Dependencies", $"Found {Dependencies.Count} call relationships");
    }

    // ══════════════════════════════════════════════════════════════
    //  PREMIUM FEATURE 7 — EXPORT
    // ══════════════════════════════════════════════════════════════
    private void ExecuteExportBas()
    {
        var mod = Modules.FirstOrDefault(m => m.Name == SelectedModule);
        if (mod == null) { StatusMessage = "Select a module to export"; return; }
        var content = ExportService.ExportAsBas(mod);
        ExportPath = $"{mod.Name}.bas";
        LogAction("Export", $"Exported {mod.Name} as .bas ({content.Length} chars)");
    }

    private void ExecuteExportHtmlDoc()
    {
        var html = ExportService.ExportDocumentationHtml(Modules);
        ExportPath = "VBA_Documentation.html";
        LogAction("Export", $"Exported HTML documentation ({html.Length} chars)");
    }

    private void ExecuteExportHealthCsv()
    {
        var csv = ExportService.ExportHealthCsv(Issues);
        ExportPath = "HealthCheck_Report.csv";
        LogAction("Export", $"Exported health CSV ({Issues.Count} issues)");
    }

    private void ExecuteExportAllText()
    {
        var text = ExportService.ExportAllAsText(Modules);
        ExportPath = "AllVBACode.txt";
        LogAction("Export", $"Exported all code as text ({text.Length} chars)");
    }

    // Content getters for the pane to write to disk
    public string GetExportBasContent() => Modules.FirstOrDefault(m => m.Name == SelectedModule) is { } mod ? ExportService.ExportAsBas(mod) : "";
    public string GetExportHtmlContent() => ExportService.ExportDocumentationHtml(Modules);
    public string GetExportCsvContent() => ExportService.ExportHealthCsv(Issues);
    public string GetExportAllTextContent() => ExportService.ExportAllAsText(Modules);

    // ══════════════════════════════════════════════════════════════
    //  PREMIUM FEATURE 8 — QUICK-FIX TEMPLATES
    // ══════════════════════════════════════════════════════════════
    private void ExecuteApplyQuickFix()
    {
        if (SelectedQuickFix == null) { StatusMessage = "Select a fix first"; return; }

        var mod = Modules.FirstOrDefault(m => string.Equals(m.Name, SelectedQuickFix.TargetModule, StringComparison.OrdinalIgnoreCase));
        if (mod == null) { StatusMessage = $"Module '{SelectedQuickFix.TargetModule}' not found"; return; }

        mod.Code = QuickFixEngine.ApplyFix(SelectedQuickFix, mod.Code);
        mod.LineCount = mod.Code.Split('\n').Length;

        LogAction("QuickFix", $"Applied '{SelectedQuickFix.Name}' to {mod.Name}");

        // Re-scan to update issues and fixes
        ExecuteScan();
    }

    // ── INotifyPropertyChanged ──────────────────────────────────
    public event PropertyChangedEventHandler? PropertyChanged;
    private void OnPropertyChanged([CallerMemberName] string? name = null)
        => PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(name));
    private void Set<T>(ref T field, T value, [CallerMemberName] string? name = null)
    {
        if (EqualityComparer<T>.Default.Equals(field, value)) return;
        field = value;
        OnPropertyChanged(name);
    }
}
