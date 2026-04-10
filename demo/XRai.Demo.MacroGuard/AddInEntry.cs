using System.Diagnostics;
using System.Windows;
using System.Windows.Threading;
using ExcelDna.Integration;
using XRai.Demo.MacroGuard.Scanner;
using XRai.Hooks;

namespace XRai.Demo.MacroGuard;

public class AddInEntry : IExcelAddIn
{
    private static MacroGuardViewModel? _viewModel;
    private static MacroGuardPane? _pane;
    private static Window? _window;
    private static Thread? _wpfThread;

    public void AutoOpen()
    {
        Pilot.Start();

        _viewModel = new MacroGuardViewModel();
        Pilot.ExposeModel(_viewModel, "MacroGuard");

        // Check VBA trust access
        try
        {
            _viewModel.VbaAccessEnabled = VbaAnalyzer.CheckVbaAccess();
            if (!_viewModel.VbaAccessEnabled)
            {
                _viewModel.StatusMessage = "VBA access disabled — enable 'Trust access to the VBA project object model' in Trust Center";
            }
        }
        catch
        {
            _viewModel.VbaAccessEnabled = false;
        }

        // Launch the WPF pane as a floating window on a dedicated STA thread
        _wpfThread = new Thread(() =>
        {
            try
            {
                _pane = new MacroGuardPane(_viewModel);
                _window = new Window
                {
                    Title = "XRai MacroGuard",
                    Content = _pane,
                    Width = 460,
                    Height = 780,
                    WindowStyle = WindowStyle.ToolWindow,
                    Topmost = true,
                    ShowInTaskbar = false,
                };

                _window.Loaded += (s, e) =>
                {
                    // Position to the right of Excel
                    try
                    {
                        dynamic app = ExcelDnaUtil.Application;
                        _window.Left = (double)app.Left + (double)app.Width - 10;
                        _window.Top = (double)app.Top + 50;
                    }
                    catch { }

                    // Expose controls to hooks
                    Pilot.Expose(_pane);
                    Pilot.Log($"MacroGuard pane visible with {_pane.ViewModel.ModuleCount} modules, {_pane.ViewModel.MacroCount} macros");
                };

                _window.Show();
                Dispatcher.Run();
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"WPF window error: {ex}");
            }
        });
        _wpfThread.SetApartmentState(ApartmentState.STA);
        _wpfThread.IsBackground = true;
        _wpfThread.Name = "XRai-MacroGuardPane";
        _wpfThread.Start();

        Pilot.Log("MacroGuard add-in loaded");
    }

    public void AutoClose()
    {
        try { _window?.Dispatcher.InvokeShutdown(); } catch { }
        Pilot.Stop();
    }

    public static MacroGuardViewModel? ViewModel => _viewModel;
}

public static class MacroGuardFunctions
{
    [ExcelFunction(Name = "MACROGUARD.ISSUES", Description = "Returns issue count for a VBA module")]
    public static object Issues(string moduleName)
    {
        var vm = AddInEntry.ViewModel;
        if (vm == null) return ExcelError.ExcelErrorNA;

        var module = vm.Modules.FirstOrDefault(m =>
            string.Equals(m.Name, moduleName, StringComparison.OrdinalIgnoreCase));
        if (module == null) return ExcelError.ExcelErrorNA;

        var analyzer = new VbaAnalyzer();
        var rules = analyzer.GetRules(vm.StrictMode);
        var issues = analyzer.Analyze(module, rules);
        return issues.Count;
    }

    [ExcelFunction(Name = "MACROGUARD.COMPLEXITY", Description = "Returns line count + procedure count for a VBA module")]
    public static object Complexity(string moduleName)
    {
        var vm = AddInEntry.ViewModel;
        if (vm == null) return ExcelError.ExcelErrorNA;

        var module = vm.Modules.FirstOrDefault(m =>
            string.Equals(m.Name, moduleName, StringComparison.OrdinalIgnoreCase));
        if (module == null) return ExcelError.ExcelErrorNA;

        return $"{module.LineCount} lines, {module.MacroNames.Count} procedures";
    }
}
