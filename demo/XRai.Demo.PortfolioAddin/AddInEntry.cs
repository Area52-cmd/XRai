using ExcelDna.Integration;
using ExcelDna.Integration.CustomUI;
using XRai.Hooks;

namespace XRai.Demo.PortfolioAddin;

/// <summary>
/// Demo Excel-DNA addin demonstrating XRai integration with full WPF UI.
///
/// CRITICAL design choices for clean .NET 8 + Excel-DNA + WPF teardown:
///   1. .dna sets LoadFromBytes="false" — addin loads into the DEFAULT
///      AssemblyLoadContext (non-collectible). CoreCLR's strict ALC unload
///      sweep does NOT run on the default context, sidestepping the
///      0x80131506 FailFast that WPF static state otherwise triggers.
///   2. WPF lives inside a CustomTaskPane via WinForms ElementHost — on
///      Excel's MAIN UI thread. No dedicated STA thread, no Dispatcher.Run().
///      Side-thread WPF roots WPF statics in ways that break unload even
///      with LoadFromBytes=false.
///   3. AutoClose calls Pilot.Shutdown — single-call cleanup.
/// </summary>
public class AddInEntry : IExcelAddIn
{
    private static PortfolioViewModel? _viewModel;
    private static CustomTaskPane? _ctp;
    private static TaskPaneHost? _host;

    public void AutoOpen()
    {
        Pilot.Start();

        _viewModel = new PortfolioViewModel();
        Pilot.ExposeModel(_viewModel, "Portfolio");

        // CTP creation must happen on Excel's main thread once Excel is ready.
        // Doing it directly in AutoOpen sometimes runs before Excel finishes
        // initializing the COM bridge needed by CustomTaskPaneFactory.
        ExcelAsyncUtil.QueueAsMacro(() =>
        {
            try
            {
                _host = new TaskPaneHost(_viewModel!);
                _ctp = CustomTaskPaneFactory.CreateCustomTaskPane(_host, "XRai Portfolio Tracker");
                _ctp.DockPosition = MsoCTPDockPosition.msoCTPDockPositionRight;
                _ctp.Width = 420;
                _ctp.Visible = true;

                _host.Pane.Loaded += (_, _) =>
                {
                    try
                    {
                        Pilot.Expose(_host.Pane);
                        Pilot.Log($"Portfolio pane visible with {_host.Pane.ViewModel.Holdings.Count} stocks");
                    }
                    catch (Exception ex) { Pilot.Log($"Expose failed: {ex.Message}"); }
                };
            }
            catch (Exception ex) { Pilot.Log($"CTP create failed: {ex.Message}"); }
        });

        Pilot.Log("Portfolio add-in loaded (CTP+ElementHost, LoadFromBytes=false)");
    }

    public void AutoClose()
    {
        try { _ctp?.Delete(); } catch { }
        _ctp = null;
        _host = null;
        Pilot.Shutdown();
    }

    public static PortfolioViewModel? ViewModel => _viewModel;
}

public static class PortfolioFunctions
{
    [ExcelFunction(Name = "XRAI.PRICE", Description = "Get current mock stock price")]
    public static object XraiPrice(string symbol)
    {
        var vm = AddInEntry.ViewModel;
        if (vm == null) return ExcelError.ExcelErrorNA;
        var h = vm.Holdings.FirstOrDefault(x => string.Equals(x.Symbol, symbol, StringComparison.OrdinalIgnoreCase));
        return h?.CurrentPrice ?? (object)ExcelError.ExcelErrorNA;
    }

    [ExcelFunction(Name = "XRAI.PNL", Description = "Calculate P&L for position")]
    public static object XraiPnl(string symbol, int quantity, double costBasis)
    {
        var vm = AddInEntry.ViewModel;
        if (vm == null) return ExcelError.ExcelErrorNA;
        var h = vm.Holdings.FirstOrDefault(x => string.Equals(x.Symbol, symbol, StringComparison.OrdinalIgnoreCase));
        if (h == null) return ExcelError.ExcelErrorNA;
        return (h.CurrentPrice - costBasis) * quantity;
    }

    [ExcelFunction(Name = "XRAI.PORTFOLIO.VALUE", Description = "Total portfolio market value")]
    public static object XraiPortfolioValue() => AddInEntry.ViewModel?.TotalValue ?? (object)ExcelError.ExcelErrorNA;

    [ExcelFunction(Name = "XRAI.PORTFOLIO.PNL", Description = "Total portfolio P&L")]
    public static object XraiPortfolioPnl() => AddInEntry.ViewModel?.TotalPnL ?? (object)ExcelError.ExcelErrorNA;
}
