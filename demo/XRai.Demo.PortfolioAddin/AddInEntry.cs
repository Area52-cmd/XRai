using System.Diagnostics;
using System.Windows;
using System.Windows.Threading;
using ExcelDna.Integration;
using XRai.Hooks;

namespace XRai.Demo.PortfolioAddin;

public class AddInEntry : IExcelAddIn
{
    private static PortfolioViewModel? _viewModel;
    private static PortfolioPane? _pane;
    private static Window? _window;
    private static Thread? _wpfThread;

    public void AutoOpen()
    {
        Pilot.Start();

        _viewModel = new PortfolioViewModel();
        Pilot.ExposeModel(_viewModel, "Portfolio");

        // Launch the WPF pane as a floating window on a dedicated STA thread
        _wpfThread = new Thread(() =>
        {
            try
            {
                _pane = new PortfolioPane(_viewModel);
                _window = new Window
                {
                    Title = "XRai Portfolio Tracker",
                    Content = _pane,
                    Width = 420,
                    Height = 700,
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
                    Pilot.Log($"Portfolio pane visible with {_pane.ViewModel.Holdings.Count} stocks");
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
        _wpfThread.Name = "XRai-PortfolioPane";
        _wpfThread.Start();

        Pilot.Log("Portfolio add-in loaded");
    }

    public void AutoClose()
    {
        try { _window?.Dispatcher.InvokeShutdown(); } catch { }
        Pilot.Stop();
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
