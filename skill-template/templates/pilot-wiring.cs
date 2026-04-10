// ═══════════════════════════════════════════════════════════════════════
// XRai.Hooks wiring reference for IExcelAddIn implementations.
//
// This is a TEMPLATE — adapt to your existing AddInEntry / IExcelAddIn class.
// Claude Code: when setting up XRai in a new add-in project, merge this
// code into the user's existing IExcelAddIn implementation.
// ═══════════════════════════════════════════════════════════════════════

using System;
using System.Threading;
using System.Windows;
using System.Windows.Threading;
using ExcelDna.Integration;
using XRai.Hooks;

namespace YourNamespace;

public class AddInEntry : IExcelAddIn
{
    // Keep references so they don't get garbage collected
    private static YourViewModel? _viewModel;
    private static YourTaskPane? _pane;
    private static Window? _window;
    private static Thread? _wpfThread;

    public void AutoOpen()
    {
        // 1. Start the XRai hooks pipe server
        //    This creates a named pipe "xrai_{excel_pid}" that the XRai CLI connects to.
        Pilot.Start();

        // 2. Create your ViewModel and expose it
        //    XRai will be able to read ALL INotifyPropertyChanged properties + collections
        _viewModel = new YourViewModel();
        Pilot.ExposeModel(_viewModel, "YourModelName");

        // 3. Create and expose your WPF task pane
        //
        // There are two approaches — pick whichever matches your existing setup:

        // ─── APPROACH A: Excel-DNA Custom Task Pane (docks inside Excel) ─────
        // Requires COM-visible UserControl. See Excel-DNA docs for CustomTaskPaneFactory.
        // Uncomment if you use this approach:
        /*
        ExcelAsyncUtil.QueueAsMacro(() =>
        {
            var taskPaneControl = new YourTaskPaneControl(_viewModel);
            var ctp = ExcelDna.Integration.CustomUI.CustomTaskPaneFactory
                .CreateCustomTaskPane(taskPaneControl, "Your Pane Title");
            ctp.Visible = true;
            ctp.Width = 400;

            // Expose to XRai after the pane is actually visible
            Pilot.Expose(taskPaneControl);
            Pilot.Log("Task pane ready");
        });
        */

        // ─── APPROACH B: Floating WPF Window (separate window near Excel) ────
        // Simpler, no COM visibility needed. Runs on a dedicated STA thread.
        _wpfThread = new Thread(() =>
        {
            _pane = new YourTaskPane(_viewModel);
            _window = new Window
            {
                Title = "Your Add-In Pane",
                Content = _pane,
                Width = 420,
                Height = 700,
                WindowStyle = WindowStyle.ToolWindow,
                Topmost = true,
                ShowInTaskbar = false,
            };

            _window.Loaded += (s, e) =>
            {
                // Position near Excel (optional)
                try
                {
                    dynamic app = ExcelDnaUtil.Application;
                    _window.Left = (double)app.Left + (double)app.Width - 10;
                    _window.Top = (double)app.Top + 50;
                }
                catch { }

                // Expose WPF controls to XRai after the visual tree is rendered
                Pilot.Expose(_pane);
                Pilot.Log("WPF pane exposed");
            };

            _window.Show();
            Dispatcher.Run();
        });
        _wpfThread.SetApartmentState(ApartmentState.STA);
        _wpfThread.IsBackground = true;
        _wpfThread.Name = "YourAddin-WPF";
        _wpfThread.Start();
    }

    public void AutoClose()
    {
        // Clean shutdown — stops the pipe server and releases resources
        try { _window?.Dispatcher.InvokeShutdown(); } catch { }
        Pilot.Stop();
    }

    // Expose static accessors so your ExcelFunction UDFs can reach the ViewModel
    public static YourViewModel? ViewModel => _viewModel;
}

// ═══════════════════════════════════════════════════════════════════════
// CRITICAL: Name your WPF controls with x:Name="..." in XAML
//
// XRai discovers controls by walking the visual tree and reading FrameworkElement.Name.
// Controls WITHOUT x:Name are invisible to XRai.
//
// Example XAML:
//   <TextBox x:Name="SpotInput" Text="{Binding Spot}" />
//   <Button x:Name="CalcButton" Content="Calculate" Click="Calc_Click" />
//   <DataGrid x:Name="TradesGrid" ItemsSource="{Binding Trades}" />
//   <TabControl x:Name="MainTabs">
//       <TabItem Header="Summary">...</TabItem>
//       <TabItem Header="Detail">...</TabItem>
//   </TabControl>
//
// Then from XRai:
//   {"cmd":"pane.type","control":"SpotInput","value":"105"}
//   {"cmd":"pane.click","control":"CalcButton"}
//   {"cmd":"pane.grid.read","control":"TradesGrid"}
//   {"cmd":"pane.tab","control":"MainTabs","tab":"Detail"}
// ═══════════════════════════════════════════════════════════════════════
