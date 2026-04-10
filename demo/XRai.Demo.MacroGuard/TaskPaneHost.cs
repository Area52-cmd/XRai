using System.Windows.Forms;
using System.Windows.Forms.Integration;

namespace XRai.Demo.MacroGuard;

/// <summary>
/// WinForms UserControl that hosts the WPF MacroGuardPane via ElementHost.
/// Required because Excel custom task panes only accept WinForms controls.
/// </summary>
public class TaskPaneHost : UserControl
{
    private readonly ElementHost _host;
    private readonly MacroGuardPane _pane;

    public TaskPaneHost(MacroGuardViewModel viewModel)
    {
        _pane = new MacroGuardPane(viewModel);
        _host = new ElementHost
        {
            Dock = DockStyle.Fill,
            Child = _pane
        };
        Controls.Add(_host);
    }

    public MacroGuardPane Pane => _pane;
}
