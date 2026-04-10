using System.Windows.Forms;
using System.Windows.Forms.Integration;

namespace XRai.Demo.PortfolioAddin;

/// <summary>
/// WinForms UserControl that hosts the WPF PortfolioPane via ElementHost.
/// Required because Excel custom task panes only accept WinForms controls.
/// </summary>
public class TaskPaneHost : UserControl
{
    private readonly ElementHost _host;
    private readonly PortfolioPane _pane;

    public TaskPaneHost(PortfolioViewModel viewModel)
    {
        _pane = new PortfolioPane(viewModel);
        _host = new ElementHost
        {
            Dock = DockStyle.Fill,
            Child = _pane
        };
        Controls.Add(_host);
    }

    public PortfolioPane Pane => _pane;
}
