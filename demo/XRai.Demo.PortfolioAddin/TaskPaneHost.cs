using System.Runtime.InteropServices;
using System.Windows.Forms;
using System.Windows.Forms.Integration;

namespace XRai.Demo.PortfolioAddin;

/// <summary>Marker interface for COM IDispatch (.NET 6+ Office CTP requirement).</summary>
public interface ITaskPaneHost { }

/// <summary>
/// WinForms UserControl bridging the Office CTP to the WPF PortfolioPane via
/// ElementHost. Office CTPs only accept ActiveX-compatible (WinForms) controls;
/// WPF can't expose itself as ActiveX, so this thin shell wraps an ElementHost.
///
/// Critically, this lives on Excel's MAIN UI thread — not a dedicated STA
/// thread. Hosting WPF on a side thread inside an Excel-DNA addin triggers
/// CoreCLR's strict-shutdown FailFast (0x80131506) on Excel exit. Side-thread
/// WPF was the previous demo pattern; this CTP+ElementHost pattern is the one
/// Excel-DNA officially documents.
/// </summary>
[ComVisible(true)]
[ComDefaultInterface(typeof(ITaskPaneHost))]
[Guid("F8B2C19A-3D7E-4D9F-A45E-1B3C56C7F901")]
public sealed class TaskPaneHost : UserControl, ITaskPaneHost
{
    private readonly ElementHost _host;
    private readonly PortfolioPane _pane;

    public TaskPaneHost() : this(new PortfolioViewModel()) { }

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

    protected override void Dispose(bool disposing)
    {
        if (disposing) _host.Dispose();
        base.Dispose(disposing);
    }
}
