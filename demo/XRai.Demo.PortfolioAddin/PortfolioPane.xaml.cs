using System.Windows;

namespace XRai.Demo.PortfolioAddin;

public partial class PortfolioPane : System.Windows.Controls.UserControl
{
    private readonly PortfolioViewModel _vm;

    public PortfolioPane() : this(new PortfolioViewModel()) { }

    public PortfolioPane(PortfolioViewModel vm)
    {
        _vm = vm;
        DataContext = _vm;
        InitializeComponent();
        TradeDatePicker.SelectedDate = DateTime.Today;
    }

    public PortfolioViewModel ViewModel => _vm;

    private void RefreshButton_Click(object sender, RoutedEventArgs e)
    {
        _vm.Progress = 50;
        _vm.RefreshPrices();
        _vm.Progress = 100;
        _vm.StatusMessage = "Prices refreshed successfully";
    }

    private void ExecuteTrade_Click(object sender, RoutedEventArgs e)
    {
        _vm.SelectedSide = BuyRadio.IsChecked == true ? "Buy" : "Sell";
        _vm.ExecuteTrade();
    }

    private void Connect_Click(object sender, RoutedEventArgs e)
    {
        _vm.IsConnected = true;
        _vm.StatusMessage = "Connected to market data";
    }

    private void Disconnect_Click(object sender, RoutedEventArgs e)
    {
        _vm.IsConnected = false;
        _vm.StatusMessage = "Disconnected";
    }
}
