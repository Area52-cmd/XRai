using System.ComponentModel;
using System.Runtime.CompilerServices;

namespace XRai.Demo.PortfolioAddin.Models;

public class StockPosition : INotifyPropertyChanged
{
    private string _symbol = "";
    private string _name = "";
    private int _quantity;
    private double _costBasis;
    private double _currentPrice;
    private string _sector = "";

    public string Symbol { get => _symbol; set => Set(ref _symbol, value); }
    public string Name { get => _name; set => Set(ref _name, value); }
    public int Quantity { get => _quantity; set { Set(ref _quantity, value); OnPropertyChanged(nameof(MarketValue)); OnPropertyChanged(nameof(PnL)); OnPropertyChanged(nameof(PnLPercent)); } }
    public double CostBasis { get => _costBasis; set { Set(ref _costBasis, value); OnPropertyChanged(nameof(PnL)); OnPropertyChanged(nameof(PnLPercent)); } }
    public double CurrentPrice { get => _currentPrice; set { Set(ref _currentPrice, value); OnPropertyChanged(nameof(MarketValue)); OnPropertyChanged(nameof(PnL)); OnPropertyChanged(nameof(PnLPercent)); } }
    public string Sector { get => _sector; set => Set(ref _sector, value); }

    public double MarketValue => Quantity * CurrentPrice;
    public double PnL => MarketValue - (Quantity * CostBasis);
    public double PnLPercent => CostBasis > 0 ? (CurrentPrice - CostBasis) / CostBasis * 100 : 0;

    public event PropertyChangedEventHandler? PropertyChanged;
    private void OnPropertyChanged([CallerMemberName] string? name = null) => PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(name));
    private void Set<T>(ref T field, T value, [CallerMemberName] string? name = null) { field = value; OnPropertyChanged(name); }
}
