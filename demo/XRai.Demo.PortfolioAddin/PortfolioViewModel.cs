using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Runtime.CompilerServices;
using XRai.Demo.PortfolioAddin.Models;

namespace XRai.Demo.PortfolioAddin;

public class PortfolioViewModel : INotifyPropertyChanged
{
    private string _portfolioName = "My Portfolio";
    private string _searchSymbol = "";
    private string _selectedAccount = "Personal";
    private bool _isAutoRefresh;
    private int _refreshInterval = 5;
    private double _riskTolerance = 50;
    private bool _isConnected = true;
    private double _totalValue;
    private double _totalPnL;
    private string _lastUpdated = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss");
    private string _statusMessage = "Ready";
    private double _progress;
    private string _orderType = "Market";
    private int _orderQuantity = 100;
    private double _limitPrice;
    private string _selectedSide = "Buy";

    // Core properties
    public string PortfolioName { get => _portfolioName; set => Set(ref _portfolioName, value); }
    public string SearchSymbol { get => _searchSymbol; set => Set(ref _searchSymbol, value); }
    public string SelectedAccount { get => _selectedAccount; set => Set(ref _selectedAccount, value); }
    public bool IsAutoRefresh { get => _isAutoRefresh; set => Set(ref _isAutoRefresh, value); }
    public int RefreshInterval { get => _refreshInterval; set => Set(ref _refreshInterval, value); }
    public double RiskTolerance { get => _riskTolerance; set => Set(ref _riskTolerance, value); }
    public bool IsConnected { get => _isConnected; set => Set(ref _isConnected, value); }
    public double TotalValue { get => _totalValue; set => Set(ref _totalValue, value); }
    public double TotalPnL { get => _totalPnL; set => Set(ref _totalPnL, value); }
    public string LastUpdated { get => _lastUpdated; set => Set(ref _lastUpdated, value); }
    public string StatusMessage { get => _statusMessage; set => Set(ref _statusMessage, value); }
    public double Progress { get => _progress; set => Set(ref _progress, value); }

    // Trade form
    public string OrderType { get => _orderType; set => Set(ref _orderType, value); }
    public int OrderQuantity { get => _orderQuantity; set => Set(ref _orderQuantity, value); }
    public double LimitPrice { get => _limitPrice; set => Set(ref _limitPrice, value); }
    public string SelectedSide { get => _selectedSide; set => Set(ref _selectedSide, value); }

    // Collections
    public ObservableCollection<StockPosition> Holdings { get; } = new();
    public ObservableCollection<TradeOrder> TradeHistory { get; } = new();
    public ObservableCollection<string> Accounts { get; } = new() { "Personal", "IRA", "401k", "Joint" };
    public ObservableCollection<string> OrderTypes { get; } = new() { "Market", "Limit", "Stop", "Stop Limit" };

    public PortfolioViewModel()
    {
        LoadSampleData();
    }

    private void LoadSampleData()
    {
        Holdings.Add(new StockPosition { Symbol = "AAPL", Name = "Apple Inc.", Quantity = 50, CostBasis = 142.50, CurrentPrice = 178.25, Sector = "Technology" });
        Holdings.Add(new StockPosition { Symbol = "MSFT", Name = "Microsoft Corp.", Quantity = 30, CostBasis = 285.00, CurrentPrice = 415.80, Sector = "Technology" });
        Holdings.Add(new StockPosition { Symbol = "GOOGL", Name = "Alphabet Inc.", Quantity = 15, CostBasis = 120.00, CurrentPrice = 175.50, Sector = "Technology" });
        Holdings.Add(new StockPosition { Symbol = "JPM", Name = "JPMorgan Chase", Quantity = 40, CostBasis = 148.75, CurrentPrice = 195.20, Sector = "Finance" });
        Holdings.Add(new StockPosition { Symbol = "JNJ", Name = "Johnson & Johnson", Quantity = 25, CostBasis = 165.00, CurrentPrice = 152.30, Sector = "Healthcare" });
        Holdings.Add(new StockPosition { Symbol = "AMZN", Name = "Amazon.com", Quantity = 20, CostBasis = 135.00, CurrentPrice = 185.60, Sector = "Consumer" });
        Holdings.Add(new StockPosition { Symbol = "NVDA", Name = "NVIDIA Corp.", Quantity = 10, CostBasis = 450.00, CurrentPrice = 875.50, Sector = "Technology" });
        Holdings.Add(new StockPosition { Symbol = "TSLA", Name = "Tesla Inc.", Quantity = 15, CostBasis = 220.00, CurrentPrice = 175.40, Sector = "Automotive" });

        TradeHistory.Add(new TradeOrder { Symbol = "AAPL", Side = "Buy", Quantity = 50, Price = 142.50, Status = "Filled" });
        TradeHistory.Add(new TradeOrder { Symbol = "NVDA", Side = "Buy", Quantity = 10, Price = 450.00, Status = "Filled" });
        TradeHistory.Add(new TradeOrder { Symbol = "TSLA", Side = "Buy", Quantity = 15, Price = 220.00, Status = "Filled" });

        RecalcTotals();
    }

    public void RecalcTotals()
    {
        TotalValue = Holdings.Sum(h => h.MarketValue);
        TotalPnL = Holdings.Sum(h => h.PnL);
        LastUpdated = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss");
    }

    public void ExecuteTrade()
    {
        if (string.IsNullOrEmpty(SearchSymbol)) return;
        var order = new TradeOrder
        {
            Symbol = SearchSymbol.ToUpper(),
            Side = SelectedSide,
            Quantity = OrderQuantity,
            Price = LimitPrice > 0 ? LimitPrice : 100.0, // Mock price
            Status = "Filled"
        };
        TradeHistory.Insert(0, order);
        StatusMessage = $"Trade executed: {order}";
    }

    public void RefreshPrices()
    {
        var rng = new Random();
        foreach (var h in Holdings)
        {
            h.CurrentPrice *= 1 + (rng.NextDouble() - 0.48) * 0.02; // Small random move
        }
        RecalcTotals();
        StatusMessage = "Prices refreshed";
    }

    public event PropertyChangedEventHandler? PropertyChanged;
    private void OnPropertyChanged([CallerMemberName] string? name = null) => PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(name));
    private void Set<T>(ref T field, T value, [CallerMemberName] string? name = null) { field = value; OnPropertyChanged(name); }
}
