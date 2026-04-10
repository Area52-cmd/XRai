namespace XRai.Demo.PortfolioAddin.Models;

public class TradeOrder
{
    public DateTime Timestamp { get; set; } = DateTime.Now;
    public string Symbol { get; set; } = "";
    public string Side { get; set; } = "Buy";   // Buy or Sell
    public int Quantity { get; set; }
    public double Price { get; set; }
    public string Status { get; set; } = "Pending";

    public double Total => Quantity * Price;
    public override string ToString() => $"{Timestamp:HH:mm:ss} {Side} {Quantity} {Symbol} @ {Price:C2}";
}
