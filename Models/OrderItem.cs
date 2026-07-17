namespace AtolGenerator.Models;

public class OrderItem
{
    public string Name     { get; set; } = string.Empty;
    public double Quantity { get; set; } = 1;
    public double Sum      { get; set; }
    /// <summary>Ставка позиции. Пустое значение наследует ставку стороны чека.</summary>
    public string VatType  { get; set; } = string.Empty;
}
