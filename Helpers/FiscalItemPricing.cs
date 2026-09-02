namespace AtolGenerator.Helpers;

/// <summary>
/// АТОЛ требует точного равенства price × quantity = sum.
/// Количество сохраняем из 1С; сумму строки берём как price × quantity,
/// где price = Round(сумма1С / quantity, 2). Отличие от суммы в 1С — копейки округления.
/// </summary>
public static class FiscalItemPricing
{
    public static (double Price, double Quantity, double Sum) Align(double quantity, double sum)
    {
        var sourceSum = Math.Round(sum, 2);
        if (sourceSum <= 0)
            return (0, quantity > 0 ? quantity : 1, 0);

        var qty = quantity > 0 ? quantity : 1;
        var price = Math.Round(sourceSum / qty, 2);
        var alignedSum = Math.Round(price * qty, 2);
        return (price, qty, alignedSum);
    }
}
