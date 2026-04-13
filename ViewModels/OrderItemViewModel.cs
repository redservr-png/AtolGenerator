using AtolGenerator.Models;

namespace AtolGenerator.ViewModels;

public class OrderItemViewModel : BaseViewModel
{
    private string _name     = string.Empty;
    private double _quantity = 1;
    private double _sum;

    public string Name
    {
        get => _name;
        set { Set(ref _name, value); OnPropertyChanged(nameof(Price)); }
    }

    public double Quantity
    {
        get => _quantity;
        set { Set(ref _quantity, value); OnPropertyChanged(nameof(Price)); }
    }

    public double Sum
    {
        get => _sum;
        set { Set(ref _sum, value); OnPropertyChanged(nameof(Price)); }
    }

    public double Price => Quantity > 0 ? Math.Round(Sum / Quantity, 2) : 0;

    public OrderItem ToModel() => new() { Name = Name, Quantity = Quantity, Sum = Sum };
}
