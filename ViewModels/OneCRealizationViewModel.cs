using AtolGenerator.Services;
using AtolGenerator.ViewModels;

namespace AtolGenerator.ViewModels;

public class OneCRealizationViewModel : BaseViewModel
{
    public OneCRealization Source { get; }

    private bool _isSelected = true;
    public bool IsSelected
    {
        get => _isSelected;
        set => Set(ref _isSelected, value);
    }

    public string DocNumber    => Source.DocNumber;
    public string DocDate      => Source.DocDate;
    public string OrderNumber  => Source.OrderNumber;
    public string OrderDate    => Source.OrderDate;
    public string CustomerName => Source.CustomerName;
    public double Amount       => Source.Amount;
    public bool   IsService    => Source.IsService;
    public string City         => Source.City;
    public bool   HasCheck     => Source.HasCheck;
    public string CheckDate    => Source.CheckDate;
    public string CheckNumber  => Source.CheckNumber;

    public OneCRealizationViewModel(OneCRealization r) => Source = r;
}
