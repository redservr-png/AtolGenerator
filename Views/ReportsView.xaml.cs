using System.Windows;
using System.Windows.Controls;
using AtolGenerator.Views.Converters;

namespace AtolGenerator.Views;

public partial class ReportsView : UserControl
{
    public ReportsView()
    {
        InitializeComponent();
        DataContextChanged += (_, _) =>
        {
            if (TryFindResource("Proxy") is BindingProxy proxy)
                proxy.Data = DataContext;
        };
    }
}
