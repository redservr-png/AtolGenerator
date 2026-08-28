using System.Windows;
using AtolGenerator.ViewModels;

namespace AtolGenerator.Views;

public partial class OneCRealizationsWindow : Window
{
    public OneCRealizationsWindow()
    {
        InitializeComponent();
        Loaded += OnLoaded;
        Closed += OnClosed;
    }

    private void OnLoaded(object sender, RoutedEventArgs e)
    {
        if (DataContext is MainViewModel viewModel)
            viewModel.SetOneCRealizationsCompact(false);
    }

    private void OnClosed(object? sender, EventArgs e)
    {
        if (DataContext is MainViewModel viewModel)
            viewModel.SetOneCRealizationsCompact(true);
    }
}
