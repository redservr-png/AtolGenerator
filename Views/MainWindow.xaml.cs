using System.Windows;
using AtolGenerator.ViewModels;

namespace AtolGenerator.Views;

public partial class MainWindow : Window
{
    public MainWindow() => InitializeComponent();

    private async void OnLoaded(object sender, RoutedEventArgs e)
    {
        if (DataContext is MainViewModel viewModel)
            await viewModel.InitializeAsync();
    }
}
