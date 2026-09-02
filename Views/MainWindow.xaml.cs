using System.Windows;
using System.Windows.Input;
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

    private void OnResultDoubleClick(object sender, MouseButtonEventArgs e)
    {
        if (DataContext is MainViewModel viewModel)
            viewModel.OpenReceiptPreviewWindow();
    }
}
