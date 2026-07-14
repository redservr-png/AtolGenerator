using System.ComponentModel;
using System.Windows;
using AtolGenerator.ViewModels;

namespace AtolGenerator.Views;

public partial class SettingsWindow : Window
{
    private readonly SettingsViewModel _viewModel;
    private bool _completed;

    public SettingsWindow()
    {
        InitializeComponent();
        _viewModel = new SettingsViewModel();
        _viewModel.RequestClose += Complete;
        DataContext = _viewModel;
    }

    private void Complete(bool saved)
    {
        _completed = true;
        try
        {
            DialogResult = saved;
        }
        catch (InvalidOperationException)
        {
            // Поддерживаем и немодальный запуск окна (например, в UI smoke-тесте).
            Close();
        }
    }

    protected override void OnClosing(CancelEventArgs e)
    {
        if (!_completed)
            _viewModel.CancelPreview();
        base.OnClosing(e);
    }
}
