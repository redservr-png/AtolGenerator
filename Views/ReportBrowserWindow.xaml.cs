using System.Diagnostics;
using System.IO;
using System.Windows;
using AtolGenerator.Helpers;
using Microsoft.Web.WebView2.Core;

namespace AtolGenerator.Views;

public enum ReportPortal
{
    Taxcom,
    AtolOnline,
}

public partial class ReportBrowserWindow : Window
{
    private readonly ReportPortalOptions _options;
    private CoreWebView2DownloadOperation? _downloadOperation;

    public ReportBrowserWindow(ReportPortal portal)
    {
        _options = ReportPortalOptions.Create(portal);
        InitializeComponent();
        Title = _options.DisplayName;
        ProviderTitle.Text = _options.DisplayName;
        AddressText.Text = _options.StartUri.Host;
        LoadingText.Text = $"Открываем {_options.DisplayName}...";
    }

    public string? DownloadedReportPath { get; private set; }

    private async void OnLoaded(object sender, RoutedEventArgs e)
    {
        try
        {
            Directory.CreateDirectory(_options.DownloadDirectory);
            var profileDirectory = Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                "AtolGenerator", "WebView2", _options.ProfileName);
            var environment = await CoreWebView2Environment.CreateAsync(userDataFolder: profileDirectory);

            await Browser.EnsureCoreWebView2Async(environment);
            Browser.CoreWebView2.Settings.IsStatusBarEnabled = false;
            Browser.CoreWebView2.Settings.IsPasswordAutosaveEnabled = true;
            Browser.CoreWebView2.Settings.IsGeneralAutofillEnabled = true;
            Browser.CoreWebView2.DownloadStarting += Browser_DownloadStarting;
            Browser.CoreWebView2.NavigationStarting += Browser_NavigationStarting;
            Browser.CoreWebView2.NavigationCompleted += Browser_NavigationCompleted;
            Browser.CoreWebView2.NewWindowRequested += Browser_NewWindowRequested;

            Browser.CoreWebView2.Navigate(_options.StartUri.AbsoluteUri);
        }
        catch (Exception ex)
        {
            LoadingPanel.Visibility = Visibility.Collapsed;
            StatusText.Text = $"Не удалось запустить браузер: {ex.Message}";
            MessageBox.Show(
                "Не удалось открыть встроенный браузер. Установите Microsoft Edge WebView2 Runtime и повторите попытку.\n\n" + ex.Message,
                _options.DisplayName, MessageBoxButton.OK, MessageBoxImage.Error);
        }
    }

    private void Browser_NavigationStarting(object? sender, CoreWebView2NavigationStartingEventArgs e)
    {
        StatusText.Text = "Загружаем страницу...";
        AddressText.Text = GetDisplayAddress(e.Uri);
    }

    private void Browser_NavigationCompleted(object? sender, CoreWebView2NavigationCompletedEventArgs e)
    {
        LoadingPanel.Visibility = Visibility.Collapsed;
        UpdateNavigationButtons();
        AddressText.Text = GetDisplayAddress(Browser.Source?.AbsoluteUri);
        StatusText.Text = e.IsSuccess ? "Ожидание отчёта" : "Страница не загрузилась";
    }

    private void Browser_NewWindowRequested(object? sender, CoreWebView2NewWindowRequestedEventArgs e)
    {
        if (IsTrustedAddress(e.Uri))
        {
            e.Handled = true;
            Browser.CoreWebView2.Navigate(e.Uri);
            return;
        }

        e.Handled = true;
        try
        {
            Process.Start(new ProcessStartInfo(e.Uri) { UseShellExecute = true });
        }
        catch (Exception ex)
        {
            StatusText.Text = $"Не удалось открыть ссылку: {ex.Message}";
        }
    }

    private void Browser_DownloadStarting(object? sender, CoreWebView2DownloadStartingEventArgs e)
    {
        if (!IsTrustedAddress(Browser.Source?.AbsoluteUri))
        {
            e.Cancel = true;
            StatusText.Text = $"Загрузка отменена: открыт не сайт {_options.DisplayName}";
            return;
        }

        var suggestedName = Path.GetFileName(e.ResultFilePath);
        var extension = Path.GetExtension(suggestedName);
        if (!string.Equals(extension, _options.Extension, StringComparison.OrdinalIgnoreCase))
        {
            e.Cancel = true;
            StatusText.Text = $"Для импорта нужен отчёт {_options.FormatName}";
            MessageBox.Show(
                $"Выбранный файл не подходит для импорта. Нужен отчёт {_options.DisplayName} в формате {_options.FormatName}.",
                _options.DisplayName, MessageBoxButton.OK, MessageBoxImage.Information);
            return;
        }

        var destination = CreateUniqueReportPath(suggestedName);
        e.ResultFilePath = destination;
        e.Handled = true;

        _downloadOperation = e.DownloadOperation;
        _downloadOperation.StateChanged += DownloadOperation_StateChanged;
        StatusText.Text = $"Скачиваем {Path.GetFileName(destination)}...";
    }

    private void DownloadOperation_StateChanged(object? sender, object e)
    {
        if (sender is not CoreWebView2DownloadOperation operation) return;

        Dispatcher.Invoke(() =>
        {
            switch (operation.State)
            {
                case CoreWebView2DownloadState.Completed:
                    DownloadedReportPath = operation.ResultFilePath;
                    StatusText.Text = $"Отчёт получен: {Path.GetFileName(DownloadedReportPath)}";
                    DialogResult = true;
                    break;
                case CoreWebView2DownloadState.Interrupted:
                    StatusText.Text = $"Загрузка прервана: {operation.InterruptReason}";
                    break;
            }
        });
    }

    private void BackButton_Click(object sender, RoutedEventArgs e)
    {
        if (Browser.CanGoBack) Browser.GoBack();
    }

    private void ForwardButton_Click(object sender, RoutedEventArgs e)
    {
        if (Browser.CanGoForward) Browser.GoForward();
    }

    private void ReloadButton_Click(object sender, RoutedEventArgs e)
    {
        if (Browser.CoreWebView2 is not null) Browser.Reload();
    }

    private void HomeButton_Click(object sender, RoutedEventArgs e)
    {
        Browser.CoreWebView2?.Navigate(_options.StartUri.AbsoluteUri);
    }

    private void OpenFolderButton_Click(object sender, RoutedEventArgs e) =>
        FileHelper.OpenFolder(_options.DownloadDirectory);

    private void CloseButton_Click(object sender, RoutedEventArgs e) => Close();

    private void UpdateNavigationButtons()
    {
        BackButton.IsEnabled = Browser.CanGoBack;
        ForwardButton.IsEnabled = Browser.CanGoForward;
    }

    private bool IsTrustedAddress(string? address)
    {
        if (!Uri.TryCreate(address, UriKind.Absolute, out var uri)) return false;
        return string.Equals(uri.Scheme, Uri.UriSchemeHttps, StringComparison.OrdinalIgnoreCase) &&
               _options.TrustedDomains.Any(domain =>
                   string.Equals(uri.Host, domain, StringComparison.OrdinalIgnoreCase) ||
                   uri.Host.EndsWith('.' + domain, StringComparison.OrdinalIgnoreCase));
    }

    private string GetDisplayAddress(string? address)
    {
        if (!Uri.TryCreate(address, UriKind.Absolute, out var uri)) return _options.StartUri.Host;
        return uri.Host + uri.AbsolutePath;
    }

    private string CreateUniqueReportPath(string suggestedName)
    {
        var name = FileHelper.SafeFilename(Path.GetFileNameWithoutExtension(suggestedName));
        if (string.IsNullOrWhiteSpace(name))
            name = $"{_options.DefaultFilePrefix}_{DateTime.Now:yyyyMMdd_HHmmss}";

        var path = Path.Combine(_options.DownloadDirectory, name + _options.Extension);
        if (!File.Exists(path)) return path;

        return Path.Combine(
            _options.DownloadDirectory,
            $"{name}_{DateTime.Now:yyyyMMdd_HHmmss}{_options.Extension}");
    }

    private sealed record ReportPortalOptions(
        string DisplayName,
        Uri StartUri,
        string ProfileName,
        string Extension,
        string FormatName,
        string DownloadDirectory,
        string DefaultFilePrefix,
        IReadOnlyList<string> TrustedDomains)
    {
        public static ReportPortalOptions Create(ReportPortal portal) => portal switch
        {
            ReportPortal.AtolOnline => new ReportPortalOptions(
                "АТОЛ Online",
                new Uri("https://online.atol.ru/lk/"),
                "AtolOnline",
                ".csv",
                "CSV",
                FileHelper.AtolReportDir,
                "atol_report",
                new[] { "atol.ru", "atol.online" }),
            _ => new ReportPortalOptions(
                "Такском-Касса",
                new Uri("https://lk-ofd.taxcom.ru/#dashboard/widgets"),
                "Taxcom",
                ".xlsx",
                "XLSX",
                FileHelper.TaxcomReportDir,
                "taxcom_report",
                new[] { "taxcom.ru" }),
        };
    }
}
