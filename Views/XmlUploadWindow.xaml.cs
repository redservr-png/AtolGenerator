using System.Diagnostics;
using System.IO;
using System.Windows;
using AtolGenerator.Helpers;
using AtolGenerator.Services;
using Microsoft.Web.WebView2.Core;
using Shape = System.Windows.Shapes.Shape;

namespace AtolGenerator.Views;

public partial class XmlUploadWindow : Window
{
    private static readonly Uri HomeUri = new("https://online.atol.ru/lk/");
    private static readonly string[] TrustedDomains = { "atol.ru", "atol.online" };

    private readonly CorrectionPunchPlan _plan;
    private bool _uploadDone;

    public XmlUploadWindow(CorrectionPunchPlan plan)
    {
        _plan = plan;
        InitializeComponent();
        FileSummaryText.Text = BuildFileSummary(plan);
        if (plan.ReceiptXmlPaths.Count > 0)
        {
            WarningText.Visibility = Visibility.Visible;
            WarningText.Text =
                "В кабинет кладите только файл(ы) коррекций. XML возвратов/приходов уже ушли через API.";
        }
    }

    private async void OnLoaded(object sender, RoutedEventArgs e)
    {
        try
        {
            var profileDirectory = Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                "AtolGenerator", "WebView2", "AtolOnline");
            var environment = await CoreWebView2Environment.CreateAsync(userDataFolder: profileDirectory);

            await Browser.EnsureCoreWebView2Async(environment);
            Browser.CoreWebView2.Settings.IsStatusBarEnabled = false;
            Browser.CoreWebView2.Settings.IsPasswordAutosaveEnabled = true;
            Browser.CoreWebView2.Settings.IsGeneralAutofillEnabled = true;
            Browser.CoreWebView2.NavigationStarting += Browser_NavigationStarting;
            Browser.CoreWebView2.NavigationCompleted += Browser_NavigationCompleted;
            Browser.CoreWebView2.NewWindowRequested += Browser_NewWindowRequested;

            Browser.CoreWebView2.Navigate(HomeUri.AbsoluteUri);
        }
        catch (Exception ex)
        {
            LoadingPanel.Visibility = Visibility.Collapsed;
            SetStatus($"Не удалось запустить браузер: {ex.Message}", isError: true);
            MessageBox.Show(
                WebView2ErrorHelper.GetStartupMessage(ex),
                "Загрузка XML", MessageBoxButton.OK, MessageBoxImage.Error);
        }
    }

    private void Browser_NavigationStarting(object? sender, CoreWebView2NavigationStartingEventArgs e)
    {
        if (!_uploadDone)
            SetStatus("Загружаем страницу...");
        AddressText.Text = GetDisplayAddress(e.Uri);
    }

    private void Browser_NavigationCompleted(object? sender, CoreWebView2NavigationCompletedEventArgs e)
    {
        LoadingPanel.Visibility = Visibility.Collapsed;
        UpdateNavigationButtons();
        AddressText.Text = GetDisplayAddress(Browser.Source?.AbsoluteUri);
        if (_uploadDone) return;

        SetStatus(e.IsSuccess
            ? "Откройте «Загрузка чеков из файлов XML (1.05/1.2)» и укажите файл коррекции."
            : "Страница не загрузилась",
            isError: !e.IsSuccess);
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
            SetStatus($"Не удалось открыть ссылку: {ex.Message}", isError: true);
        }
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

    private void HomeButton_Click(object sender, RoutedEventArgs e) =>
        Browser.CoreWebView2?.Navigate(HomeUri.AbsoluteUri);

    private void CopyPathButton_Click(object sender, RoutedEventArgs e)
    {
        var path = _plan.CorrectionXmlPaths.FirstOrDefault();
        if (string.IsNullOrWhiteSpace(path))
        {
            SetStatus("Нет пути к XML коррекции", isError: true);
            return;
        }

        Clipboard.SetText(path);
        SetStatus(_plan.CorrectionXmlPaths.Count == 1
            ? "Путь к XML скопирован"
            : $"Скопирован первый из {_plan.CorrectionXmlPaths.Count} файлов. Остальные в той же папке.");
    }

    private void OpenFolderButton_Click(object sender, RoutedEventArgs e)
    {
        var path = _plan.CorrectionXmlPaths.FirstOrDefault();
        if (!string.IsNullOrWhiteSpace(path))
            FileHelper.RevealInExplorer(path);
        else
            FileHelper.OpenFolder(FileHelper.OutputDir);
    }

    private void MarkDoneButton_Click(object sender, RoutedEventArgs e)
    {
        _uploadDone = true;
        WarningText.Visibility = Visibility.Collapsed;
        MarkDoneButton.IsEnabled = false;
        MarkDoneButton.Content = "Готово";
        SetStatus("XML принят АТОЛ. Можно закрыть окно и перейти к сверке / записи в 1С.", isSuccess: true);
    }

    private void CloseButton_Click(object sender, RoutedEventArgs e) => Close();

    private void SetStatus(string text, bool isError = false, bool isSuccess = false)
    {
        StatusText.Text = text;
        if (isSuccess)
        {
            StatusText.SetResourceReference(ForegroundProperty, "BrushGreen");
            StatusIndicator.SetResourceReference(Shape.FillProperty, "BrushGreen");
            return;
        }

        if (isError)
        {
            StatusText.SetResourceReference(ForegroundProperty, "BrushRed");
            StatusIndicator.SetResourceReference(Shape.FillProperty, "BrushRed");
            return;
        }

        StatusText.SetResourceReference(ForegroundProperty, "BrushText2");
        StatusIndicator.SetResourceReference(Shape.FillProperty, "BrushAccent");
    }

    private void UpdateNavigationButtons()
    {
        BackButton.IsEnabled = Browser.CanGoBack;
        ForwardButton.IsEnabled = Browser.CanGoForward;
    }

    private static bool IsTrustedAddress(string? address)
    {
        if (!Uri.TryCreate(address, UriKind.Absolute, out var uri)) return false;
        return string.Equals(uri.Scheme, Uri.UriSchemeHttps, StringComparison.OrdinalIgnoreCase) &&
               TrustedDomains.Any(domain =>
                   string.Equals(uri.Host, domain, StringComparison.OrdinalIgnoreCase) ||
                   uri.Host.EndsWith('.' + domain, StringComparison.OrdinalIgnoreCase));
    }

    private static string GetDisplayAddress(string? address)
    {
        if (!Uri.TryCreate(address, UriKind.Absolute, out var uri)) return HomeUri.Host;
        return uri.Host + uri.AbsolutePath;
    }

    private static string BuildFileSummary(CorrectionPunchPlan plan)
    {
        if (plan.CorrectionXmlPaths.Count == 0)
            return "Нет XML коррекций";

        var names = plan.CorrectionXmlPaths.Select(Path.GetFileName);
        return plan.CorrectionXmlPaths.Count == 1
            ? $"Файл: {names.First()}"
            : $"Файлов коррекций: {plan.CorrectionXmlPaths.Count} — {string.Join(", ", names)}";
    }
}
