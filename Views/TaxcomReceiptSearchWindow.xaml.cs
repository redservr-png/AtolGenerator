using System.IO;
using System.Text.Json;
using System.Windows;
using AtolGenerator.Models;
using AtolGenerator.Services;
using Microsoft.Web.WebView2.Core;

namespace AtolGenerator.Views;

public partial class TaxcomReceiptSearchWindow : Window
{
    private const string SearchUrl = "https://lk-ofd.taxcom.ru/#receipts";
    private static readonly JsonSerializerOptions JsonOptions = new()
    {
        PropertyNameCaseInsensitive = true,
    };

    private readonly IReadOnlyList<TaxcomReceiptSearchRequest> _requests;
    private bool _isRunning;
    private bool _isComplete;
    private int _nextRequestIndex;

    public TaxcomReceiptSearchWindow(IReadOnlyList<TaxcomReceiptSearchRequest> requests)
    {
        _requests = requests;
        InitializeComponent();
        ProgressText.Text = $"Подготовлено запросов: {_requests.Count}";
        SearchButton.IsEnabled = false;
    }

    public List<OfdReportRow> FoundRows { get; } = new();
    public List<TaxcomReceiptSearchFailure> Failures { get; } = new();

    private async void OnLoaded(object sender, RoutedEventArgs e)
    {
        try
        {
            var profileDirectory = Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                "AtolGenerator", "WebView2", "Taxcom");
            var environment = await CoreWebView2Environment.CreateAsync(userDataFolder: profileDirectory);
            await Browser.EnsureCoreWebView2Async(environment);
            Browser.CoreWebView2.Settings.IsStatusBarEnabled = false;
            Browser.CoreWebView2.Settings.IsPasswordAutosaveEnabled = true;
            Browser.CoreWebView2.Settings.IsGeneralAutofillEnabled = true;
            Browser.CoreWebView2.NavigationStarting += Browser_NavigationStarting;
            Browser.CoreWebView2.NavigationCompleted += Browser_NavigationCompleted;
            Browser.CoreWebView2.NewWindowRequested += Browser_NewWindowRequested;
            Browser.CoreWebView2.Navigate(SearchUrl);
        }
        catch (Exception ex)
        {
            StatusText.Text = $"Не удалось открыть встроенный браузер: {ex.Message}";
            SearchButton.IsEnabled = false;
        }
    }

    private void Browser_NavigationStarting(object? sender, CoreWebView2NavigationStartingEventArgs e)
    {
        AddressText.Text = GetDisplayAddress(e.Uri);
        StatusText.Text = "Открываем Такском-Кассу...";
    }

    private async void Browser_NavigationCompleted(object? sender, CoreWebView2NavigationCompletedEventArgs e)
    {
        AddressText.Text = GetDisplayAddress(Browser.Source?.AbsoluteUri);
        if (!e.IsSuccess)
        {
            StatusText.Text = "Страница Такскома не загрузилась";
            SearchButton.IsEnabled = true;
            return;
        }

        if (await IsLoginPageAsync())
        {
            StatusText.Text = "Войдите в Такском. После авторизации поиск продолжится автоматически.";
            SearchButton.Content = "Продолжить после входа";
            SearchButton.IsEnabled = true;
            return;
        }

        var url = Browser.Source?.AbsoluteUri ?? string.Empty;
        if (!url.Contains("#receipts", StringComparison.OrdinalIgnoreCase))
        {
            Browser.CoreWebView2.Navigate(SearchUrl);
            return;
        }

        SearchButton.IsEnabled = true;
        await RunSearchAsync();
    }

    private void Browser_NewWindowRequested(object? sender, CoreWebView2NewWindowRequestedEventArgs e)
    {
        if (!IsTrustedAddress(e.Uri)) return;
        e.Handled = true;
        Browser.CoreWebView2.Navigate(e.Uri);
    }

    private async void SearchButton_Click(object sender, RoutedEventArgs e)
    {
        if (_isComplete)
        {
            DialogResult = true;
            return;
        }

        if (await IsLoginPageAsync())
        {
            StatusText.Text = "Сначала завершите вход в Такском в левой части окна.";
            return;
        }

        if (!(Browser.Source?.AbsoluteUri ?? string.Empty).Contains("#receipts", StringComparison.OrdinalIgnoreCase))
        {
            Browser.CoreWebView2.Navigate(SearchUrl);
            return;
        }

        await RunSearchAsync();
    }

    private async Task RunSearchAsync()
    {
        if (_isRunning || _isComplete || Browser.CoreWebView2 is null) return;
        _isRunning = true;
        SearchButton.IsEnabled = false;

        try
        {
            StatusText.Text = "Готовим страницу поиска...";
            if (!await WaitForSearchPageAsync())
            {
                StatusText.Text = "Страница поиска ещё не готова. Дождитесь формы и нажмите «Продолжить поиск».";
                SearchButton.IsEnabled = true;
                return;
            }

            for (var index = _nextRequestIndex; index < _requests.Count; index++)
            {
                var request = _requests[index];
                ProgressText.Text = $"Запрос {index + 1} из {_requests.Count} · найдено {FoundRows.Count}";
                CurrentDocumentText.Text = request.DocumentLabel;
                CurrentFiscalSignText.Text = $"ФПД {request.FiscalSign}";
                StatusText.Text = "Заполняем параметры поиска...";

                var action = await FillAndSubmitAsync(request);
                if (!action.Ok)
                {
                    AddFailure(request, action.Message);
                    _nextRequestIndex = index + 1;
                    continue;
                }

                StatusText.Text = "Ждём карточку чека...";
                var extraction = await WaitForReceiptAsync(request.FiscalSign);
                if (!string.Equals(extraction.State, "found", StringComparison.OrdinalIgnoreCase))
                {
                    AddFailure(request, extraction.Message.Length > 0
                        ? extraction.Message
                        : "Чек не найден в выбранном периоде");
                    _nextRequestIndex = index + 1;
                    continue;
                }

                if (!TaxcomReceiptParser.TryParse(
                        extraction.Text, request.FiscalSign, out var row, out var parseError))
                {
                    AddFailure(request, parseError);
                    _nextRequestIndex = index + 1;
                    continue;
                }

                FoundRows.Add(row);
                ResultList.Items.Add($"Найден · ФП {request.FiscalSign} · ФД {row.FiscalDocument} · {row.Amount:N2} ₽");
                _nextRequestIndex = index + 1;
            }

            _isComplete = true;
            ProgressText.Text = $"Завершено · найдено {FoundRows.Count} из {_requests.Count}";
            StatusText.Text = Failures.Count == 0
                ? "Все исходные чеки найдены и будут сохранены в локальный архив."
                : $"Не найдено или не прочитано: {Failures.Count}. Найденные чеки будут сохранены.";
            SearchButton.Content = "Применить результаты";
            SearchButton.IsEnabled = true;

            if (Failures.Count == 0)
            {
                await Task.Delay(600);
                DialogResult = true;
            }
        }
        catch (Exception ex)
        {
            StatusText.Text = $"Поиск остановлен: {ex.Message}";
            SearchButton.Content = "Продолжить поиск";
            SearchButton.IsEnabled = true;
        }
        finally
        {
            _isRunning = false;
        }
    }

    private async Task<DomActionResult> FillAndSubmitAsync(TaxcomReceiptSearchRequest request)
    {
        var script = FillAndSubmitScript
            .Replace("__FISCAL_SIGN__", JsonSerializer.Serialize(request.FiscalSign.ToString()))
            .Replace("__PERIOD__", JsonSerializer.Serialize(
                $"{request.PeriodFrom:dd.MM.yyyy} - {request.PeriodTo:dd.MM.yyyy}"));
        return await ExecuteAsync<DomActionResult>(script) ?? new DomActionResult
        {
            Message = "Страница поиска не вернула результат заполнения",
        };
    }

    private async Task<DomReceiptResult> WaitForReceiptAsync(long fiscalSign)
    {
        var script = ExtractReceiptScript.Replace(
            "__FISCAL_SIGN__", JsonSerializer.Serialize(fiscalSign.ToString()));
        for (var attempt = 0; attempt < 30; attempt++)
        {
            await Task.Delay(500);
            var result = await ExecuteAsync<DomReceiptResult>(script);
            if (result is not null &&
                string.Equals(result.State, "found", StringComparison.OrdinalIgnoreCase))
                return result;
            if (attempt >= 3 && result is not null &&
                string.Equals(result.State, "not_found", StringComparison.OrdinalIgnoreCase))
                return result;
        }

        return new DomReceiptResult
        {
            State = "not_found",
            Message = "Такском не показал карточку чека за 15 секунд",
        };
    }

    private async Task<bool> WaitForSearchPageAsync()
    {
        for (var attempt = 0; attempt < 20; attempt++)
        {
            var result = await ExecuteAsync<DomSearchPageResult>(SearchPageScript);
            if (result?.Ready == true) return true;
            await Task.Delay(500);
        }

        return false;
    }

    private async Task<bool> IsLoginPageAsync()
    {
        if (Browser.CoreWebView2 is null) return true;
        var result = await ExecuteAsync<DomLoginResult>(LoginPageScript);
        return result?.IsLogin == true;
    }

    private async Task<T?> ExecuteAsync<T>(string script) where T : class
    {
        if (Browser.CoreWebView2 is null) return null;
        var json = await Browser.CoreWebView2.ExecuteScriptAsync(script);
        return JsonSerializer.Deserialize<T>(json, JsonOptions);
    }

    private void AddFailure(TaxcomReceiptSearchRequest request, string message)
    {
        Failures.Add(new TaxcomReceiptSearchFailure
        {
            FiscalSign = request.FiscalSign,
            DocumentLabel = request.DocumentLabel,
            Message = message,
        });
        ResultList.Items.Add($"Не найден · ФП {request.FiscalSign} · {message}");
    }

    private void CloseButton_Click(object sender, RoutedEventArgs e)
    {
        DialogResult = FoundRows.Count > 0;
    }

    private static string GetDisplayAddress(string? address)
    {
        if (!Uri.TryCreate(address, UriKind.Absolute, out var uri)) return "lk-ofd.taxcom.ru";
        return uri.Host + uri.Fragment;
    }

    private static bool IsTrustedAddress(string? address)
    {
        if (!Uri.TryCreate(address, UriKind.Absolute, out var uri)) return false;
        return string.Equals(uri.Scheme, Uri.UriSchemeHttps, StringComparison.OrdinalIgnoreCase) &&
               (string.Equals(uri.Host, "taxcom.ru", StringComparison.OrdinalIgnoreCase) ||
                uri.Host.EndsWith(".taxcom.ru", StringComparison.OrdinalIgnoreCase));
    }

    private const string LoginPageScript = """
        (() => {
          const inputs = Array.from(document.querySelectorAll('input'));
          const isLogin = inputs.some(x => (x.placeholder || '').toLowerCase().includes('парол'));
          return { isLogin };
        })()
        """;

    private const string SearchPageScript = """
        (() => {
          const inputs = Array.from(document.querySelectorAll('input'));
          const hasFiscalSign = inputs.some(x => (x.placeholder || '').toUpperCase().includes('ФПД'));
          const controls = Array.from(document.querySelectorAll('button, input[type="submit"], a'));
          const hasSearch = controls.some(x => ((x.textContent || x.value || '').trim().toUpperCase() === 'ПОИСК ЧЕКА'));
          return { ready: hasFiscalSign && hasSearch };
        })()
        """;

    private const string FillAndSubmitScript = """
        (() => {
          const fiscalSign = __FISCAL_SIGN__;
          const period = __PERIOD__;
          const inputs = Array.from(document.querySelectorAll('input'));
          const setValue = (input, value) => {
            const setter = Object.getOwnPropertyDescriptor(HTMLInputElement.prototype, 'value').set;
            setter.call(input, value);
            for (const name of ['input', 'change', 'keyup', 'blur']) {
              input.dispatchEvent(new Event(name, { bubbles: true }));
            }
          };
          const fpInput = inputs.find(x => (x.placeholder || '').toUpperCase().includes('ФПД'));
          if (!fpInput) return { ok: false, message: 'На странице не найдено поле ФПД' };
          const periodInput = inputs.find(x => /^\s*\d{2}\.\d{2}\.\d{4}\s*-\s*\d{2}\.\d{2}\.\d{4}\s*$/.test(x.value || '')) ||
            inputs.find(x => ((x.parentElement && x.parentElement.innerText) || '').toUpperCase().includes('ПЕРИОД'));
          if (periodInput) setValue(periodInput, period);
          setValue(fpInput, fiscalSign);
          const controls = Array.from(document.querySelectorAll('button, input[type="submit"], a'));
          const searchButton = controls.find(x => ((x.textContent || x.value || '').trim().toUpperCase() === 'ПОИСК ЧЕКА'));
          if (!searchButton) return { ok: false, message: 'На странице не найдена кнопка поиска' };
          if (searchButton.disabled) return { ok: false, message: 'Кнопка поиска недоступна: проверьте период' };
          searchButton.click();
          return { ok: true, message: periodInput ? '' : 'Использован период, уже выбранный в Такскоме' };
        })()
        """;

    private const string ExtractReceiptScript = """
        (() => {
          const fiscalSign = __FISCAL_SIGN__;
          const elements = Array.from(document.querySelectorAll('main, article, section, div'));
          const candidates = elements
            .map(x => ({ text: (x.innerText || '').trim() }))
            .filter(x => x.text.includes(fiscalSign) && x.text.toUpperCase().includes('КАССОВЫЙ ЧЕК') && x.text.toUpperCase().includes('ИТОГО'))
            .filter(x => x.text.length > 100 && x.text.length < 15000)
            .sort((a, b) => a.text.length - b.text.length);
          if (candidates.length > 0) return { state: 'found', text: candidates[0].text, message: '' };
          const pageText = (document.body.innerText || '').toLowerCase();
          const notFound = pageText.includes('ничего не найдено') || pageText.includes('чеки не найдены') || pageText.includes('результаты не найдены');
          return { state: notFound ? 'not_found' : 'pending', text: '', message: notFound ? 'Чек не найден в Такскоме' : '' };
        })()
        """;

    private sealed class DomLoginResult
    {
        public bool IsLogin { get; init; }
    }

    private sealed class DomActionResult
    {
        public bool Ok { get; init; }
        public string Message { get; init; } = string.Empty;
    }

    private sealed class DomSearchPageResult
    {
        public bool Ready { get; init; }
    }

    private sealed class DomReceiptResult
    {
        public string State { get; init; } = "pending";
        public string Text { get; init; } = string.Empty;
        public string Message { get; init; } = string.Empty;
    }
}
