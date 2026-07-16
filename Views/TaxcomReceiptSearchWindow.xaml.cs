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
    private int? _manualPeriodRequestIndex;

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
        if (!IsReceiptSearchAddress(url))
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

        if (!IsReceiptSearchAddress(Browser.Source?.AbsoluteUri))
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
                CurrentPeriodText.Text = $"Период {request.PeriodFrom:dd.MM.yyyy} - {request.PeriodTo:dd.MM.yyyy}";
                StatusText.Text = "Заполняем параметры поиска...";

                var manualPeriodConfirmed = _manualPeriodRequestIndex == index;
                var action = await FillAndSubmitAsync(request, manualPeriodConfirmed);
                if (!action.Ok)
                {
                    if (action.RequiresManualPeriod)
                    {
                        _manualPeriodRequestIndex = index;
                        StatusText.Text = action.Message;
                        SearchButton.Content = "Продолжить после ввода периода";
                        SearchButton.IsEnabled = true;
                        return;
                    }

                    _manualPeriodRequestIndex = null;
                    AddFailure(request, action.Message);
                    _nextRequestIndex = index + 1;
                    continue;
                }

                _manualPeriodRequestIndex = null;
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

    private async Task<DomActionResult> FillAndSubmitAsync(
        TaxcomReceiptSearchRequest request,
        bool manualPeriodConfirmed)
    {
        var period = $"{request.PeriodFrom:dd.MM.yyyy} - {request.PeriodTo:dd.MM.yyyy}";
        if (manualPeriodConfirmed)
        {
            StatusText.Text = $"Используем выбранный вручную период. Вводим ФПД {request.FiscalSign}...";
            return await SubmitSearchAsync(request, skipPeriodValidation: true);
        }

        var periodScript = ApplyPeriodScript
            .Replace("__DATE_FROM__", JsonSerializer.Serialize($"{request.PeriodFrom:dd.MM.yyyy}"))
            .Replace("__DATE_TO__", JsonSerializer.Serialize($"{request.PeriodTo:dd.MM.yyyy}"))
            .Replace("__PERIOD__", JsonSerializer.Serialize(period));
        var periodResult = await ExecuteAsync<DomPeriodResult>(periodScript);
        if (periodResult?.Found != true)
        {
            return new DomActionResult
            {
                Message = periodResult?.Message ?? "На странице не найдено поле периода. Выберите период вручную и продолжите поиск.",
                RequiresManualPeriod = true,
            };
        }

        await Task.Delay(350);
        var periodApplied = await IsPeriodAppliedAsync(request);
        if (!periodApplied && periodResult?.CanType == true)
        {
            StatusText.Text = "Календарь не принял период напрямую. Вводим даты с клавиатуры...";
            await TypePeriodWithDevToolsAsync(period);
            await Task.Delay(350);
            await ExecuteAsync<DomPeriodResult>(FinalizePeriodScript);
            periodApplied = await IsPeriodAppliedAsync(request);
        }

        if (!periodApplied)
        {
            return new DomActionResult
            {
                Message = $"Такском не принял период {period}. Введите этот диапазон в поле слева и нажмите «Продолжить после ввода периода».",
                RequiresManualPeriod = true,
            };
        }

        StatusText.Text = $"Период установлен: {period}. Вводим ФПД...";
        return await SubmitSearchAsync(request, skipPeriodValidation: false);
    }

    private async Task<DomActionResult> SubmitSearchAsync(
        TaxcomReceiptSearchRequest request,
        bool skipPeriodValidation)
    {
        var script = FillAndSubmitScript
            .Replace("__FISCAL_SIGN__", JsonSerializer.Serialize(request.FiscalSign.ToString()))
            .Replace("__DATE_FROM__", JsonSerializer.Serialize($"{request.PeriodFrom:dd.MM.yyyy}"))
            .Replace("__DATE_TO__", JsonSerializer.Serialize($"{request.PeriodTo:dd.MM.yyyy}"))
            .Replace("__SKIP_PERIOD_VALIDATION__", skipPeriodValidation ? "true" : "false");
        return await ExecuteAsync<DomActionResult>(script) ?? new DomActionResult
        {
            Message = "Страница поиска не вернула результат заполнения",
        };
    }

    private async Task<bool> IsPeriodAppliedAsync(TaxcomReceiptSearchRequest request)
    {
        var script = VerifyPeriodScript
            .Replace("__DATE_FROM__", JsonSerializer.Serialize($"{request.PeriodFrom:dd.MM.yyyy}"))
            .Replace("__DATE_TO__", JsonSerializer.Serialize($"{request.PeriodTo:dd.MM.yyyy}"));
        var result = await ExecuteAsync<DomPeriodResult>(script);
        return result?.Applied == true;
    }

    private async Task TypePeriodWithDevToolsAsync(string period)
    {
        if (Browser.CoreWebView2 is null) return;
        await ExecuteAsync<DomPeriodResult>(FocusPeriodScript);
        await Browser.CoreWebView2.CallDevToolsProtocolMethodAsync(
            "Input.insertText", JsonSerializer.Serialize(new { text = period }));
        const int tabKeyCode = 9;
        await Browser.CoreWebView2.CallDevToolsProtocolMethodAsync(
            "Input.dispatchKeyEvent",
            JsonSerializer.Serialize(new
            {
                type = "keyDown",
                key = "Tab",
                code = "Tab",
                windowsVirtualKeyCode = tabKeyCode,
                nativeVirtualKeyCode = tabKeyCode,
            }));
        await Browser.CoreWebView2.CallDevToolsProtocolMethodAsync(
            "Input.dispatchKeyEvent",
            JsonSerializer.Serialize(new
            {
                type = "keyUp",
                key = "Tab",
                code = "Tab",
                windowsVirtualKeyCode = tabKeyCode,
                nativeVirtualKeyCode = tabKeyCode,
            }));
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
        return uri.Host + uri.AbsolutePath + uri.Fragment;
    }

    private static bool IsReceiptSearchAddress(string? address)
    {
        if (!IsTrustedAddress(address) ||
            !Uri.TryCreate(address, UriKind.Absolute, out var uri))
            return false;

        var fragment = Uri.UnescapeDataString(uri.Fragment);
        var route = fragment
            .TrimStart('#', '/', '!')
            .TrimEnd('/');
        return string.Equals(route, "receipts", StringComparison.OrdinalIgnoreCase) ||
               route.EndsWith("/receipts", StringComparison.OrdinalIgnoreCase) ||
               fragment.Contains("returnUrl=receipts", StringComparison.OrdinalIgnoreCase);
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

    private const string ApplyPeriodScript = """
        (() => {
          const period = __PERIOD__;
          const dateFrom = __DATE_FROM__;
          const dateTo = __DATE_TO__;
          const parseDate = (value, endOfDay) => {
            const parts = value.split('.').map(Number);
            if (parts.length !== 3 || parts.some(Number.isNaN)) return null;
            return new Date(
              parts[2], parts[1] - 1, parts[0],
              endOfDay ? 23 : 0,
              endOfDay ? 59 : 0,
              endOfDay ? 59 : 0,
              endOfDay ? 999 : 0).getTime();
          };
          const findReceiptSearch = element => {
            if (!element) return null;
            const key = Object.getOwnPropertyNames(element).find(name =>
              name.startsWith('__reactFiber$') ||
              name.startsWith('__reactInternalInstance$'));
            let fiber = key ? element[key] : null;
            while (fiber) {
              const instance = fiber.stateNode;
              if (instance &&
                  typeof instance._handlerChangeDatePickerArea === 'function' &&
                  typeof instance._submitRequest === 'function' &&
                  instance.state &&
                  Object.prototype.hasOwnProperty.call(instance.state, 'fiscalSign')) {
                return instance;
              }
              fiber = fiber.return;
            }
            return null;
          };
          const periodControl = document.querySelector('.receiptsSearch__filter .date-picker-area') ||
            document.querySelector('.date-picker-area');
          if (periodControl) {
            periodControl.dataset.atolPeriodControl = 'true';
            const instance = findReceiptSearch(periodControl);
            const fromMs = parseDate(dateFrom, false);
            const toMs = parseDate(dateTo, true);
            if (instance && fromMs !== null && toMs !== null) {
              window.__atolGeneratorReceiptSearch = instance;
              instance._handlerChangeDatePickerArea(fromMs, toMs);
              const caption = periodControl.querySelector('.date-picker-area__caption');
              const value = (caption && caption.textContent || periodControl.textContent || '').trim();
              return {
                found: true,
                applied: value.includes(dateFrom) && value.includes(dateTo),
                canType: false,
                value,
                message: ''
              };
            }
            return {
              found: true,
              applied: false,
              canType: false,
              value: '',
              message: 'Не удалось подключиться к календарю Такскома'
            };
          }

          const inputs = Array.from(document.querySelectorAll('input'));
          const periodPattern = /\d{2}\.\d{2}\.\d{4}\s*[-–—]\s*\d{2}\.\d{2}\.\d{4}/;
          const fieldText = input => [
            input.id, input.name, input.placeholder, input.title,
            input.getAttribute('aria-label'),
            input.parentElement && input.parentElement.innerText,
            input.closest('.form-group, .control-group, .field, .row') &&
              input.closest('.form-group, .control-group, .field, .row').innerText
          ].filter(Boolean).join(' ').toUpperCase();
          const periodInput = inputs.find(x => periodPattern.test(x.value || '')) ||
            inputs.find(x => fieldText(x).includes('ПЕРИОД'));
          if (!periodInput) return { found: false, applied: false, canType: false, value: '', message: 'На странице не найдено поле периода. Выберите период вручную и продолжите поиск.' };

          periodInput.dataset.atolPeriodInput = 'true';
          periodInput.readOnly = false;
          periodInput.disabled = false;
          const setValue = (input, value) => {
            const setter = Object.getOwnPropertyDescriptor(HTMLInputElement.prototype, 'value').set;
            input.focus();
            input.select();
            setter.call(input, value);
            for (const name of ['input', 'keyup', 'change', 'blur']) {
              input.dispatchEvent(new Event(name, { bubbles: true, cancelable: true }));
            }
          };

          setValue(periodInput, period);
          try {
            if (window.jQuery) {
              const field = window.jQuery(periodInput);
              const picker = field.data('daterangepicker') || field.data('dateRangePicker');
              if (picker) {
                if (typeof picker.setStartDate === 'function') picker.setStartDate(dateFrom);
                if (typeof picker.setEndDate === 'function') picker.setEndDate(dateTo);
                field.trigger('apply.daterangepicker', picker);
              }
              field.val(period).trigger('input').trigger('keyup').trigger('change').trigger('blur');
            }
          } catch (_) {}
          try {
            if (window.angular) {
              const element = window.angular.element(periodInput);
              const model = element.controller('ngModel');
              if (model) {
                model.$setViewValue(period);
                model.$render();
                const scope = element.scope();
                if (scope && !scope.$root.$$phase) scope.$apply();
              }
            }
          } catch (_) {}

          periodInput.focus();
          periodInput.select();
          const value = periodInput.value || '';
          return { found: true, applied: value.includes(dateFrom) && value.includes(dateTo), canType: true, value, message: '' };
        })()
        """;

    private const string FocusPeriodScript = """
        (() => {
          const input = document.querySelector('input[data-atol-period-input="true"]');
          if (!input) return { found: false, applied: false, value: '', message: 'Поле периода потеряно' };
          input.readOnly = false;
          input.disabled = false;
          input.focus();
          input.select();
          return { found: true, applied: false, value: input.value || '', message: '' };
        })()
        """;

    private const string FinalizePeriodScript = """
        (() => {
          const input = document.querySelector('input[data-atol-period-input="true"]');
          if (!input) return { found: false, applied: false, value: '', message: 'Поле периода потеряно' };
          for (const name of ['input', 'keyup', 'change', 'blur']) {
            input.dispatchEvent(new Event(name, { bubbles: true, cancelable: true }));
          }
          try {
            if (window.jQuery) window.jQuery(input).trigger('input').trigger('keyup').trigger('change').trigger('blur');
          } catch (_) {}
          try {
            if (window.angular) {
              const element = window.angular.element(input);
              const model = element.controller('ngModel');
              if (model) {
                model.$setViewValue(input.value);
                model.$render();
                const scope = element.scope();
                if (scope && !scope.$root.$$phase) scope.$apply();
              }
            }
          } catch (_) {}
          return { found: true, applied: false, value: input.value || '', message: '' };
        })()
        """;

    private const string VerifyPeriodScript = """
        (() => {
          const dateFrom = __DATE_FROM__;
          const dateTo = __DATE_TO__;
          const periodControl = document.querySelector('[data-atol-period-control="true"]') ||
            document.querySelector('.receiptsSearch__filter .date-picker-area') ||
            document.querySelector('.date-picker-area');
          if (periodControl) {
            const caption = periodControl.querySelector('.date-picker-area__caption');
            const value = (caption && caption.textContent || periodControl.textContent || '').trim();
            return {
              found: true,
              applied: value.includes(dateFrom) && value.includes(dateTo),
              canType: false,
              value,
              message: ''
            };
          }
          const input = document.querySelector('input[data-atol-period-input="true"]');
          if (!input) return { found: false, applied: false, canType: false, value: '', message: 'Поле периода потеряно' };
          const value = input.value || '';
          return { found: true, applied: value.includes(dateFrom) && value.includes(dateTo), canType: true, value, message: '' };
        })()
        """;

    private const string FillAndSubmitScript = """
        (() => {
          const fiscalSign = __FISCAL_SIGN__;
          const dateFrom = __DATE_FROM__;
          const dateTo = __DATE_TO__;
          const skipPeriodValidation = __SKIP_PERIOD_VALIDATION__;
          const inputs = Array.from(document.querySelectorAll('input'));
          const periodControl = document.querySelector('[data-atol-period-control="true"]') ||
            document.querySelector('.receiptsSearch__filter .date-picker-area') ||
            document.querySelector('.date-picker-area');
          const periodInput = document.querySelector('input[data-atol-period-input="true"]');
          const caption = periodControl && periodControl.querySelector('.date-picker-area__caption');
          const periodValue = periodControl
            ? (caption && caption.textContent || periodControl.textContent || '')
            : (periodInput && periodInput.value || '');
          if (!skipPeriodValidation &&
              (!periodValue.includes(dateFrom) || !periodValue.includes(dateTo))) {
            return { ok: false, message: 'Период поиска не подтверждён на странице Такскома' };
          }
          const receiptSearch = window.__atolGeneratorReceiptSearch;
          if (receiptSearch &&
              typeof receiptSearch.setState === 'function' &&
              typeof receiptSearch._submitRequest === 'function') {
            receiptSearch.setState({ fiscalSign }, () => receiptSearch._submitRequest());
            return { ok: true, message: '' };
          }
          const setValue = (input, value) => {
            const setter = Object.getOwnPropertyDescriptor(HTMLInputElement.prototype, 'value').set;
            setter.call(input, value);
            for (const name of ['input', 'keyup', 'change', 'blur']) {
              input.dispatchEvent(new Event(name, { bubbles: true, cancelable: true }));
            }
            try {
              if (window.jQuery) window.jQuery(input).val(value).trigger('input').trigger('keyup').trigger('change').trigger('blur');
            } catch (_) {}
          };
          const fpInput = inputs.find(x => (x.placeholder || '').toUpperCase().includes('ФПД'));
          if (!fpInput) return { ok: false, message: 'На странице не найдено поле ФПД' };
          setValue(fpInput, fiscalSign);
          const controls = Array.from(document.querySelectorAll('button, input[type="submit"], a'));
          const searchButton = controls.find(x => ((x.textContent || x.value || '').trim().toUpperCase() === 'ПОИСК ЧЕКА'));
          if (!searchButton) return { ok: false, message: 'На странице не найдена кнопка поиска' };
          if (searchButton.disabled) return { ok: false, message: 'Кнопка поиска недоступна: проверьте период' };
          searchButton.click();
          return { ok: true, message: '' };
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
        public bool RequiresManualPeriod { get; init; }
        public string Message { get; init; } = string.Empty;
    }

    private sealed class DomPeriodResult
    {
        public bool Found { get; init; }
        public bool Applied { get; init; }
        public bool CanType { get; init; }
        public string Value { get; init; } = string.Empty;
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
