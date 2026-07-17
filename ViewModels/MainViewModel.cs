using System.Collections.ObjectModel;
using System.IO;
using System.Windows;
using System.Windows.Input;
using AtolGenerator.Constants;
using AtolGenerator.Helpers;
using AtolGenerator.Models;
using AtolGenerator.Services;
using Microsoft.Win32;

namespace AtolGenerator.ViewModels;

public class MainViewModel : BaseViewModel
{
    private string _activeWorkspace = "main";
    private bool _startupInitialized;
    private Views.CorrectionWorkWindow? _correctionWorkWindow;

    public ReportsViewModel Reports { get; } = new();
    public ObsidianCasesViewModel ObsidianCases { get; }
    public CorrectionWorkViewModel CorrectionWork { get; }
    public bool ShowMainWorkspace => _activeWorkspace == "main";
    public bool ShowReportsWorkspace => _activeWorkspace == "reports";
    public bool ShowObsidianWorkspace => _activeWorkspace == "obsidian";

    // ── Tab ──────────────────────────────────────────────────────────────────
    private string _tab = "payment";
    public string Tab
    {
        get => _tab;
        set
        {
            Set(ref _tab, value);
            OnPropertyChanged(nameof(IsPaymentTab));
            OnPropertyChanged(nameof(IsRealizationTab));
            OnPropertyChanged(nameof(ShowPaymentType));
            OnPropertyChanged(nameof(ShowItemsSection));
            OnPropertyChanged(nameof(ShowBuyRefundOption));
            OnPropertyChanged(nameof(OneCPanelVisible));
            // Reset to sell on tab switch
            if (CheckType is "buy_refund" && value == "payment")
                CheckType = "sell";
        }
    }
    public bool IsPaymentTab     => Tab == "payment";
    public bool IsRealizationTab => Tab == "realization";

    // ── Check Type ───────────────────────────────────────────────────────────
    private string _checkType = "sell";
    public string CheckType
    {
        get => _checkType;
        set
        {
            Set(ref _checkType, value);
            OnPropertyChanged(nameof(IsCorrection));
            OnPropertyChanged(nameof(ShowCorrectionBox));
            OnPropertyChanged(nameof(CanPunchOrdersViaAtol));
            OnPropertyChanged(nameof(CorrectionPunchHint));
            OnPropertyChanged(nameof(ShowCorrectionPunchHint));
            OnPropertyChanged(nameof(ShowPaymentType));
            OnPropertyChanged(nameof(ShowItemsSection));
            OnPropertyChanged(nameof(CanPunchSelectedRealizationsViaAtol));
        }
    }
    public bool IsCorrection         => CheckType is "sell_correction" or "buy_correction";
    public bool ShowCorrectionBox    => IsCorrection;

    /// <summary>В списке есть исправления, которые разрешено формировать только через XML.</summary>
    public bool HasXmlOnlyCorrection =>
        Orders.Any(IsXmlOnlyOrder);

    /// <summary>Количество таких «XML-только» коррекций в списке.</summary>
    public int XmlOnlyCorrectionCount =>
        Orders.Count(IsXmlOnlyOrder);

    // Пробитие через API: коррекции не поддерживаются (ошибка 31)
    public bool CanPunchOrdersViaAtol =>
        OrderCount > 0 && !IsCorrection && !HasXmlOnlyCorrection;

    private static bool IsXmlOnlyOrder(OrderEntry order) =>
        order.IsCorrection ||
        !string.IsNullOrWhiteSpace(order.ObsidianCaseId) ||
        order.CorrectionScenario.RequiresXmlOnly();

    public string CorrectionPunchHint
    {
        get
        {
            if (IsCorrection)
                return "Коррекция не поддерживается через АТОЛ API.\nИспользуйте сформированный XML-файл.";
            if (HasXmlOnlyCorrection)
                return $"В списке есть {XmlOnlyCorrectionCount} коррекций, требующих XML.\n" +
                       "Пробитие через API недоступно — используйте «Сформировать XML».";
            return string.Empty;
        }
    }

    /// <summary>Показывать ли подсказку про невозможность пробития через API.</summary>
    public bool ShowCorrectionPunchHint => IsCorrection || HasXmlOnlyCorrection;
    public bool ShowPaymentType   => IsPaymentTab;  // на вкладке реализации тип оплаты всегда 14 (аванс)
    public bool ShowBuyRefundOption => IsRealizationTab;
    public bool ShowItemsSection  => IsRealizationTab && !IsCorrection;

    // ── Кассир ───────────────────────────────────────────────────────────────
    public ObservableCollection<CashierInfo> AvailableCashiers { get; } = new();

    private CashierInfo _selectedCashier = AppConstants.DefaultCashier;
    public CashierInfo SelectedCashier
    {
        get => _selectedCashier;
        set
        {
            Set(ref _selectedCashier, value);
            OnPropertyChanged(nameof(SelectedCashierShort));
        }
    }
    public string SelectedCashierShort => SelectedCashier?.ShortName ?? AppConstants.CashierShort;

    // ── Merge XML ────────────────────────────────────────────────────────────
    private bool _mergeXml = true;
    public bool MergeXml { get => _mergeXml; set => Set(ref _mergeXml, value); }

    // ── Payment Type ─────────────────────────────────────────────────────────
    private string _paymentType = "card";
    public string PaymentType
    {
        get => _paymentType;
        set => Set(ref _paymentType, value);
    }

    // ── Correction fields ────────────────────────────────────────────────────
    private string _correctionDate   = string.Empty;
    private string _correctionNumber = string.Empty;
    public string CorrectionDate
    {
        get => _correctionDate;
        set => Set(ref _correctionDate, value);
    }
    public string CorrectionNumber
    {
        get => _correctionNumber;
        set => Set(ref _correctionNumber, value);
    }

    // ── Bulk input ───────────────────────────────────────────────────────────
    private string _bulkText = string.Empty;
    public string BulkText
    {
        get => _bulkText;
        set => Set(ref _bulkText, value);
    }

    // ── Single order fields ──────────────────────────────────────────────────
    private string _singleOrderNum  = string.Empty;
    private string _singleOrderDate = string.Empty;
    private string _singleAmount    = string.Empty;
    private string _singleCustomer  = string.Empty;
    public string SingleOrderNum  { get => _singleOrderNum;  set => Set(ref _singleOrderNum,  value); }
    public string SingleOrderDate { get => _singleOrderDate; set => Set(ref _singleOrderDate, value); }
    public string SingleAmount    { get => _singleAmount;    set => Set(ref _singleAmount,    value); }
    public string SingleCustomer  { get => _singleCustomer;  set => Set(ref _singleCustomer,  value); }

    // ── Agent / service ──────────────────────────────────────────────────────
    private bool   _isServiceProvider;
    private string _selectedServiceType = string.Empty;
    private string _selectedCity        = string.Empty;

    public bool IsServiceProvider
    {
        get => _isServiceProvider;
        set
        {
            Set(ref _isServiceProvider, value);
            OnPropertyChanged(nameof(ShowAgentPanel));
            if (!value) SelectedAgent = null;
        }
    }
    public bool ShowAgentPanel => _isServiceProvider;

    public string SelectedServiceType
    {
        get => _selectedServiceType;
        set
        {
            Set(ref _selectedServiceType, value);
            RefreshCities();
            SelectedCity = string.Empty;
            SelectedAgent = null;
        }
    }

    public string SelectedCity
    {
        get => _selectedCity;
        set
        {
            Set(ref _selectedCity, value);
            PickAgent();
        }
    }

    private ServiceProvider? _selectedAgent;
    public ServiceProvider? SelectedAgent
    {
        get => _selectedAgent;
        set
        {
            Set(ref _selectedAgent, value);
            OnPropertyChanged(nameof(AgentInfoText));
        }
    }

    public string AgentInfoText => SelectedAgent is { } a
        ? $"{a.Name}\nИНН: {a.Inn}   Тел: {a.Phone}"
        : string.Empty;

    public ObservableCollection<string> AvailableCities { get; } = new();

    // ── Order items (for realization) ────────────────────────────────────────
    public ObservableCollection<OrderItemViewModel> CurrentItems { get; } = new();

    // ── Order list ───────────────────────────────────────────────────────────
    public ObservableCollection<OrderEntry> Orders { get; } = new();

    // ── 1C staging table ─────────────────────────────────────────────────────
    public ObservableCollection<OneCRealizationViewModel> LoadedRealizations { get; } = new();
    private bool _showLoadedRealizations;
    public bool ShowLoadedRealizations
    {
        get => _showLoadedRealizations;
        set => Set(ref _showLoadedRealizations, value);
    }
    // Все выбранные строки из запроса 1С.
    public int SelectedOneCCount        => LoadedRealizations.Count(r => r.IsSelected);
    // Только без чека — иногда вручную пробиваются как обычный "Приход" через API.
    public int SelectedNoCheckCount     => LoadedRealizations.Count(r => r.IsSelected && !r.HasCheck);
    // С пробитым не в день реализации чеком — исправительный комплект XML.
    public int SelectedHasCheckCount    => LoadedRealizations.Count(r => r.IsSelected && r.HasCheck);
    public int TotalOneCLoaded       => LoadedRealizations.Count;
    public bool CanPunchSelectedRealizationsViaAtol =>
        SelectedNoCheckCount > 0 && CheckType == "sell";

    // ── Results ──────────────────────────────────────────────────────────────
    public ObservableCollection<GenerationResult>    Results         { get; } = new();
    public ObservableCollection<ResultDisplayEntry>  AllResultEntries { get; } = new();

    private bool _showResults;
    public bool ShowResults { get => _showResults; set => Set(ref _showResults, value); }

    // ── Выбранный элемент + предпросмотр ─────────────────────────────────────
    private ResultDisplayEntry? _selectedEntry;
    public ResultDisplayEntry? SelectedEntry
    {
        get => _selectedEntry;
        set
        {
            Set(ref _selectedEntry, value);
            if (value is null)
            {
                SelectedReceiptPreview = string.Empty;
            }
            else if (value.IsPair)
            {
                var checks = new List<CheckData>();
                if (value.Refund?.CheckData     is not null) checks.Add(value.Refund.CheckData);
                if (value.Correction?.CheckData is not null) checks.Add(value.Correction.CheckData);
                SelectedReceiptPreview = ReceiptPreviewService.Generate(checks);
            }
            else
            {
                SelectedReceiptPreview = value.Single?.CheckData is not null
                    ? ReceiptPreviewService.Generate(new[] { value.Single.CheckData })
                    : string.Empty;
            }
        }
    }

    private string _selectedReceiptPreview = string.Empty;
    public string SelectedReceiptPreview
    {
        get => _selectedReceiptPreview;
        set => Set(ref _selectedReceiptPreview, value);
    }

    private string _statusText = "Готов к работе";
    public string StatusText { get => _statusText; set => Set(ref _statusText, value); }

    // ── Toast ────────────────────────────────────────────────────────────────
    private string _toastMessage = string.Empty;
    private bool   _toastVisible;
    private bool   _toastIsError;
    public string ToastMessage { get => _toastMessage; set => Set(ref _toastMessage, value); }
    public bool   ToastVisible { get => _toastVisible; set => Set(ref _toastVisible, value); }
    public bool   ToastIsError { get => _toastIsError; set => Set(ref _toastIsError, value); }

    // ── 1C Connection ────────────────────────────────────────────────────────
    private string _oneCServer   = string.Empty;
    private string _oneCDatabase = string.Empty;
    private string _oneCUser     = string.Empty;
    private string _oneCPassword = string.Empty;
    private string _oneCFromDate = DateTime.Today.AddDays(-7).ToString("dd.MM.yyyy");
    private string _oneCToDate   = DateTime.Today.ToString("dd.MM.yyyy");
    private string _oneCStatus   = string.Empty;
    private bool   _showOneCPanel;

    public string OneCServer   { get => _oneCServer;   set => Set(ref _oneCServer,   value); }
    public string OneCDatabase { get => _oneCDatabase; set => Set(ref _oneCDatabase, value); }
    public string OneCUser     { get => _oneCUser;     set => Set(ref _oneCUser,     value); }
    public string OneCPassword { get => _oneCPassword; set => Set(ref _oneCPassword, value); }
    public string OneCFromDate { get => _oneCFromDate; set => Set(ref _oneCFromDate, value); }
    public string OneCToDate   { get => _oneCToDate;   set => Set(ref _oneCToDate,   value); }
    public string OneCStatus   { get => _oneCStatus;   set => Set(ref _oneCStatus,   value); }

    public bool ShowOneCPanel
    {
        get => _showOneCPanel;
        set { Set(ref _showOneCPanel, value); OnPropertyChanged(nameof(OneCPanelVisible)); }
    }
    public bool OneCPanelVisible => ShowOneCPanel && IsRealizationTab;

    public bool IsOneCAvailable => OneCService.IsAvailable();

    // ── АТОЛ Online ──────────────────────────────────────────────────────────
    private string _atolLogin     = string.Empty;
    private string _atolPassword  = string.Empty;
    private string _atolGroupCode = string.Empty;
    private string _atolStatus    = string.Empty;
    private bool   _showAtolPanel;

    public string AtolLogin     { get => _atolLogin;     set => Set(ref _atolLogin,     value); }
    public string AtolPassword  { get => _atolPassword;  set => Set(ref _atolPassword,  value); }
    public string AtolGroupCode { get => _atolGroupCode; set => Set(ref _atolGroupCode, value); }
    public string AtolStatus    { get => _atolStatus;    set => Set(ref _atolStatus,    value); }

    private string _atolLastError = string.Empty;
    public string AtolLastError
    {
        get => _atolLastError;
        set { Set(ref _atolLastError, value); OnPropertyChanged(nameof(HasAtolError)); }
    }
    public bool   HasAtolError => !string.IsNullOrEmpty(_atolLastError);
    public string AtolLogPath  => AtolApiService.LogPath;

    public bool   ShowAtolPanel
    {
        get => _showAtolPanel;
        set => Set(ref _showAtolPanel, value);
    }

    private bool _showOfdToolsPanel;
    public bool ShowOfdToolsPanel
    {
        get => _showOfdToolsPanel;
        set => Set(ref _showOfdToolsPanel, value);
    }

    // ── Commands ─────────────────────────────────────────────────────────────
    public ICommand ParseBulkCommand       { get; }
    public ICommand AddSingleCommand       { get; }
    public ICommand DeleteOrderCommand     { get; }
    public ICommand ClearOrdersCommand     { get; }
    public ICommand GenerateCommand        { get; }
    public ICommand OpenFolderCommand      { get; }
    public ICommand AddItemCommand         { get; }
    public ICommand DeleteItemCommand      { get; }
    public ICommand SwitchTabCommand       { get; }
    public ICommand OpenWorkspaceCommand   { get; }
    public ICommand OpenSettingsCommand    { get; }
    public ICommand ImportExcelCommand     { get; }
    public ICommand ToggleOneCPanelCommand     { get; }
    public ICommand TestOneCCommand            { get; }
    public ICommand LoadFromOneCCommand        { get; }
    public ICommand SelectAllOneCCommand       { get; }
    public ICommand DeselectAllOneCCommand     { get; }
    public ICommand AddSelectedToOrdersCommand { get; }
    public ICommand ToggleAtolPanelCommand     { get; }
    public ICommand SaveAtolSettingsCommand    { get; }
    public ICommand TestAtolConnectionCommand  { get; }
    public ICommand PunchViaAtolCommand          { get; }
    public ICommand PunchOrdersViaAtolCommand  { get; }
    public ICommand MatchOfdReportCommand      { get; }
    public ICommand ApplyToOneCCommand         { get; }
    public ICommand ApplyOfdReportToOneCCommand { get; }
    public ICommand ApplyXmlAndOfdToOneCCommand { get; }
    public ICommand EditOrderCommand            { get; }

    // ── Skipped rows from Excel import ───────────────────────────────────────
    public ObservableCollection<SkippedRow> SkippedRows { get; } = new();
    private bool _showSkipped;
    public bool ShowSkipped { get => _showSkipped; set => Set(ref _showSkipped, value); }

    public MainViewModel()
    {
        ObsidianCases = new ObsidianCasesViewModel();
        CorrectionWork = new CorrectionWorkViewModel();
        ObsidianCases.OfdReportProvider = () =>
            (Reports.OfdReportPath, Reports.OfdChecks.ToList());
        ObsidianCases.OnlineOfdRowsImported = Reports.AddOnlineOfdReceipts;
        ObsidianCases.SendToWorkRequested += AddObsidianCasesToWork;
        CorrectionWork.BackRequested += CloseCorrectionWorkWindow;
        CorrectionWork.EditRequested += EditCorrectionWorkItem;
        CorrectionWork.Generated += results =>
        {
            ObsidianCases.RecordGenerated(results);
            StatusText = $"Сформировано исправительных чеков: {results.Count}";
            ShowToast($"XML готов: {results.Count} чеков", false);
        };

        // Любое изменение списка заказов автоматически обновляет
        // флаги доступности кнопки «Пробить в АТОЛ» и подсказку.
        Orders.CollectionChanged += (_, __) =>
        {
            OnPropertyChanged(nameof(OrderCount));
            OnPropertyChanged(nameof(HasXmlOnlyCorrection));
            OnPropertyChanged(nameof(XmlOnlyCorrectionCount));
            OnPropertyChanged(nameof(CanPunchOrdersViaAtol));
            OnPropertyChanged(nameof(CorrectionPunchHint));
            OnPropertyChanged(nameof(ShowCorrectionPunchHint));
        };

        ParseBulkCommand   = new AsyncRelayCommand(ParseBulkAsync);
        AddSingleCommand   = new RelayCommand(AddSingleOrder);
        DeleteOrderCommand = new RelayCommand(o => DeleteOrder(o as OrderEntry));
        ClearOrdersCommand = new RelayCommand(_ => ClearOrders());
        GenerateCommand    = new AsyncRelayCommand(GenerateChecksAsync);
        OpenFolderCommand  = new RelayCommand(_ => FileHelper.OpenFolder(FileHelper.OutputDir));
        AddItemCommand     = new RelayCommand(_ => CurrentItems.Add(new OrderItemViewModel()));
        DeleteItemCommand  = new RelayCommand(o => { if (o is OrderItemViewModel vm) CurrentItems.Remove(vm); });
        SwitchTabCommand       = new RelayCommand(t => { if (t is string s) Tab = s; });
        OpenWorkspaceCommand   = new RelayCommand(t => OpenWorkspace(t?.ToString()));
        OpenSettingsCommand    = new RelayCommand(OpenSettings);
        ImportExcelCommand     = new RelayCommand(_ => ImportExcel());
        ToggleOneCPanelCommand     = new RelayCommand(_ => ShowOneCPanel = !ShowOneCPanel);
        TestOneCCommand            = new AsyncRelayCommand(TestOneCAsync);
        LoadFromOneCCommand        = new AsyncRelayCommand(LoadFromOneCAsync);
        SelectAllOneCCommand       = new RelayCommand(_ => SetAllOneCSelected(true));
        DeselectAllOneCCommand     = new RelayCommand(_ => SetAllOneCSelected(false));
        AddSelectedToOrdersCommand = new AsyncRelayCommand(AddSelectedToOrdersAsync);
        ToggleAtolPanelCommand     = new RelayCommand(_ => ShowAtolPanel = !ShowAtolPanel);
        SaveAtolSettingsCommand    = new RelayCommand(_ => SaveAtolSettings());
        TestAtolConnectionCommand  = new AsyncRelayCommand(TestAtolConnectionAsync);
        PunchViaAtolCommand        = new AsyncRelayCommand(PunchViaAtolAsync);
        PunchOrdersViaAtolCommand  = new AsyncRelayCommand(PunchOrdersViaAtolAsync);
        MatchOfdReportCommand      = new RelayCommand(_ => MatchOfdReport());
        ApplyToOneCCommand         = new AsyncRelayCommand(ApplyToOneCAsync);
        ApplyOfdReportToOneCCommand = new AsyncRelayCommand(ApplyOfdReportToOneCAsync);
        ApplyXmlAndOfdToOneCCommand = new AsyncRelayCommand(ApplyXmlAndOfdToOneCAsync);
        EditOrderCommand            = new RelayCommand(o => EditOrder(o as OrderEntry));

        // Загружаем сохранённые настройки АТОЛ
        var saved = AtolCredentials.Load();
        AtolLogin     = saved.Login;
        AtolPassword  = saved.Password;
        AtolGroupCode = saved.GroupCode;

        // Загружаем сохранённые настройки 1С
        var savedOneC = OneCConnectionSettings.Load();
        if (!string.IsNullOrEmpty(savedOneC.Server))
        {
            OneCServer   = savedOneC.Server;
            OneCDatabase = savedOneC.Database;
            OneCUser     = savedOneC.User;
            OneCPassword = savedOneC.Password;
        }

        ReloadApplicationSettings();
    }

    public async Task InitializeAsync()
    {
        if (_startupInitialized) return;
        _startupInitialized = true;

        var reportsTask = Reports.RestoreLatestReportsAsync();
        var updateTask = ApplicationUpdateService.CheckForUpdateAsync();

        try
        {
            await reportsTask;
        }
        catch
        {
            // Автозагрузка отчётов не должна блокировать остальные рабочие сценарии.
        }

        ApplicationUpdateResult update;
        try
        {
            update = await updateTask;
        }
        catch
        {
            return;
        }

        if (!update.CheckSucceeded || !update.IsUpdateAvailable) return;

        var answer = MessageBox.Show(
            Application.Current?.MainWindow,
            $"Установлена версия {update.CurrentVersion}. Доступна версия {update.LatestVersion}.\n\n" +
            "Скачать новую версию сейчас?",
            "Доступно обновление",
            MessageBoxButton.YesNo,
            MessageBoxImage.Information);
        if (answer != MessageBoxResult.Yes) return;

        try
        {
            ApplicationUpdateService.OpenUrl(string.IsNullOrWhiteSpace(update.DownloadUrl)
                ? update.ReleaseUrl
                : update.DownloadUrl);
        }
        catch (Exception ex)
        {
            MessageBox.Show(
                Application.Current?.MainWindow,
                $"Не удалось открыть страницу обновления: {ex.Message}",
                "Обновление",
                MessageBoxButton.OK,
                MessageBoxImage.Error);
        }
    }

    // ── Logic ─────────────────────────────────────────────────────────────────

    private void OpenWorkspace(string? target)
    {
        if (target is "reports" or "ofd")
        {
            SetWorkspace("reports");
            ShowOfdToolsPanel = false;
            StatusText = "Работа с отчётами";
            return;
        }

        if (target == "obsidian")
        {
            SetWorkspace("obsidian");
            ShowOfdToolsPanel = false;
            ObsidianCases.Activate();
            StatusText = "Исправление чеков";
            return;
        }

        if (target == "correctionWork")
        {
            OpenCorrectionWorkWindow();
            return;
        }

        if (target is not ("atol" or "settings"))
            SetWorkspace("main");

        switch (target)
        {
            case "refunds":
                Tab = "payment";
                CheckType = "sell_refund";
                ShowOneCPanel = false;
                ShowOfdToolsPanel = false;
                StatusText = "Возвраты по заказам";
                break;

            case "realizationCorrections":
                Tab = "realization";
                CheckType = "sell_correction";
                ShowOneCPanel = true;
                ShowOfdToolsPanel = false;
                StatusText = "Коррекции реализаций 1С";
                break;

            case "atol":
            case "settings":
                OpenSettings();
                break;

            case "payment":
                Tab = "payment";
                CheckType = "sell";
                ShowOneCPanel = false;
                ShowOfdToolsPanel = false;
                StatusText = "Оплаты";
                break;

            case "realization":
                Tab = "realization";
                CheckType = "sell";
                ShowOneCPanel = false;
                ShowOfdToolsPanel = false;
                StatusText = "Реализации";
                break;
        }
    }

    private void SetWorkspace(string workspace)
    {
        if (string.Equals(_activeWorkspace, workspace, StringComparison.Ordinal)) return;
        _activeWorkspace = workspace;
        OnPropertyChanged(nameof(ShowMainWorkspace));
        OnPropertyChanged(nameof(ShowReportsWorkspace));
        OnPropertyChanged(nameof(ShowObsidianWorkspace));
    }

    private void OpenCorrectionWorkWindow()
    {
        CorrectionWork.SyncCashiers(AvailableCashiers, SelectedCashier);
        if (_correctionWorkWindow is not null)
        {
            if (_correctionWorkWindow.WindowState == WindowState.Minimized)
                _correctionWorkWindow.WindowState = WindowState.Normal;
            _correctionWorkWindow.Activate();
            return;
        }

        var owner = Application.Current?.MainWindow;
        var window = new Views.CorrectionWorkWindow { DataContext = CorrectionWork };
        if (owner is not null && !ReferenceEquals(owner, window))
            window.Owner = owner;
        window.Closed += (_, _) => _correctionWorkWindow = null;
        _correctionWorkWindow = window;
        window.Show();
        StatusText = "Открыто отдельное окно пробития исправлений";
    }

    private void CloseCorrectionWorkWindow()
    {
        var window = _correctionWorkWindow;
        _correctionWorkWindow = null;
        window?.Close();
        SetWorkspace("obsidian");
        ObsidianCases.Activate();
        StatusText = "Исправление чеков";
        Application.Current?.MainWindow?.Activate();
    }

    private void OpenSettings()
    {
        var dialog = new Views.SettingsWindow
        {
            Owner = System.Windows.Application.Current?.MainWindow,
        };
        if (dialog.ShowDialog() != true) return;

        ReloadApplicationSettings();
        ObsidianCases.ReloadSettings();

        var atol = AtolCredentials.Load();
        AtolLogin = atol.Login;
        AtolPassword = atol.Password;
        AtolGroupCode = atol.GroupCode;

        var oneC = OneCConnectionSettings.Load();
        OneCServer = oneC.Server;
        OneCDatabase = oneC.Database;
        OneCUser = oneC.User;
        OneCPassword = oneC.Password;

        StatusText = "Настройки сохранены";
        ShowToast("Настройки сохранены", false);
    }

    private void ReloadApplicationSettings()
    {
        var settings = ApplicationSettingsStore.Current;
        AvailableCashiers.Clear();
        foreach (var cashier in settings.Cashiers)
            AvailableCashiers.Add(cashier);

        SelectedCashier = AvailableCashiers.FirstOrDefault(c => string.Equals(
            c.ShortName, settings.SelectedCashierShortName, StringComparison.OrdinalIgnoreCase))
            ?? AvailableCashiers.FirstOrDefault()
            ?? AppConstants.DefaultCashier;
        CorrectionWork.SyncCashiers(AvailableCashiers, SelectedCashier);

        if (!string.IsNullOrWhiteSpace(SelectedServiceType))
            RefreshCities();
        if (!string.IsNullOrWhiteSpace(SelectedCity))
            PickAgent();
    }

    private void RefreshCities()
    {
        AvailableCities.Clear();
        foreach (var c in AppConstants.ServiceProviders
                     .Where(p => p.Service == SelectedServiceType)
                     .Select(p => p.City)
                     .Distinct())
            AvailableCities.Add(c);
    }

    private void PickAgent()
    {
        SelectedAgent = AppConstants.ServiceProviders
            .FirstOrDefault(p => p.Service == SelectedServiceType && p.City == SelectedCity);
    }

    private async Task ParseBulkAsync()
    {
        if (string.IsNullOrWhiteSpace(BulkText))
        { ShowToast("Введите текст с заказами", true); return; }

        var parsed = OrderParserService.Parse(BulkText);
        if (parsed.Count == 0)
        { ShowToast("Заказы не распознаны. Проверьте формат.", true); return; }

        // Обогащаем из 1С ТОЛЬКО для услуг (IsService=true но агент не определён).
        // Обычные чеки (sell) обогащение не требуют — не запрашиваем 1С без нужды.
        var hasOneC = !string.IsNullOrWhiteSpace(OneCServer) && !string.IsNullOrWhiteSpace(OneCDatabase);
        var serviceOrdersWithoutAgent = parsed.Where(o => o.IsService && o.AgentInfo is null).ToList();
        if (hasOneC && serviceOrdersWithoutAgent.Count > 0)
        {
            StatusText = $"Запрашиваем подразделение из 1С для {serviceOrdersWithoutAgent.Count} услуг…";
            try
            {
                var oneCSettings = BuildOneCSettings();
                await Task.Run(() => OneCService.EnrichOrdersFromOneC(oneCSettings, serviceOrdersWithoutAgent));
            }
            catch { /* если 1С недоступна — продолжаем без обогащения */ }
            StatusText = "Готов к работе";
        }

        var existing = new HashSet<string>(Orders.Select(o => o.OrderNum));
        int added = 0;
        foreach (var o in parsed)
        {
            if (!existing.Contains(o.OrderNum))
            { ApplyCorrectionFields(o); Orders.Add(o); added++; }
        }
        OnPropertyChanged(nameof(OrderCount));
        OnPropertyChanged(nameof(CanPunchOrdersViaAtol));

        ShowToast($"Добавлено {added} из {parsed.Count} заказ(ов)", added > 0);
    }

    private void AddSingleOrder()
    {
        if (string.IsNullOrWhiteSpace(SingleOrderNum))
        { ShowToast("Введите номер заказа", true); return; }
        if (!double.TryParse(SingleAmount.Replace(',', '.'),
            System.Globalization.NumberStyles.Any,
            System.Globalization.CultureInfo.InvariantCulture, out var amt) || amt <= 0)
        { ShowToast("Введите корректную сумму", true); return; }

        var order = new OrderEntry
        {
            OrderNum     = SingleOrderNum.Trim(),
            OrderDate    = SingleOrderDate.Trim(),
            Amount       = amt,
            CustomerName = SingleCustomer.Trim(),
            AgentInfo    = IsServiceProvider ? SelectedAgent : null,
            ServiceType  = IsServiceProvider ? SelectedServiceType : string.Empty,
            Items        = CurrentItems.Select(i => i.ToModel()).ToList(),
        };
        ApplyCorrectionFields(order);
        Orders.Add(order);
        OnPropertyChanged(nameof(OrderCount));
        OnPropertyChanged(nameof(CanPunchOrdersViaAtol));
        ShowToast($"Добавлен {order.OrderNum}", false);
    }

    private void ApplyCorrectionFields(OrderEntry o)
    {
        if (IsCorrection)
        {
            o.CorrectionDate   = CorrectionDate;
            o.CorrectionNumber = CorrectionNumber;
        }
    }

    private void DeleteOrder(OrderEntry? order)
    {
        if (order is not null)
        {
            Orders.Remove(order);
            OnPropertyChanged(nameof(OrderCount));
        OnPropertyChanged(nameof(CanPunchOrdersViaAtol));
        }
    }

    private void ClearOrders()
    {
        Orders.Clear();
        OnPropertyChanged(nameof(OrderCount));
        OnPropertyChanged(nameof(CanPunchOrdersViaAtol));
    }

    private void ImportExcel()
    {
        var dlg = new OpenFileDialog
        {
            Title  = "Выберите файл Excel",
            Filter = "Excel файлы (*.xlsx)|*.xlsx|Все файлы (*.*)|*.*",
        };
        if (dlg.ShowDialog() != true) return;

        try
        {
            var result = ExcelImportService.Import(dlg.FileName);

            int added = 0;
            var existing = new HashSet<string>(Orders.Select(o => o.OrderNum));
            foreach (var order in result.Orders)
            {
                if (!existing.Contains(order.OrderNum))
                { Orders.Add(order); added++; }
            }
            OnPropertyChanged(nameof(OrderCount));
        OnPropertyChanged(nameof(CanPunchOrdersViaAtol));

            SkippedRows.Clear();
            foreach (var s in result.SkippedRows) SkippedRows.Add(s);
            ShowSkipped = SkippedRows.Count > 0;

            var msg = $"Импортировано {added} из {result.Orders.Count} заказ(ов)";
            if (result.SkippedRows.Count > 0)
                msg += $". Пропущено {result.SkippedRows.Count} — уже пробиты чеки (требуют ручного исправления)";
            ShowToast(msg, result.SkippedRows.Count > 0 && added == 0);
        }
        catch (Exception ex)
        {
            ShowToast($"Ошибка импорта: {ex.Message}", true);
        }
    }

    private void MatchOfdReport()
    {
        var dlg = new Microsoft.Win32.OpenFileDialog
        {
            Title  = "Выберите сводный отчёт ОФД (Excel)",
            Filter = "Excel-файлы|*.xlsx;*.xls",
        };
        if (dlg.ShowDialog() != true) return;

        try
        {
            var ts  = DateTime.Now.ToString("yyyyMMdd_HHmmss");
            var dir = System.IO.Path.GetDirectoryName(dlg.FileName) ?? FileHelper.OutputDir;
            var output = System.IO.Path.Combine(dir, $"ОФД_сопоставление_{ts}.xlsx");

            var result = OfdReportMatcherService.MatchAndExport(dlg.FileName, output);
            var msg = $"Обработано чеков: {result.TotalRows}\n" +
                      $"Сопоставлено с пробитиями: {result.MatchedRows}\n" +
                      $"Не найдено в логе: {result.UnmatchedRows}\n\n" +
                      $"Файл сохранён:\n{result.OutputPath}";

            System.Windows.MessageBox.Show(msg, "Сопоставление ОФД",
                System.Windows.MessageBoxButton.OK,
                System.Windows.MessageBoxImage.Information);

            FileHelper.OpenFolder(dir);
        }
        catch (Exception ex)
        {
            System.Windows.MessageBox.Show($"Ошибка сопоставления: {ex.Message}",
                "Ошибка", System.Windows.MessageBoxButton.OK,
                System.Windows.MessageBoxImage.Error);
        }
    }

    private async Task ApplyToOneCAsync()
    {
        if (string.IsNullOrWhiteSpace(OneCServer) ||
            string.IsNullOrWhiteSpace(OneCDatabase))
        {
            System.Windows.MessageBox.Show(
                "Заполните настройки подключения к 1С (Сервер + База).",
                "Применение к 1С", System.Windows.MessageBoxButton.OK,
                System.Windows.MessageBoxImage.Warning);
            return;
        }

        var path = AtolApiService.PunchedJsonPath;
        if (!System.IO.File.Exists(path))
        {
            System.Windows.MessageBox.Show(
                $"Файл журнала пробитий не найден:\n{path}\n\n" +
                "Сначала пробейте чеки через АТОЛ Online — журнал заполнится автоматически.",
                "Применение к 1С", System.Windows.MessageBoxButton.OK,
                System.Windows.MessageBoxImage.Information);
            return;
        }

        // Читаем jsonl и формируем список записей
        var records = new List<OneCService.PunchedRecord>();
        try
        {
            foreach (var line in System.IO.File.ReadLines(path))
            {
                if (string.IsNullOrWhiteSpace(line)) continue;
                try
                {
                    using var doc = System.Text.Json.JsonDocument.Parse(line);
                    var root = doc.RootElement;
                    var rec = new OneCService.PunchedRecord
                    {
                        RealizationNum = root.TryGetProperty("realization_num", out var rn) ? rn.GetString() ?? "" : "",
                        FiscalDoc      = root.TryGetProperty("fiscal_doc",  out var fd) && fd.ValueKind == System.Text.Json.JsonValueKind.Number ? fd.GetInt64() : null,
                        FiscalSign     = root.TryGetProperty("fiscal_sign", out var fs) && fs.ValueKind == System.Text.Json.JsonValueKind.Number ? fs.GetInt64() : null,
                        ReceiptDt      = root.TryGetProperty("receipt_dt",  out var dt) ? dt.GetString() ?? "" : "",
                    };
                    if (!string.IsNullOrEmpty(rec.RealizationNum) && rec.FiscalDoc.HasValue && rec.FiscalSign.HasValue)
                        records.Add(rec);
                }
                catch { /* битая строка — пропускаем */ }
            }
        }
        catch (Exception ex)
        {
            System.Windows.MessageBox.Show($"Не удалось прочитать журнал: {ex.Message}",
                "Применение к 1С", System.Windows.MessageBoxButton.OK,
                System.Windows.MessageBoxImage.Error);
            return;
        }

        if (records.Count == 0)
        {
            System.Windows.MessageBox.Show(
                "В журнале нет пригодных записей (нужны № реализации, ФПД и № ФД).",
                "Применение к 1С", System.Windows.MessageBoxButton.OK,
                System.Windows.MessageBoxImage.Information);
            return;
        }

        // Подтверждение: Да = пропускать заполненные, Нет = перезаписать всё, Отмена = выход
        var confirm = System.Windows.MessageBox.Show(
            $"Найдено {records.Count} пробитых чеков в журнале.\n\n" +
            $"Программа подключится к 1С и для каждой реализации запишет:\n" +
            $"  ЧекНомерФП    — ФПД\n" +
            $"  НомерЧекаККМ  — № ФД\n" +
            $"  ДатаПечатиЧека — дата чека\n\n" +
            $"Да — записать ТОЛЬКО в пустые документы\n" +
            $"Нет — ПЕРЕЗАПИСАТЬ ВСЕ (включая уже заполненные)\n" +
            $"Отмена — выход",
            "Подтверждение", System.Windows.MessageBoxButton.YesNoCancel,
            System.Windows.MessageBoxImage.Question);
        if (confirm == System.Windows.MessageBoxResult.Cancel) return;
        bool skipFilled = (confirm == System.Windows.MessageBoxResult.Yes);

        AtolStatus = $"Применяем к 1С: 0/{records.Count}...";
        StatusText = AtolStatus;

        var settings = new OneCConnectionSettings
        {
            Server   = OneCServer,
            Database = OneCDatabase,
            User     = OneCUser,
            Password = OneCPassword,
        };

        var result = await Task.Run(() => OneCService.ApplyPunchedChecks(settings, records, skipFilled));

        var msg = $"Обновлено: {result.Updated}\n" +
                  $"Пропущено: {result.Skipped}\n" +
                  $"Ошибок:    {result.Failed}";
        if (result.SkippedSamples.Count > 0)
            msg += "\n\nПервые пропуски:\n  " + string.Join("\n  ", result.SkippedSamples.Take(10));
        if (result.Errors.Count > 0)
            msg += "\n\nПервые ошибки:\n" + string.Join("\n", result.Errors.Take(10));
        if (!string.IsNullOrEmpty(result.CsvBackupPath))
            msg += $"\n\n📄 CSV для ручного импорта:\n{result.CsvBackupPath}";

        AtolStatus = $"Применено к 1С: ✓ {result.Updated}, ✗ {result.Failed}";
        StatusText = AtolStatus;

        System.Windows.MessageBox.Show(msg, "Применение к 1С завершено",
            System.Windows.MessageBoxButton.OK,
            result.Failed > 0 ? System.Windows.MessageBoxImage.Warning : System.Windows.MessageBoxImage.Information);
    }

    private async Task ApplyOfdReportToOneCAsync()
    {
        if (string.IsNullOrWhiteSpace(OneCServer) ||
            string.IsNullOrWhiteSpace(OneCDatabase))
        {
            System.Windows.MessageBox.Show(
                "Заполните настройки подключения к 1С (Сервер + База).",
                "Применение отчёта ОФД", System.Windows.MessageBoxButton.OK,
                System.Windows.MessageBoxImage.Warning);
            return;
        }

        var dlg = new Microsoft.Win32.OpenFileDialog
        {
            Title  = "Выберите сводный отчёт ОФД (Excel)",
            Filter = "Excel-файлы|*.xlsx;*.xls",
        };
        if (dlg.ShowDialog() != true) return;

        List<OneCService.PunchedRecord> records;
        try
        {
            records = OneCService.ReadOfdReport(dlg.FileName);
        }
        catch (Exception ex)
        {
            System.Windows.MessageBox.Show($"Не удалось прочитать отчёт: {ex.Message}",
                "Применение отчёта ОФД", System.Windows.MessageBoxButton.OK,
                System.Windows.MessageBoxImage.Error);
            return;
        }

        if (records.Count == 0)
        {
            System.Windows.MessageBox.Show(
                "В отчёте не найдены строки с заполненным «Значением доп.реквизита пользователя».\n\n" +
                "Чеки, пробитые до v1.7.4, не содержат тега 1086 (номер реализации) — их нужно " +
                "переносить вручную или через локальный журнал punched_checks.jsonl.",
                "Применение отчёта ОФД", System.Windows.MessageBoxButton.OK,
                System.Windows.MessageBoxImage.Information);
            return;
        }

        var confirm = System.Windows.MessageBox.Show(
            $"В отчёте найдено {records.Count} чеков с привязкой к реализации.\n\n" +
            $"Программа подключится к 1С и для каждой реализации запишет:\n" +
            $"  ЧекНомерФП    — ФПД\n" +
            $"  НомерЧекаККМ  — № ФД\n" +
            $"  ДатаПечатиЧека — дата чека из ОФД\n\n" +
            $"Да — записать ТОЛЬКО в пустые документы\n" +
            $"Нет — ПЕРЕЗАПИСАТЬ ВСЕ (включая уже заполненные)\n" +
            $"Отмена — выход",
            "Подтверждение", System.Windows.MessageBoxButton.YesNoCancel,
            System.Windows.MessageBoxImage.Question);
        if (confirm == System.Windows.MessageBoxResult.Cancel) return;
        bool skipFilled = (confirm == System.Windows.MessageBoxResult.Yes);

        AtolStatus = $"Применяем отчёт к 1С: 0/{records.Count}...";
        StatusText = AtolStatus;

        var settings = new OneCConnectionSettings
        {
            Server   = OneCServer,
            Database = OneCDatabase,
            User     = OneCUser,
            Password = OneCPassword,
        };

        var result = await Task.Run(() => OneCService.ApplyPunchedChecks(settings, records, skipFilled));

        var msg = $"Обновлено: {result.Updated}\n" +
                  $"Пропущено: {result.Skipped}\n" +
                  $"Ошибок:    {result.Failed}";
        if (result.SkippedSamples.Count > 0)
            msg += "\n\nПервые пропуски:\n  " + string.Join("\n  ", result.SkippedSamples.Take(10));
        if (result.Errors.Count > 0)
            msg += "\n\nПервые ошибки:\n" + string.Join("\n", result.Errors.Take(10));
        if (!string.IsNullOrEmpty(result.CsvBackupPath))
            msg += $"\n\n📄 CSV для ручного импорта:\n{result.CsvBackupPath}";

        AtolStatus = $"Отчёт ОФД применён: ✓ {result.Updated}, ✗ {result.Failed}";
        StatusText = AtolStatus;

        System.Windows.MessageBox.Show(msg, "Применение отчёта ОФД завершено",
            System.Windows.MessageBoxButton.OK,
            result.Failed > 0 ? System.Windows.MessageBoxImage.Warning : System.Windows.MessageBoxImage.Information);
    }

    private async Task ApplyXmlAndOfdToOneCAsync()
    {
        if (string.IsNullOrWhiteSpace(OneCServer) ||
            string.IsNullOrWhiteSpace(OneCDatabase))
        {
            System.Windows.MessageBox.Show(
                "Заполните настройки подключения к 1С (Сервер + База).",
                "Применение XML+ОФД", System.Windows.MessageBoxButton.OK,
                System.Windows.MessageBoxImage.Warning);
            return;
        }

        // 1. Выбираем XML
        var xmlDlg = new Microsoft.Win32.OpenFileDialog
        {
            Title  = "Шаг 1/2: выберите XML-файл с пробитыми чеками коррекции",
            Filter = "XML-файлы|*.xml",
        };
        if (xmlDlg.ShowDialog() != true) return;

        // 2. Выбираем отчёт ОФД
        var ofdDlg = new Microsoft.Win32.OpenFileDialog
        {
            Title  = "Шаг 2/2: выберите сводный отчёт ОФД (Excel)",
            Filter = "Excel-файлы|*.xlsx;*.xls",
        };
        if (ofdDlg.ShowDialog() != true) return;

        List<XmlOfdMatcherService.XmlCheck> xml;
        List<XmlOfdMatcherService.OfdRow>   ofd;
        try
        {
            xml = XmlOfdMatcherService.ReadXmlChecks(xmlDlg.FileName);
            ofd = XmlOfdMatcherService.ReadOfdCorrections(ofdDlg.FileName);
        }
        catch (Exception ex)
        {
            System.Windows.MessageBox.Show($"Ошибка чтения: {ex.Message}",
                "Применение XML+ОФД", System.Windows.MessageBoxButton.OK,
                System.Windows.MessageBoxImage.Error);
            return;
        }

        if (xml.Count == 0)
        {
            System.Windows.MessageBox.Show(
                "В XML не найдено чеков коррекции с номером реализации.",
                "Применение XML+ОФД", System.Windows.MessageBoxButton.OK,
                System.Windows.MessageBoxImage.Information);
            return;
        }
        if (ofd.Count == 0)
        {
            System.Windows.MessageBox.Show(
                "В отчёте ОФД не найдено чеков коррекции.",
                "Применение XML+ОФД", System.Windows.MessageBoxButton.OK,
                System.Windows.MessageBoxImage.Information);
            return;
        }

        var match = XmlOfdMatcherService.Match(xml, ofd);
        var unmatchedTxt = match.Unmatched > 0
            ? $"\n\nНе сопоставлено: {match.Unmatched}\n  " +
              string.Join("\n  ", match.Warnings.Take(10))
            : string.Empty;

        // Первый диалог — да/нет/отмена. По умолчанию (Да) пропускаем уже заполненные,
        // Нет → перезаписать ВСЕ (даже заполненные), Отмена → выход.
        var confirm = System.Windows.MessageBox.Show(
            $"XML: {xml.Count} чеков\n" +
            $"ОФД: {ofd.Count} коррекций\n" +
            $"Сопоставлено по сумме: {match.Matched}{unmatchedTxt}\n\n" +
            $"Да — записать ТОЛЬКО в пустые документы (уже заполненные пропустить)\n" +
            $"Нет — ПЕРЕЗАПИСАТЬ ВСЕ (включая уже заполненные)\n" +
            $"Отмена — выход",
            "Подтверждение", System.Windows.MessageBoxButton.YesNoCancel,
            System.Windows.MessageBoxImage.Question);
        if (confirm == System.Windows.MessageBoxResult.Cancel) return;
        bool skipFilled = (confirm == System.Windows.MessageBoxResult.Yes);

        AtolStatus = $"Записываем в 1С: 0/{match.Records.Count}...";
        StatusText = AtolStatus;

        var settings = new OneCConnectionSettings
        {
            Server   = OneCServer,
            Database = OneCDatabase,
            User     = OneCUser,
            Password = OneCPassword,
        };

        var result = await Task.Run(() =>
            OneCService.ApplyPunchedChecks(settings, match.Records, skipFilled: skipFilled));

        var modeStr = skipFilled ? "(пропуская заполненные)" : "(перезаписать всё)";
        var msg = $"Сопоставлено XML↔ОФД: {match.Matched}\n" +
                  $"Не сопоставлено:    {match.Unmatched}\n" +
                  $"────────────────────\n" +
                  $"Записано в 1С:      {result.Updated}\n" +
                  $"Пропущено (заполн.):{result.Skipped}\n" +
                  $"Ошибок:             {result.Failed}\n" +
                  $"Режим:              {modeStr}";

        if (result.SkippedSamples.Count > 0 && result.Skipped > 0)
        {
            msg += "\n\nПервые пропуски (дата+значение):\n  " +
                   string.Join("\n  ", result.SkippedSamples.Take(10));
        }

        if (result.Errors.Count > 0)
            msg += "\n\nПервые ошибки:\n" + string.Join("\n", result.Errors.Take(10));

        if (!string.IsNullOrEmpty(result.CsvBackupPath))
            msg += $"\n\n📄 CSV для ручного импорта в 1С (через внешнюю обработку):\n{result.CsvBackupPath}";

        AtolStatus = $"XML+ОФД → 1С: ✓ {result.Updated}, пропущено {result.Skipped}";
        StatusText = AtolStatus;

        System.Windows.MessageBox.Show(msg, "Применение XML+ОФД завершено",
            System.Windows.MessageBoxButton.OK,
            (result.Failed > 0 || match.Unmatched > 0)
                ? System.Windows.MessageBoxImage.Warning
                : System.Windows.MessageBoxImage.Information);
    }

    /// <summary>Открывает модальное окно редактирования записи.</summary>
    private void EditOrder(OrderEntry? order)
    {
        if (order is null) return;
        try
        {
            // Работаем с копией — если пользователь отменит, оригинал не пострадает.
            var copy = CloneOrderEntry(order);
            var dlg = new Views.CorrectionEditorWindow(copy)
            {
                Owner = ResolveDialogOwner(),
            };

            if (dlg.ShowDialog() != true) return;

            ApplyOrderEntry(copy, order);
            var idx = Orders.IndexOf(order);
            if (idx >= 0)
            {
                Orders.RemoveAt(idx);
                Orders.Insert(idx, order);
            }
        }
        catch (Exception ex)
        {
            ShowEditorError(ex);
        }
    }

    private void EditCorrectionWorkItem(CorrectionWorkItemViewModel item)
    {
        try
        {
            var copy = CloneOrderEntry(item.Entry);
            var dialog = new Views.CorrectionEditorWindow(copy)
            {
                Owner = ResolveDialogOwner(),
            };
            if (dialog.ShowDialog() != true) return;

            ApplyOrderEntry(copy, item.Entry);
            CorrectionWork.RefreshItem(item);
            StatusText = $"Исправление {item.DocumentNumber} обновлено";
        }
        catch (Exception ex)
        {
            ShowEditorError(ex);
        }
    }

    private Window? ResolveDialogOwner() =>
        _correctionWorkWindow is { IsVisible: true }
            ? _correctionWorkWindow
            : System.Windows.Application.Current?.MainWindow;

    private void ShowEditorError(Exception ex)
    {
        StatusText = "Не удалось открыть редактор исправления";
        ShowToast(StatusText, true);
        MessageBox.Show(
            $"Редактор не смог открыть выбранный чек.\n\n{ex.Message}",
            "Редактирование исправления",
            MessageBoxButton.OK,
            MessageBoxImage.Error);
    }

    /// <summary>Создаёт независимую копию OrderEntry (мелкое клонирование).</summary>
    private static OrderEntry CloneOrderEntry(OrderEntry s) => new()
    {
        ObsidianCaseId        = s.ObsidianCaseId,
        OrderNum             = s.OrderNum,
        OrderDate            = s.OrderDate,
        SourceDocumentDate   = s.SourceDocumentDate,
        Amount               = s.Amount,
        CustomerName         = s.CustomerName,
        Items                = s.Items.Select(i => new OrderItem
        {
            Name     = i.Name,
            Quantity = i.Quantity,
            Sum      = i.Sum,
            VatType  = i.VatType,
        }).ToList(),
        OriginalItems        = s.OriginalItems.Select(i => new OrderItem
        {
            Name     = i.Name,
            Quantity = i.Quantity,
            Sum      = i.Sum,
            VatType  = i.VatType,
        }).ToList(),
        AgentInfo            = s.AgentInfo,
        CorrectionDate       = s.CorrectionDate,
        CorrectionNumber     = s.CorrectionNumber,
        IsService            = s.IsService,
        IsOwnService         = s.IsOwnService,
        ServiceType          = s.ServiceType,
        City                 = s.City,
        Kind                 = s.Kind,
        DocumentType         = s.DocumentType,
        CorrectionScenario   = s.CorrectionScenario,
        OriginalFiscalNumber = s.OriginalFiscalNumber,
        OriginalCheckAmount  = s.OriginalCheckAmount,
        CorrectAmount        = s.CorrectAmount,
        Notes                = s.Notes,
        OriginalPaymentWasCash = s.OriginalPaymentWasCash,
        CorrectPaymentIsCash   = s.CorrectPaymentIsCash,
        PlannedReverseOperation = s.PlannedReverseOperation,
        PlannedCorrectOperation = s.PlannedCorrectOperation,
        OriginalVatType = s.OriginalVatType,
        CorrectVatType = s.CorrectVatType,
        PlannedVatType = s.PlannedVatType,
        OriginalCheckDate = s.OriginalCheckDate,
        OriginalCheckOperation = s.OriginalCheckOperation,
        OneCComment = s.OneCComment,
        OneCCheckNumber = s.OneCCheckNumber,
        OneCCheckDate = s.OneCCheckDate,
    };

    /// <summary>Копирует поля из источника в цель (используется после редактирования).</summary>
    private static void ApplyOrderEntry(OrderEntry from, OrderEntry to)
    {
        to.ObsidianCaseId        = from.ObsidianCaseId;
        to.OrderNum             = from.OrderNum;
        to.OrderDate            = from.OrderDate;
        to.SourceDocumentDate   = from.SourceDocumentDate;
        to.Amount               = from.Amount;
        to.CustomerName         = from.CustomerName;
        to.Items                = from.Items.Select(i => new OrderItem
        {
            Name     = i.Name,
            Quantity = i.Quantity,
            Sum      = i.Sum,
            VatType  = i.VatType,
        }).ToList();
        to.OriginalItems        = from.OriginalItems.Select(i => new OrderItem
        {
            Name     = i.Name,
            Quantity = i.Quantity,
            Sum      = i.Sum,
            VatType  = i.VatType,
        }).ToList();
        to.AgentInfo            = from.AgentInfo;
        to.CorrectionDate       = from.CorrectionDate;
        to.CorrectionNumber     = from.CorrectionNumber;
        to.IsService            = from.IsService;
        to.IsOwnService         = from.IsOwnService;
        to.ServiceType          = from.ServiceType;
        to.City                 = from.City;
        to.Kind                 = from.Kind;
        to.DocumentType         = from.DocumentType;
        to.CorrectionScenario   = from.CorrectionScenario;
        to.OriginalFiscalNumber = from.OriginalFiscalNumber;
        to.OriginalCheckAmount  = from.OriginalCheckAmount;
        to.CorrectAmount        = from.CorrectAmount;
        to.Notes                = from.Notes;
        to.OriginalPaymentWasCash = from.OriginalPaymentWasCash;
        to.CorrectPaymentIsCash   = from.CorrectPaymentIsCash;
        to.PlannedReverseOperation = from.PlannedReverseOperation;
        to.PlannedCorrectOperation = from.PlannedCorrectOperation;
        to.OriginalVatType = from.OriginalVatType;
        to.CorrectVatType = from.CorrectVatType;
        to.PlannedVatType = from.PlannedVatType;
        to.OriginalCheckDate = from.OriginalCheckDate;
        to.OriginalCheckOperation = from.OriginalCheckOperation;
        to.OneCComment = from.OneCComment;
        to.OneCCheckNumber = from.OneCCheckNumber;
        to.OneCCheckDate = from.OneCCheckDate;
    }

    private void SaveAtolSettings()
    {
        var creds = new AtolCredentials
        {
            Login     = AtolLogin.Trim(),
            Password  = AtolPassword,
            GroupCode = AtolGroupCode.Trim(),
        };
        creds.Save();
        AtolStatus = "Настройки сохранены";
        ShowToast("Настройки АТОЛ сохранены", false);
    }

    private async Task TestAtolConnectionAsync()
    {
        if (string.IsNullOrWhiteSpace(AtolLogin))
        { AtolStatus = "⚠️ Заполните логин"; return; }
        if (string.IsNullOrWhiteSpace(AtolPassword))
        { AtolStatus = "⚠️ Заполните пароль"; return; }
        if (string.IsNullOrWhiteSpace(AtolGroupCode))
        { AtolStatus = "⚠️ Заполните Group Code"; return; }

        AtolApiService.InvalidateToken();   // сбрасываем кэш

        var creds = new AtolCredentials
        {
            Login     = AtolLogin.Trim(),
            Password  = AtolPassword,
            GroupCode = AtolGroupCode.Trim(),
        };

        // Шаг 1: токен (логин + пароль)
        AtolStatus = "⏳ Шаг 1/2: проверяем логин и пароль…";
        var (token, tokenErr) = await AtolApiService.GetTokenAsync(creds);
        if (token is null)
        {
            AtolStatus = $"❌ Ошибка авторизации: {tokenErr}";
            ShowToast($"АТОЛ: {tokenErr}", true);
            return;
        }

        // Шаг 2: group_code
        AtolStatus = "⏳ Шаг 2/2: проверяем Group Code…";
        var (gcOk, gcDetail) = await AtolApiService.TestGroupCodeAsync(creds.GroupCode, token);
        if (gcOk)
        {
            AtolStatus = $"✅ Подключено успешно. Логин и Group Code корректны.";
            ShowToast("АТОЛ Online: подключение успешно", false);
        }
        else
        {
            AtolStatus = $"⚠️ Логин/пароль верны, но Group Code недоступен.\n{gcDetail}";
            ShowToast("АТОЛ: неверный Group Code", true);
        }
    }

    private async Task<List<string>> PrepareOneCRealizationsAsync(List<OneCRealizationViewModel> rows)
    {
        if (rows.Count == 0) return new List<string>();

        var settings = BuildOneCSettings();
        return await Task.Run(() =>
        {
            var errors = new List<string>();
            foreach (var row in rows)
            {
                try
                {
                    OneCService.EnrichRealizationForReceipt(settings, row.Source);
                }
                catch (Exception ex)
                {
                    errors.Add($"{row.DocNumber}: не удалось загрузить номенклатуру из 1С ({ex.Message})");
                    continue;
                }

                if (!row.Source.IsService) continue;

                if (string.IsNullOrWhiteSpace(row.Source.ServiceType))
                {
                    errors.Add($"{row.DocNumber}: не определена услуга по номенклатуре (доставка/сборка)");
                    continue;
                }

                if (row.Source.AgentInfo is null && !row.Source.IsOwnService)
                {
                    var city = string.IsNullOrWhiteSpace(row.City) ? "подразделение не заполнено" : row.City;
                    errors.Add($"{row.DocNumber}: для услуги «{row.Source.ServiceType}» и подразделения «{city}» не найдена ставка НДС");
                }
            }

            return errors;
        });
    }

    private static bool IsRealizationVatReady(OneCRealizationViewModel row) =>
        !row.Source.IsService ||
        (!string.IsNullOrWhiteSpace(row.Source.ServiceType) &&
         (row.Source.AgentInfo is not null || row.Source.IsOwnService));

    private void ShowRealizationVatErrors(List<string> errors, bool stopped)
    {
        var title = stopped
            ? "Не удалось определить ставку НДС для выбранных агентских реализаций.\n\n"
            : "Часть агентских реализаций пропущена: не удалось определить ставку НДС.\n\n";
        var text = title +
                   string.Join("\n", errors.Take(15));
        if (errors.Count > 15)
            text += $"\n...и еще {errors.Count - 15}";

        AtolStatus = stopped
            ? "Операция остановлена: не определена услуга/ставка НДС"
            : "Часть реализаций пропущена: не определена услуга/ставка НДС";
        StatusText = AtolStatus;
        ShowToast(stopped ? "Не определена услуга/ставка НДС" : "Часть строк пропущена", true);
        MessageBox.Show(text, "Проверка реализаций из 1С",
            MessageBoxButton.OK, MessageBoxImage.Warning);
    }

    private static OrderEntry BuildCorrectionOrderFromRealization(OneCRealizationViewModel r)
    {
        var entry = new OrderEntry
        {
            OrderNum         = r.DocNumber,
            OrderDate        = r.DocDate,
            Amount           = r.Amount,
            CustomerName     = r.CustomerName,
            AgentInfo        = r.Source.AgentInfo,
            CorrectionDate   = r.DocDate,
            CorrectionNumber = r.DocNumber,
            IsService        = r.IsService,
            IsOwnService     = r.Source.IsOwnService,
            ServiceType      = r.Source.ServiceType,
            City             = r.City,
            DocumentType     = SourceDocumentType.Realization,
            CorrectAmount    = r.Amount,
            Items            = r.Source.Items.Select(i => new OrderItem
            {
                Name     = i.Name,
                Quantity = i.Quantity,
                Sum      = i.Sum,
            }).ToList(),
            Notes            = string.IsNullOrWhiteSpace(r.OrderNumber)
            ? "Реализация из 1С без фискального чека — формируется чек коррекции"
            : $"Реализация из 1С без фискального чека — заказ покупателя {r.OrderNumber}",
        };

        if (r.HasCheck)
        {
            entry.OriginalFiscalNumber = r.FiscalNumber;
            entry.OriginalCheckAmount  = r.Amount;
            entry.CorrectionScenario   = CorrectionScenario.WrongDate;
            entry.Kind                 = OrderKind.RefundCorrectionPair;
            entry.Notes                = string.IsNullOrWhiteSpace(r.OrderNumber)
                ? "Реализация из 1С пробита не в день реализации — исправительный комплект"
                : $"Реализация из 1С пробита не в день реализации — заказ покупателя {r.OrderNumber}";
        }
        else
        {
            entry.CorrectionScenario = CorrectionScenario.CheckNotPunched;
            entry.Kind               = OrderKind.SingleCorrection;
        }

        return entry;
    }

    private static OrderEntry BuildApiSellOrderFromRealization(OneCRealizationViewModel r) => new()
    {
        OrderNum         = string.IsNullOrWhiteSpace(r.OrderNumber) ? r.DocNumber : r.OrderNumber,
        OrderDate        = string.IsNullOrWhiteSpace(r.OrderDate) ? r.DocDate : r.OrderDate,
        Amount           = r.Amount,
        CustomerName     = r.CustomerName,
        Items            = r.Source.Items.Select(i => new OrderItem
        {
            Name     = i.Name,
            Quantity = i.Quantity,
            Sum      = i.Sum,
        }).ToList(),
        AgentInfo        = r.Source.AgentInfo,
        CorrectionDate   = r.DocDate,
        CorrectionNumber = r.DocNumber,
        IsService        = r.IsService,
        IsOwnService     = r.Source.IsOwnService,
        ServiceType      = r.Source.ServiceType,
        City             = r.City,
        DocumentType     = SourceDocumentType.Realization,
    };

    private async Task PunchViaAtolAsync()
    {
        if (CheckType != "sell")
        {
            ShowToast("Через API по реализациям доступен только тип «Приход»", true);
            return;
        }

        var rows = LoadedRealizations
            .Where(r => r.IsSelected && !r.HasCheck)
            .ToList();

        if (rows.Count == 0)
        { ShowToast("Нет выбранных реализаций без чека для пробития прихода", true); return; }

        if (string.IsNullOrWhiteSpace(AtolGroupCode) || string.IsNullOrWhiteSpace(AtolLogin))
        { ShowToast("Заполните настройки АТОЛ (кнопка 🔑 в шапке)", true); return; }

        AtolStatus = $"Догружаем ставки НДС из 1С для {rows.Count} реализаций…";
        StatusText = AtolStatus;
        var errors = await PrepareOneCRealizationsAsync(rows);
        if (errors.Count > 0)
        {
            ShowRealizationVatErrors(errors, stopped: false);
        }

        rows = rows.Where(IsRealizationVatReady).ToList();
        if (rows.Count == 0)
        {
            AtolStatus = "Пробитие остановлено: нет реализаций с определенной ставкой НДС";
            StatusText = AtolStatus;
            return;
        }

        var creds = new AtolCredentials
        {
            Login     = AtolLogin.Trim(),
            Password  = AtolPassword,
            GroupCode = AtolGroupCode.Trim(),
        };

        AtolStatus = $"Пробиваем приходы по реализациям 0 из {rows.Count}...";
        var errorsText = new System.Text.StringBuilder();
        int ok = 0, fail = 0;

        for (int i = 0; i < rows.Count; i++)
        {
            var row = rows[i];
            AtolStatus = $"Пробиваем приход {i + 1}/{rows.Count}: {row.DocNumber}...";
            row.PunchStatus = "⏳";

            var order = BuildApiSellOrderFromRealization(row);
            var result = await AtolApiService.PunchOrderAsync(creds, order, "sell", "advance", "realization",
                SelectedCashier?.FullName ?? AppConstants.CashierName);

            if (result.Success)
            {
                ok++;
                row.PunchOk = true;
                row.PunchFail = false;
                row.PunchStatus = $"✅ {result.Uuid?[..8]}...";
                row.MarkAsPunched(result);
                OnPropertyChanged(nameof(SelectedNoCheckCount));
                OnPropertyChanged(nameof(SelectedHasCheckCount));
                OnPropertyChanged(nameof(CanPunchSelectedRealizationsViaAtol));
            }
            else
            {
                fail++;
                row.PunchOk = false;
                row.PunchFail = true;
                row.PunchStatus = $"❌ {result.Error}";
                errorsText.AppendLine($"❌ {row.DocNumber}: {result.Error}");
            }
        }

        AtolStatus = $"Готово: приходов пробито {ok}, ошибок {fail}";
        AtolLastError = errorsText.ToString().TrimEnd();
        StatusText = "Готов к работе";
        ShowToast($"АТОЛ Online: пробито {ok} из {rows.Count}" + (fail > 0 ? $", ошибок {fail}" : ""),
            fail > 0 && ok == 0);
    }

    // ── Пробить заказы из основного списка ───────────────────────────────────
    private async Task PunchOrdersViaAtolAsync()
    {
        if (Orders.Count == 0)
        { ShowToast("Нет заказов для пробития", true); return; }

        if (Orders.Any(IsXmlOnlyOrder))
        {
            ShowToast("Исправления из Obsidian формируются только в XML на отдельном экране", true);
            return;
        }

        if (string.IsNullOrWhiteSpace(AtolGroupCode) || string.IsNullOrWhiteSpace(AtolLogin))
        { ShowToast("Заполните настройки АТОЛ (кнопка 🔑 в шапке)", true); return; }

        var creds = new AtolCredentials
        {
            Login     = AtolLogin.Trim(),
            Password  = AtolPassword,
            GroupCode = AtolGroupCode.Trim(),
        };

        var orders = Orders.ToList();
        AtolLastError = string.Empty;
        AtolStatus    = $"Пробиваем через АТОЛ 0 из {orders.Count}...";

        var errors = new System.Text.StringBuilder();
        int ok = 0, fail = 0;

        for (int i = 0; i < orders.Count; i++)
        {
            var order = orders[i];
            AtolStatus = $"Пробиваем {i + 1} из {orders.Count}: {order.OrderNum}...";
            StatusText = AtolStatus;

            var result = await AtolApiService.PunchOrderAsync(creds, order, CheckType, PaymentType, Tab,
                SelectedCashier?.FullName ?? AppConstants.CashierName);

            if (result.Success)
            {
                ok++;
            }
            else
            {
                fail++;
                var uuidHint = !string.IsNullOrEmpty(result.Uuid) ? $" [UUID: {result.Uuid}]" : "";
                errors.AppendLine($"❌ {order.OrderNum}: {result.Error}{uuidHint}");
            }
        }

        AtolStatus    = $"Готово: пробито {ok}, ошибок {fail}" +
                        (fail > 0 ? $"  |  лог: {AtolApiService.LogPath}" : "");
        AtolLastError = errors.ToString().TrimEnd();
        StatusText    = "Готов к работе";

        ShowToast(
            $"АТОЛ Online: пробито {ok} из {orders.Count}" + (fail > 0 ? $", ошибок {fail}" : ""),
            fail > 0 && ok == 0);
    }

    private OneCConnectionSettings BuildOneCSettings() => new()
    {
        Server   = OneCServer.Trim(),
        Database = OneCDatabase.Trim(),
        User     = OneCUser.Trim(),
        Password = OneCPassword,
    };

    private async Task TestOneCAsync()
    {
        if (string.IsNullOrWhiteSpace(OneCServer) || string.IsNullOrWhiteSpace(OneCDatabase))
        { OneCStatus = "Заполните сервер и базу данных"; return; }

        OneCStatus = "Подключение...";
        var settings = BuildOneCSettings();
        var result = await Task.Run(() => OneCService.TestConnection(settings));
        OneCStatus = result;
    }

    private async Task LoadFromOneCAsync()
    {
        if (string.IsNullOrWhiteSpace(OneCServer) || string.IsNullOrWhiteSpace(OneCDatabase))
        { ShowToast("Заполните настройки 1С", true); return; }

        if (!DateTime.TryParseExact(OneCFromDate, "dd.MM.yyyy",
                System.Globalization.CultureInfo.InvariantCulture,
                System.Globalization.DateTimeStyles.None, out var from))
        { ShowToast("Неверный формат даты «от»", true); return; }

        if (!DateTime.TryParseExact(OneCToDate, "dd.MM.yyyy",
                System.Globalization.CultureInfo.InvariantCulture,
                System.Globalization.DateTimeStyles.None, out var to))
        { ShowToast("Неверный формат даты «до»", true); return; }

        OneCStatus = "Загрузка данных из 1С...";
        StatusText = "Загрузка из 1С...";
        var settings = BuildOneCSettings();
        settings.Save();

        List<OneCRealization> realizations;
        try
        {
            realizations = await Task.Run(() => OneCService.LoadRealizations(settings, from, to));
        }
        catch (Exception ex)
        {
            OneCStatus = $"Ошибка: {ex.Message} — подробности в {OneCService.LogPath}";
            ShowToast($"Ошибка загрузки из 1С. См. лог: {OneCService.LogPath}", true);
            StatusText = "Готов к работе";
            return;
        }

        // Fill staging table
        LoadedRealizations.Clear();
        foreach (var r in realizations)
        {
            var vm = new OneCRealizationViewModel(r);
            // Все строки выбираются по умолчанию.
            // Реализации из 1С без фискального чека формируются как XML чеков коррекции.
            vm.PropertyChanged += (_, _) =>
            {
                OnPropertyChanged(nameof(SelectedOneCCount));
                OnPropertyChanged(nameof(SelectedNoCheckCount));
                OnPropertyChanged(nameof(SelectedHasCheckCount));
                OnPropertyChanged(nameof(CanPunchSelectedRealizationsViaAtol));
            };
            LoadedRealizations.Add(vm);
        }

        ShowLoadedRealizations = LoadedRealizations.Count > 0;
        OnPropertyChanged(nameof(SelectedOneCCount));
        OnPropertyChanged(nameof(SelectedNoCheckCount));
        OnPropertyChanged(nameof(SelectedHasCheckCount));
        OnPropertyChanged(nameof(CanPunchSelectedRealizationsViaAtol));
        OnPropertyChanged(nameof(TotalOneCLoaded));

        var withCheck    = realizations.Count(r => r.HasCheck);
        var withoutCheck = realizations.Count - withCheck;
        OneCStatus = $"Загружено {realizations.Count}: без чека — {withoutCheck}, уже пробиты — {withCheck}";
        ShowToast(OneCStatus, false);
        StatusText = "Готов к работе";
    }

    private void SetAllOneCSelected(bool value)
    {
        foreach (var r in LoadedRealizations)
            r.IsSelected = value;
        OnPropertyChanged(nameof(SelectedOneCCount));
        OnPropertyChanged(nameof(SelectedNoCheckCount));
        OnPropertyChanged(nameof(SelectedHasCheckCount));
        OnPropertyChanged(nameof(CanPunchSelectedRealizationsViaAtol));
    }

    private async Task AddSelectedToOrdersAsync()
    {
        var selected = LoadedRealizations.Where(r => r.IsSelected).ToList();
        if (selected.Count == 0)
        {
            ShowToast("Нет выбранных реализаций для добавления", true);
            return;
        }

        OneCStatus = $"Догружаем ставки НДС из 1С для {selected.Count} реализаций…";
        var errors = await PrepareOneCRealizationsAsync(selected);
        if (errors.Count > 0)
        {
            ShowRealizationVatErrors(errors, stopped: false);
        }

        selected = selected.Where(IsRealizationVatReady).ToList();
        if (selected.Count == 0)
        {
            OneCStatus = "Добавление остановлено: нет реализаций с определенной ставкой НДС";
            return;
        }

        var existing = new HashSet<string>(Orders.Select(o =>
            !string.IsNullOrWhiteSpace(o.CorrectionNumber) ? o.CorrectionNumber : o.OrderNum));
        int addedCorrect  = 0;

        foreach (var r in selected)
        {
            var key = r.DocNumber;
            if (existing.Contains(key)) continue;

            var entry = BuildCorrectionOrderFromRealization(r);
            addedCorrect++;

            Orders.Add(entry);
            existing.Add(key);
        }

        // OrderCount / CanPunchOrdersViaAtol обновятся через Orders.CollectionChanged
        OneCStatus = addedCorrect > 0
            ? $"Добавлено в список коррекций: {addedCorrect}"
            : "Нет новых реализаций для добавления";
        if (addedCorrect > 0)
            ShowToast($"Добавлено коррекций: {addedCorrect}", false);
        else
            ShowToast("Нет новых реализаций для добавления", true);
    }

    private async Task GenerateChecksAsync()
    {
        if (Orders.Count == 0)
        { ShowToast("Добавьте хотя бы один заказ", true); return; }

        StatusText = "Генерация...";
        Results.Clear();
        ShowResults = false;

        var parms = new GenerationParams
        {
            Tab         = Tab,
            CheckType   = CheckType,
            PaymentType = PaymentType,
            IsService   = IsServiceProvider,
            MergeXml    = MergeXml,
            Orders      = Orders.ToList(),
            OutputDir   = FileHelper.OutputDir,
            Cashier     = SelectedCashier,
        };

        List<GenerationResult> results;
        try
        {
            results = await Task.Run(() => CheckGeneratorService.Generate(parms));
        }
        catch (Exception ex)
        {
            StatusText = "Ошибка генерации";
            ShowToast($"Ошибка: {ex.Message}", true);
            return;
        }

        ShowGenerationResults(results);
    }

    private void AddObsidianCasesToWork(IReadOnlyList<OrderEntry> entries)
    {
        CorrectionWork.SyncCashiers(AvailableCashiers, SelectedCashier);
        var (added, updated) = CorrectionWork.AddOrUpdate(entries);
        OpenCorrectionWorkWindow();
        ShowOfdToolsPanel = false;
        StatusText = updated > 0
            ? $"Исправления: добавлено {added}; обновлено {updated}"
            : $"Подготовлено исправлений: {added}";
        ShowToast(updated > 0
            ? $"Пробитие исправлений: добавлено {added}, обновлено {updated}"
            : $"Открыто исправлений: {added}", false);
    }

    public int OrderCount => Orders.Count;

    private void ShowGenerationResults(List<GenerationResult> results)
    {
        ObsidianCases.RecordGenerated(results);
        SelectedEntry = null;
        Results.Clear();
        AllResultEntries.Clear();
        foreach (var r in results)
        {
            Results.Add(r);
            AllResultEntries.Add(new ResultDisplayEntry { Single = r });
        }
        ShowResults = AllResultEntries.Count > 0;
        StatusText  = $"Сформировано {Results.Count} чек(ов)";

        if (AllResultEntries.Count > 0)
            SelectedEntry = AllResultEntries[0];

        var xmlCount  = Results.Select(r => r.XmlPath).Where(p => !string.IsNullOrEmpty(p)).Distinct().Count();
        var docxCount = Results.Select(r => r.DocxPath)
            .Where(p => !string.IsNullOrEmpty(p))
            .Distinct()
            .Count();
        var toastMsg  = docxCount > 0
            ? $"Готово! Создано {xmlCount} XML + {docxCount} DOCX"
            : $"Готово! Создано {xmlCount} XML";
        ShowToast(toastMsg, false);
    }

    private System.Windows.Threading.DispatcherTimer? _toastTimer;

    private void ShowToast(string msg, bool isError)
    {
        ToastMessage = msg;
        ToastIsError = isError;
        ToastVisible = true;

        _toastTimer?.Stop();
        _toastTimer = new System.Windows.Threading.DispatcherTimer
        {
            Interval = TimeSpan.FromSeconds(3)
        };
        _toastTimer.Tick += (_, _) => { ToastVisible = false; _toastTimer.Stop(); };
        _toastTimer.Start();
    }
}
