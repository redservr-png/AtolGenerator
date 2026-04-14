using System.Collections.ObjectModel;
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
            OnPropertyChanged(nameof(ShowPaymentType));
            OnPropertyChanged(nameof(ShowItemsSection));
        }
    }
    public bool IsCorrection      => CheckType is "sell_correction" or "buy_correction";
    public bool ShowCorrectionBox => IsCorrection;
    public bool ShowPaymentType   => IsPaymentTab;  // на вкладке реализации тип оплаты всегда 14 (аванс)
    public bool ShowBuyRefundOption => IsRealizationTab;
    public bool ShowItemsSection  => IsRealizationTab && !IsCorrection;

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
    public int SelectedOneCCount => LoadedRealizations.Count(r => r.IsSelected && !r.HasCheck);
    public int TotalOneCLoaded   => LoadedRealizations.Count;

    // ── Results ──────────────────────────────────────────────────────────────
    public ObservableCollection<GenerationResult> Results { get; } = new();

    private bool _showResults;
    public bool ShowResults { get => _showResults; set => Set(ref _showResults, value); }

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

    // ── Commands ─────────────────────────────────────────────────────────────
    public ICommand ParseBulkCommand       { get; }
    public ICommand AddSingleCommand       { get; }
    public ICommand DeleteOrderCommand     { get; }
    public ICommand GenerateCommand        { get; }
    public ICommand OpenFolderCommand      { get; }
    public ICommand AddItemCommand         { get; }
    public ICommand DeleteItemCommand      { get; }
    public ICommand SwitchTabCommand       { get; }
    public ICommand ImportExcelCommand     { get; }
    public ICommand ToggleOneCPanelCommand     { get; }
    public ICommand TestOneCCommand            { get; }
    public ICommand LoadFromOneCCommand        { get; }
    public ICommand SelectAllOneCCommand       { get; }
    public ICommand DeselectAllOneCCommand     { get; }
    public ICommand AddSelectedToOrdersCommand { get; }

    // ── Skipped rows from Excel import ───────────────────────────────────────
    public ObservableCollection<SkippedRow> SkippedRows { get; } = new();
    private bool _showSkipped;
    public bool ShowSkipped { get => _showSkipped; set => Set(ref _showSkipped, value); }

    public MainViewModel()
    {
        ParseBulkCommand   = new RelayCommand(ParseBulk);
        AddSingleCommand   = new RelayCommand(AddSingleOrder);
        DeleteOrderCommand = new RelayCommand(o => DeleteOrder(o as OrderEntry));
        GenerateCommand    = new AsyncRelayCommand(GenerateChecksAsync);
        OpenFolderCommand  = new RelayCommand(_ => FileHelper.OpenFolder(FileHelper.OutputDir));
        AddItemCommand     = new RelayCommand(_ => CurrentItems.Add(new OrderItemViewModel()));
        DeleteItemCommand  = new RelayCommand(o => { if (o is OrderItemViewModel vm) CurrentItems.Remove(vm); });
        SwitchTabCommand       = new RelayCommand(t => { if (t is string s) Tab = s; });
        ImportExcelCommand     = new RelayCommand(_ => ImportExcel());
        ToggleOneCPanelCommand     = new RelayCommand(_ => ShowOneCPanel = !ShowOneCPanel);
        TestOneCCommand            = new AsyncRelayCommand(TestOneCAsync);
        LoadFromOneCCommand        = new AsyncRelayCommand(LoadFromOneCAsync);
        SelectAllOneCCommand       = new RelayCommand(_ => SetAllOneCSelected(true));
        DeselectAllOneCCommand     = new RelayCommand(_ => SetAllOneCSelected(false));
        AddSelectedToOrdersCommand = new RelayCommand(_ => AddSelectedToOrders());
    }

    // ── Logic ─────────────────────────────────────────────────────────────────

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

    private void ParseBulk()
    {
        if (string.IsNullOrWhiteSpace(BulkText))
        { ShowToast("Введите текст с заказами", true); return; }

        var parsed = OrderParserService.Parse(BulkText);
        if (parsed.Count == 0)
        { ShowToast("Заказы не распознаны. Проверьте формат.", true); return; }

        var existing = new HashSet<string>(Orders.Select(o => o.OrderNum));
        int added = 0;
        foreach (var o in parsed)
        {
            if (!existing.Contains(o.OrderNum))
            { ApplyCorrectionFields(o); Orders.Add(o); added++; }
        }
        OnPropertyChanged(nameof(OrderCount));
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
            Items        = CurrentItems.Select(i => i.ToModel()).ToList(),
        };
        ApplyCorrectionFields(order);
        Orders.Add(order);
        OnPropertyChanged(nameof(OrderCount));
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
        }
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

        List<OneCRealization> realizations;
        try
        {
            realizations = await Task.Run(() => OneCService.LoadRealizations(settings, from, to));
        }
        catch (Exception ex)
        {
            OneCStatus = $"Ошибка: {ex.Message}";
            ShowToast($"Ошибка загрузки из 1С: {ex.Message}", true);
            StatusText = "Готов к работе";
            return;
        }

        // Fill staging table
        LoadedRealizations.Clear();
        foreach (var r in realizations)
        {
            var vm = new OneCRealizationViewModel(r);
            // Already-checked rows — pre-deselect so user can see them but won't add them
            if (r.HasCheck) vm.IsSelected = false;
            vm.PropertyChanged += (_, _) =>
            {
                OnPropertyChanged(nameof(SelectedOneCCount));
            };
            LoadedRealizations.Add(vm);
        }

        ShowLoadedRealizations = LoadedRealizations.Count > 0;
        OnPropertyChanged(nameof(SelectedOneCCount));
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
            if (!r.HasCheck) r.IsSelected = value;
        OnPropertyChanged(nameof(SelectedOneCCount));
    }

    private void AddSelectedToOrders()
    {
        var existing = new HashSet<string>(Orders.Select(o => o.OrderNum));
        var added = 0;

        foreach (var r in LoadedRealizations.Where(r => r.IsSelected && !r.HasCheck))
        {
            var key = string.IsNullOrEmpty(r.OrderNumber) ? r.DocNumber : r.OrderNumber;
            if (existing.Contains(key)) continue;

            Orders.Add(new OrderEntry
            {
                OrderNum         = r.OrderNumber,
                OrderDate        = r.OrderDate,
                Amount           = r.Amount,
                CustomerName     = r.CustomerName,
                CorrectionDate   = r.DocDate,
                CorrectionNumber = r.DocNumber,
                IsService        = r.IsService,
            });
            existing.Add(key);
            added++;
        }

        OnPropertyChanged(nameof(OrderCount));
        if (added > 0)
            ShowToast($"Добавлено {added} реализаций в очередь", false);
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

        foreach (var r in results) Results.Add(r);
        ShowResults = Results.Count > 0;
        StatusText  = $"Сформировано {Results.Count} чек(ов)";
        ShowToast($"Готово! Создано {Results.Count} XML + {Results.Count} DOCX", false);
    }

    public int OrderCount => Orders.Count;

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
