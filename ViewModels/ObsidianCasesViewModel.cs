using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Diagnostics;
using System.Globalization;
using System.IO;
using System.Windows;
using System.Windows.Data;
using System.Windows.Input;
using AtolGenerator.Helpers;
using AtolGenerator.Models;
using AtolGenerator.Services;
using AtolGenerator.Views;
using Microsoft.Win32;

namespace AtolGenerator.ViewModels;

public sealed class ObsidianCaseItemViewModel : BaseViewModel
{
    private bool _isSelected;
    private string _editableProblem;
    private CorrectionPlan _plan;

    public ObsidianCaseItemViewModel(ObsidianCaseRecord record, ObsidianCaseState state)
    {
        Record = record;
        State = state;
        if (Record.PrimaryDocument.CorrectionScenario == CorrectionScenario.RealRefund)
        {
            Record.PrimaryDocument.CorrectionScenario = CorrectionScenario.Unknown;
            Record.PrimaryDocument.Kind = OrderKind.SingleRefund;
            State.ScenarioOverride = null;
        }
        _editableProblem = record.PrimaryDocument.Notes;
        NormalizeOriginalReceiptLookupState();
        _plan = CorrectionPlanService.Build(record.PrimaryDocument, state.OriginalReceipt, DateTime.Today);
    }

    public ObsidianCaseRecord Record { get; }
    public ObsidianCaseState State { get; }
    public string CaseId => Record.CaseId;
    public bool IsCompleted => Record.IsCompleted;
    public string Period => Record.Period;
    public string PeriodLabel => string.IsNullOrWhiteSpace(Period) ? "Без раздела" : Period;
    public string DocumentNumber => Record.PrimaryDocument.OrderNum;
    public string DocumentDate => SourceDocumentDate(Record.PrimaryDocument);
    public string Department => Record.PrimaryDocument.City;
    public string DocumentType => DisplayDocumentType(Record.PrimaryDocument.DocumentType);
    public string RelatedDocumentsText => Record.RelatedDocuments.Count == 0
        ? "Нет связанных документов"
        : string.Join("\n", Record.RelatedDocuments.Select(x =>
            $"{DisplayDocumentType(x.DocumentType)} {x.OrderNum} от {SourceDocumentDate(x)}".Trim()));
    public double CorrectAmount => State.ExpectedChecks.LastOrDefault()?.Amount
                                   ?? Record.PrimaryDocument.CorrectAmount
                                   ?? Record.PrimaryDocument.Amount;
    public string AmountText => CorrectAmount > 0 ? $"{CorrectAmount:N2} ₽" : "Не загружена";
    public string OneCDataStatus => State.OneCSnapshot is null ? "Не загружены" : "Загружены";
    public string OneCDocumentDateText => State.OneCSnapshot is null
        ? "Не загружена"
        : ValueOrPlaceholder(Record.PrimaryDocument.OrderDate);
    public string OneCCheckDateText => Record.PrimaryDocument.OneCCheckDate?.ToString("dd.MM.yyyy HH:mm:ss")
                                       ?? "Не заполнена";
    public string OneCCheckNumberText => ValueOrPlaceholder(Record.PrimaryDocument.OneCCheckNumber);
    public string OneCFiscalNumberText => ValueOrPlaceholder(Record.PrimaryDocument.OriginalFiscalNumber);
    public string OneCCommentText => ValueOrPlaceholder(Record.PrimaryDocument.OneCComment);

    public bool IsSelected
    {
        get => _isSelected;
        set => Set(ref _isSelected, value);
    }

    public string EditableProblem
    {
        get => _editableProblem;
        set => Set(ref _editableProblem, value);
    }

    public CorrectionScenario SelectedScenario
    {
        get => Record.PrimaryDocument.CorrectionScenario;
        set
        {
            if (Record.PrimaryDocument.CorrectionScenario == value) return;
            Record.PrimaryDocument.CorrectionScenario = value;
            Record.PrimaryDocument.Kind = value.ToOrderKind();
            State.ScenarioOverride = value;
            OnPropertyChanged();
            OnPropertyChanged(nameof(ScenarioSourceLabel));
            Refresh();
        }
    }

    public string ScenarioSourceLabel => State.ScenarioOverride.HasValue
        ? "Выбран вручную"
        : SelectedScenario == CorrectionScenario.Unknown
            ? "Не определён по пояснению"
            : "Определён по пояснению";

    public string StateLabel
    {
        get
        {
            if (IsCompleted) return "Закрыта";
            if (State.CheckConfirmed && State.ServiceNoteVerified && !State.OneCRecorded) return "Запись в 1С";
            if (State.CheckConfirmed && !State.ServiceNoteVerified) return "Проверить служебку";
            if (State.ExpectedChecks.Count > 0) return "Ожидается отчёт";
            if (State.SentToWork) return "В работе";
            return "Новая";
        }
    }

    public string CheckStatus => IsCompleted || State.CheckConfirmed ? "Подтверждён" : "Не подтверждён";
    public string MemoStatus => IsCompleted || State.ServiceNoteVerified ? "Проверена" : "Не проверена";
    public string OneCStatus => IsCompleted || State.OneCRecorded ? "Записано" : "Не подтверждено";
    public string LastMessage => State.LastMessage;
    public bool CanEdit => !IsCompleted && !string.IsNullOrWhiteSpace(CaseId);
    public CorrectionPlan Plan => _plan;
    public bool PlanReady => Plan.IsReady;
    public string PlanStatus => Plan.Status switch
    {
        CorrectionPlanStatus.Ready => "План готов",
        CorrectionPlanStatus.NeedsOneC => "Нужны данные 1С",
        CorrectionPlanStatus.NeedsOriginalReceipt => "Нужен исходный чек",
        CorrectionPlanStatus.NeedsScenario => "Нужен сценарий",
        CorrectionPlanStatus.NeedsServiceRule => "Не определён НДС",
        CorrectionPlanStatus.DeferredCorrectionReceipt => "Этап 3",
        _ => "Требует проверки",
    };
    public string PlanMessage => Plan.Message;
    public string PlanChecksText => Plan.Checks.Count == 0
        ? "Чеки пока не рассчитаны"
        : string.Join("\n", Plan.Checks.Select(x =>
            $"{x.Sequence}. {x.Title} · {x.Operation} · {x.Amount:N2} ₽ · {VatLabel(x.VatType)}"));
    public string OriginalReceiptLookupText => State.OriginalReceiptLookupState switch
    {
        OriginalReceiptLookupState.Found => "Найден",
        OriginalReceiptLookupState.NotFound => "Не найден",
        OriginalReceiptLookupState.MissingFiscalSign => "Нет ФП",
        OriginalReceiptLookupState.Ambiguous => "Несколько совпадений",
        _ => "Не проверен",
    };
    public string OriginalReceiptStatus => State.OriginalReceiptLookupState switch
    {
        OriginalReceiptLookupState.Found when State.OriginalReceipt is not null =>
            $"ФП {State.OriginalReceipt.FiscalSign} · ФД {State.OriginalReceipt.FiscalDocument}",
        OriginalReceiptLookupState.NotFound => "Не найден",
        OriginalReceiptLookupState.MissingFiscalSign => "Нет ФП",
        OriginalReceiptLookupState.Ambiguous => "Несколько совпадений",
        _ => "Не проверен",
    };
    public string OriginalReceiptDetails => State.OriginalReceiptLookupState switch
    {
        OriginalReceiptLookupState.Found when State.OriginalReceipt is not null =>
            $"{State.OriginalReceipt.RegisteredAt:dd.MM.yyyy HH:mm} · " +
            $"{State.OriginalReceipt.Operation} · {Math.Abs(State.OriginalReceipt.Amount):N2} ₽",
        OriginalReceiptLookupState.NotFound =>
            string.IsNullOrWhiteSpace(State.LastMessage)
                ? $"ФП {Record.PrimaryDocument.OriginalFiscalNumber} не найден в ОФД."
                : State.LastMessage,
        OriginalReceiptLookupState.MissingFiscalSign => "В данных 1С не заполнен ФП исходного чека.",
        OriginalReceiptLookupState.Ambiguous =>
            $"ФП {Record.PrimaryDocument.OriginalFiscalNumber} встречается в отчёте больше одного раза.",
        _ => "Загрузите отчёт ОФД и найдите чек по ФП из 1С.",
    };

    public void Refresh()
    {
        _plan = CorrectionPlanService.Build(Record.PrimaryDocument, State.OriginalReceipt, DateTime.Today);
        OnPropertyChanged(nameof(StateLabel));
        OnPropertyChanged(nameof(CheckStatus));
        OnPropertyChanged(nameof(MemoStatus));
        OnPropertyChanged(nameof(OneCStatus));
        OnPropertyChanged(nameof(LastMessage));
        OnPropertyChanged(nameof(AmountText));
        OnPropertyChanged(nameof(OneCDataStatus));
        OnPropertyChanged(nameof(OneCDocumentDateText));
        OnPropertyChanged(nameof(OneCCheckDateText));
        OnPropertyChanged(nameof(OneCCheckNumberText));
        OnPropertyChanged(nameof(OneCFiscalNumberText));
        OnPropertyChanged(nameof(OneCCommentText));
        OnPropertyChanged(nameof(Department));
        OnPropertyChanged(nameof(DocumentDate));
        OnPropertyChanged(nameof(Plan));
        OnPropertyChanged(nameof(PlanReady));
        OnPropertyChanged(nameof(PlanStatus));
        OnPropertyChanged(nameof(PlanMessage));
        OnPropertyChanged(nameof(PlanChecksText));
        OnPropertyChanged(nameof(OriginalReceiptLookupText));
        OnPropertyChanged(nameof(OriginalReceiptStatus));
        OnPropertyChanged(nameof(OriginalReceiptDetails));
        OnPropertyChanged(nameof(ScenarioSourceLabel));
    }

    private void NormalizeOriginalReceiptLookupState()
    {
        if (State.OriginalReceiptLookupState != OriginalReceiptLookupState.NotChecked) return;

        if (State.OriginalReceipt is not null)
        {
            State.OriginalReceiptLookupState = OriginalReceiptLookupState.Found;
            return;
        }

        if (State.LastMessage?.Contains("не найден в отчёте ОФД", StringComparison.OrdinalIgnoreCase) == true)
            State.OriginalReceiptLookupState = OriginalReceiptLookupState.NotFound;
        else if (State.LastMessage?.Contains("не заполнен ФП исходного чека", StringComparison.OrdinalIgnoreCase) == true)
            State.OriginalReceiptLookupState = OriginalReceiptLookupState.MissingFiscalSign;
        else if (State.LastMessage?.Contains("больше одного раза", StringComparison.OrdinalIgnoreCase) == true)
            State.OriginalReceiptLookupState = OriginalReceiptLookupState.Ambiguous;
    }

    public OrderEntry CreateWorkEntry()
    {
        var source = Record.PrimaryDocument;
        var entry = new OrderEntry
        {
            ObsidianCaseId = CaseId,
            OrderNum = source.OrderNum,
            OrderDate = source.OrderDate,
            SourceDocumentDate = source.SourceDocumentDate,
            Amount = source.Amount,
            CustomerName = source.CustomerName,
            Items = source.Items.Select(x => new OrderItem
            {
                Name = x.Name,
                Quantity = x.Quantity,
                Sum = x.Sum,
                VatType = x.VatType,
            }).ToList(),
            OriginalItems = (source.OriginalItems.Count > 0 ? source.OriginalItems : source.Items)
                .Select(x => new OrderItem
                {
                    Name = x.Name,
                    Quantity = x.Quantity,
                    Sum = x.Sum,
                    VatType = x.VatType,
                }).ToList(),
            AgentInfo = source.AgentInfo,
            CorrectionDate = source.CorrectionDate,
            CorrectionNumber = source.CorrectionNumber,
            IsService = source.IsService,
            IsOwnService = source.IsOwnService,
            ServiceType = source.ServiceType,
            City = source.City,
            DocumentType = source.DocumentType,
            CorrectionScenario = source.CorrectionScenario,
            OriginalFiscalNumber = source.OriginalFiscalNumber,
            OriginalCheckAmount = source.OriginalCheckAmount,
            CorrectAmount = source.CorrectAmount,
            Notes = EditableProblem.Trim(),
            OriginalPaymentWasCash = source.OriginalPaymentWasCash,
            CorrectPaymentIsCash = source.CorrectPaymentIsCash,
            OriginalCheckDate = State.OriginalReceipt?.RegisteredAt,
            OriginalCheckOperation = State.OriginalReceipt?.Operation ?? string.Empty,
            OneCComment = source.OneCComment,
            OneCCheckNumber = source.OneCCheckNumber,
            OneCCheckDate = source.OneCCheckDate,
            PlannedReverseOperation = Plan.Checks.Count > 1 ||
                                      source.CorrectionScenario == CorrectionScenario.FullCancel
                ? Plan.Checks.FirstOrDefault()?.Operation ?? string.Empty
                : string.Empty,
            PlannedCorrectOperation = source.CorrectionScenario == CorrectionScenario.FullCancel
                ? string.Empty
                : Plan.Checks.LastOrDefault()?.Operation ?? string.Empty,
            OriginalVatType = source.OriginalVatType,
            CorrectVatType = source.CorrectVatType,
            PlannedVatType = Plan.Checks.LastOrDefault()?.VatType ?? string.Empty,
        };
        entry.OriginalCheckAmount = State.OriginalReceipt is not null
            ? Math.Abs(State.OriginalReceipt.Amount)
            : source.OriginalCheckAmount;
        entry.OriginalFiscalNumber = State.OriginalReceipt?.FiscalSign?.ToString()
                                     ?? source.OriginalFiscalNumber;
        entry.Kind = Plan.Checks.Count > 1
            ? OrderKind.RefundCorrectionPair
            : Plan.Checks.Any(x => x.Operation.EndsWith("_correction", StringComparison.OrdinalIgnoreCase))
                ? OrderKind.SingleCorrection
                : OrderKind.SingleRefund;
        return entry;
    }

    private static string VatLabel(string value) => VatRateCatalog.LabelFor(value);

    private static string SourceDocumentDate(OrderEntry entry) =>
        string.IsNullOrWhiteSpace(entry.SourceDocumentDate)
            ? entry.OrderDate
            : entry.SourceDocumentDate;

    private static string ValueOrPlaceholder(string value) =>
        string.IsNullOrWhiteSpace(value) ? "Не заполнено" : value;

    private static string DisplayDocumentType(SourceDocumentType type) => type switch
    {
        SourceDocumentType.Realization => "Реализация",
        SourceDocumentType.CardPayment => "Оплата картой",
        SourceDocumentType.CashPayment => "ПКО / наличные",
        SourceDocumentType.CashExpense => "РКО / возврат",
        SourceDocumentType.BuyerOrder => "Заказ покупателя",
        SourceDocumentType.KkmCheck => "Чек ККМ",
        SourceDocumentType.FpOnly => "Фискальный признак",
        _ => "Не распознано",
    };
}

public sealed class ObsidianCasesViewModel : BaseViewModel, IDisposable
{
    private readonly Dictionary<string, ObsidianCaseState> _states = ObsidianCaseStateStore.Load();
    private FileSystemWatcher? _watcher;
    private DateTime _ignoreWatcherUntil;
    private string _filePath = string.Empty;
    private string _searchText = string.Empty;
    private string _selectedPeriod = "Все периоды";
    private int _selectedPeriodIndex;
    private bool _showClosed;
    private string _status = "Укажите Obsidian-файл в настройках";
    private ObsidianCaseItemViewModel? _selectedCase;
    private IReadOnlyList<OfdReportRow> _sourceOfdRows = Array.Empty<OfdReportRow>();
    private string _sourceOfdReportPath = string.Empty;

    public ObsidianCasesViewModel()
    {
        CasesView = CollectionViewSource.GetDefaultView(Cases);
        CasesView.Filter = FilterCase;
        SyncCommand = new RelayCommand(Sync);
        SendToWorkCommand = new RelayCommand(SendToWork, HasSelectedReady);
        VerifyMemosCommand = new RelayCommand(VerifyMemos, HasSelectedActive);
        MarkMemoVerifiedCommand = new RelayCommand(MarkMemoVerified, HasSelectedActive);
        VerifyReportCommand = new RelayCommand(VerifyReport, HasSelectedActive);
        FetchOneCDataCommand = new AsyncRelayCommand(FetchOneCDataAsync);
        LoadSourceOfdReportCommand = new RelayCommand(LoadSourceOfdReport);
        MarkOneCRecordedCommand = new RelayCommand(MarkOneCRecorded, HasSelectedActive);
        SaveProblemCommand = new RelayCommand(SaveProblem, () => SelectedCase?.CanEdit == true);
        OpenServiceNoteCommand = new RelayCommand(OpenServiceNote, () =>
            SelectedCase is not null && File.Exists(SelectedCase.State.ServiceNotePath));
        SelectAllVisibleCommand = new RelayCommand(SelectAllVisible);
        SelectCaseCommand = new RelayCommand(SelectCase);
        ReloadSettings();
    }

    public event Action<IReadOnlyList<OrderEntry>>? SendToWorkRequested;
    public Func<(string Path, IReadOnlyList<OfdReportRow> Rows)>? OfdReportProvider { get; set; }
    public Action<IReadOnlyList<OfdReportRow>>? OnlineOfdRowsImported { get; set; }

    public ObservableCollection<ObsidianCaseItemViewModel> Cases { get; } = new();
    public ObservableCollection<string> PeriodOptions { get; } = new() { "Все периоды" };
    public CorrectionScenario[] ScenarioOptions { get; } = Enum
        .GetValues<CorrectionScenario>()
        .Where(x => x != CorrectionScenario.RealRefund)
        .ToArray();
    public ICollectionView CasesView { get; }
    public ICommand SyncCommand { get; }
    public ICommand SendToWorkCommand { get; }
    public ICommand VerifyMemosCommand { get; }
    public ICommand MarkMemoVerifiedCommand { get; }
    public ICommand VerifyReportCommand { get; }
    public ICommand FetchOneCDataCommand { get; }
    public ICommand LoadSourceOfdReportCommand { get; }
    public ICommand MarkOneCRecordedCommand { get; }
    public ICommand SaveProblemCommand { get; }
    public ICommand OpenServiceNoteCommand { get; }
    public ICommand SelectAllVisibleCommand { get; }
    public ICommand SelectCaseCommand { get; }

    public string FilePath { get => _filePath; private set => Set(ref _filePath, value); }
    public string SearchText
    {
        get => _searchText;
        set
        {
            if (!Set(ref _searchText, value)) return;
            CasesView.Refresh();
        }
    }
    public string SelectedPeriod
    {
        get => _selectedPeriod;
        set
        {
            if (!Set(ref _selectedPeriod, value ?? "Все периоды")) return;
            CasesView.Refresh();
        }
    }
    public int SelectedPeriodIndex
    {
        get => _selectedPeriodIndex;
        set
        {
            if (value < 0 || value >= PeriodOptions.Count || !Set(ref _selectedPeriodIndex, value)) return;
            var period = PeriodOptions[value];
            if (string.Equals(_selectedPeriod, period, StringComparison.Ordinal)) return;
            _selectedPeriod = period;
            OnPropertyChanged(nameof(SelectedPeriod));
            CasesView.Refresh();
        }
    }
    public bool ShowClosed
    {
        get => _showClosed;
        set
        {
            if (!Set(ref _showClosed, value)) return;
            CasesView.Refresh();
        }
    }
    public string Status { get => _status; private set => Set(ref _status, value); }
    public ObsidianCaseItemViewModel? SelectedCase
    {
        get => _selectedCase;
        set => Set(ref _selectedCase, value);
    }
    public int ActiveCount => Cases.Count(x => !x.IsCompleted);
    public int ClosedCount => Cases.Count(x => x.IsCompleted);
    public int SelectedCount => Cases.Count(x => x.IsSelected && !x.IsCompleted);
    public string SourceOfdReportInfo
    {
        get
        {
            if (string.IsNullOrWhiteSpace(_sourceOfdReportPath))
                return "Архив исходных чеков не загружен";

            var fileCount = Math.Max(1, _sourceOfdRows
                .Select(row => row.SourceFile)
                .Where(name => !string.IsNullOrWhiteSpace(name))
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .Count());
            return $"ОФД: архив · файлов {fileCount} · чеков {_sourceOfdRows.Count}";
        }
    }

    public void Activate()
    {
        ReloadSettings();
        TryUseWorkspaceOfdReport();
        Sync();
    }

    public void ReloadSettings()
    {
        var settings = ApplicationSettingsStore.Reload();
        var path = settings.ObsidianFilePath;
        if (string.IsNullOrWhiteSpace(path))
            path = ObsidianSettings.Load().MdFilePath;
        FilePath = path;
        ConfigureWatcher();
    }

    public void RecordGenerated(IEnumerable<GenerationResult> results)
    {
        var changed = false;
        foreach (var result in results.Where(x => !string.IsNullOrWhiteSpace(x.ObsidianCaseId)))
        {
            var state = GetState(result.ObsidianCaseId);
            state.SentToWork = true;
            if (!string.IsNullOrWhiteSpace(result.DocxPath)) state.ServiceNotePath = result.DocxPath;
            if (!string.IsNullOrWhiteSpace(result.ExternalId) &&
                state.ExpectedChecks.All(x => !x.ExternalId.Equals(result.ExternalId, StringComparison.OrdinalIgnoreCase)))
            {
                state.ExpectedChecks.Add(new ObsidianExpectedCheck
                {
                    ExternalId = result.ExternalId,
                    Operation = result.CheckData?.OperationType ?? string.Empty,
                    Amount = result.Amount,
                    GeneratedAt = DateTime.Now,
                });
            }
            state.LastMessage = "Документы сформированы, ожидается подтверждение отчётом";
            state.UpdatedAt = DateTime.Now;
            changed = true;

            if (ApplicationSettingsStore.Current.AutoValidateServiceNotes)
                ValidateMemo(state, result.OrderNum, result.Amount);
        }

        if (!changed) return;
        SaveStates();
        Sync();
    }

    private void Sync()
    {
        if (string.IsNullOrWhiteSpace(FilePath))
        {
            Cases.Clear();
            Status = "Укажите путь к файлу Obsidian в настройках";
            RefreshCounters();
            return;
        }
        if (!File.Exists(FilePath))
        {
            Cases.Clear();
            Status = "Файл Obsidian не найден";
            RefreshCounters();
            return;
        }

        try
        {
            _ignoreWatcherUntil = DateTime.Now.AddSeconds(1);
            var selectedId = SelectedCase?.CaseId;
            var markedCaseIds = Cases
                .Where(x => x.IsSelected && !string.IsNullOrWhiteSpace(x.CaseId))
                .Select(x => x.CaseId)
                .ToHashSet(StringComparer.OrdinalIgnoreCase);
            var records = ObsidianSyncService.LoadAndEnsureIds(FilePath);
            Cases.Clear();
            foreach (var record in records)
            {
                record.PrimaryDocument.SourceDocumentDate = record.PrimaryDocument.OrderDate;
                foreach (var related in record.RelatedDocuments)
                    related.SourceDocumentDate = related.OrderDate;
                var state = GetState(record.CaseId);
                if (state.OneCSnapshot is not null)
                    ApplySnapshot(state.OneCSnapshot, record.PrimaryDocument);
                if (state.CorrectAmount is not null)
                {
                    record.PrimaryDocument.CorrectAmount = state.CorrectAmount;
                    record.PrimaryDocument.Amount = state.CorrectAmount.Value;
                }
                if (!string.IsNullOrWhiteSpace(state.OriginalFiscalNumber))
                    record.PrimaryDocument.OriginalFiscalNumber = state.OriginalFiscalNumber;
                if (record.ScenarioOverride.HasValue)
                {
                    state.ScenarioOverride = record.ScenarioOverride;
                    record.PrimaryDocument.CorrectionScenario = record.ScenarioOverride.Value;
                    record.PrimaryDocument.Kind = record.ScenarioOverride.Value.ToOrderKind();
                }
                else if (state.ScenarioOverride.HasValue)
                {
                    record.PrimaryDocument.CorrectionScenario = state.ScenarioOverride.Value;
                    record.PrimaryDocument.Kind = state.ScenarioOverride.Value.ToOrderKind();
                }
                else
                {
                    CorrectionTypeDetector.Detect(record.PrimaryDocument);
                }
                if (record.IsCompleted)
                {
                    state.CheckConfirmed = true;
                    state.ServiceNoteVerified = true;
                    state.OneCRecorded = true;
                }
                var item = new ObsidianCaseItemViewModel(record, state);
                item.IsSelected = !record.IsCompleted && markedCaseIds.Contains(record.CaseId);
                item.PropertyChanged += (_, args) =>
                {
                    if (args.PropertyName == nameof(ObsidianCaseItemViewModel.IsSelected))
                        RefreshCounters();
                    if (args.PropertyName == nameof(ObsidianCaseItemViewModel.SelectedScenario))
                    {
                        _ignoreWatcherUntil = DateTime.Now.AddSeconds(1);
                        ObsidianSyncService.UpdateScenario(FilePath, item.CaseId, item.SelectedScenario);
                        SaveStates();
                        Status = $"Сценарий для {item.DocumentNumber} сохранён в Obsidian";
                        RefreshCounters();
                    }
                };
                Cases.Add(item);
            }
            RefreshPeriodOptions();
            SelectedCase = Cases.FirstOrDefault(x => x.CaseId == selectedId)
                           ?? Cases.FirstOrDefault(x => !x.IsCompleted)
                           ?? Cases.FirstOrDefault();
            CasesView.Refresh();
            Status = $"Синхронизировано: активных {ActiveCount}, закрытых {ClosedCount}";
            SaveStates();
            RefreshCounters();
        }
        catch (Exception ex)
        {
            Status = $"Ошибка синхронизации: {ex.Message}";
        }
    }

    private void SendToWork()
    {
        var allSelected = SelectedActive().ToList();
        var selected = allSelected.Where(x => x.PlanReady).ToList();
        if (selected.Count == 0) return;
        var entries = selected.Select(x => x.CreateWorkEntry()).ToList();
        foreach (var item in selected)
        {
            item.State.SentToWork = true;
            item.State.LastMessage = "Открыто в пробитии исправлений";
            item.State.UpdatedAt = DateTime.Now;
            item.Refresh();
        }
        SaveStates();
        var skipped = allSelected.Count - selected.Count;
        Status = skipped == 0
            ? $"Открыто в пробитии исправлений: {entries.Count}"
            : $"Открыто в пробитии исправлений: {entries.Count}; требуют подготовки: {skipped}";
        SendToWorkRequested?.Invoke(entries);
    }

    private void VerifyMemos()
    {
        var selected = SelectedActive().ToList();
        var verified = 0;
        foreach (var item in selected)
        {
            if (ValidateMemo(item.State, item.DocumentNumber, item.CorrectAmount)) verified++;
            item.Refresh();
        }
        SaveStates();
        TryCompleteSelected(selected);
        Status = $"Служебки проверены: {verified} из {selected.Count}";
    }

    private void MarkMemoVerified()
    {
        var selected = SelectedActive().ToList();
        foreach (var item in selected)
        {
            item.State.ServiceNoteVerified = true;
            item.State.LastMessage = "Служебная записка подтверждена вручную";
            item.State.UpdatedAt = DateTime.Now;
            item.Refresh();
        }
        SaveStates();
        TryCompleteSelected(selected);
        Status = $"Служебки подтверждены вручную: {selected.Count}";
    }

    private async Task FetchOneCDataAsync()
    {
        var selected = SelectedActive().ToList();
        if (selected.Count == 0)
        {
            Status = "Выберите случаи для загрузки из 1С";
            return;
        }

        var settings = OneCConnectionSettings.Load();
        if (string.IsNullOrWhiteSpace(settings.Server) || string.IsNullOrWhiteSpace(settings.Database))
        {
            Status = "Сначала заполните подключение к 1С в настройках";
            return;
        }

        var targets = selected
            .Select(x => x.Record.PrimaryDocument)
            .Concat(selected.SelectMany(x => x.Record.RelatedDocuments))
            .Where(x => !string.IsNullOrWhiteSpace(x.OrderNum) &&
                        x.DocumentType is not (SourceDocumentType.Unknown or SourceDocumentType.FpOnly or SourceDocumentType.KkmCheck))
            .ToList();
        if (targets.Count == 0)
        {
            Status = "В выбранных строках нет документов, доступных для запроса 1С";
            return;
        }

        Status = $"Загружаем данные из 1С для {targets.Count} документов...";
        var result = await Task.Run(() =>
        {
            var fetch = OneCService.FetchAmountsFromOneC(settings, targets);
            foreach (var target in targets.Where(x =>
                         x.DocumentType == SourceDocumentType.Realization && x.CorrectAmount is not null))
            {
                try
                {
                    OneCService.EnrichCorrectionOrderForReceipt(settings, target);
                }
                catch (Exception ex)
                {
                    fetch.Errors.Add($"{target.OrderNum}: табличная часть/услуга не загружена ({ex.Message})");
                }
            }
            return fetch;
        });
        foreach (var item in selected)
        {
            var source = item.Record.PrimaryDocument;
            if (source.CorrectAmount is not null) item.State.CorrectAmount = source.CorrectAmount;
            if (!string.IsNullOrWhiteSpace(source.OriginalFiscalNumber))
                item.State.OriginalFiscalNumber = source.OriginalFiscalNumber;
            item.State.OneCSnapshot = CloneSnapshot(source);
            item.State.LastMessage = source.CorrectAmount is not null
                ? "Правильные сумма, дата и реквизиты загружены из 1С"
                : "Документ не найден в 1С";
            item.State.UpdatedAt = DateTime.Now;
            item.Refresh();
        }
        SaveStates();
        Sync();
        Status = $"Из 1С загружено: {result.Filled} из {result.Total}; не найдено: {result.NotFound}" +
                 (result.Errors.Count > 0 ? $"; предупреждений: {result.Errors.Count}" : string.Empty);
    }

    private void LoadSourceOfdReport()
    {
        var selected = SelectedActive().ToList();
        if (selected.Count == 0)
        {
            Status = "Отметьте строки, для которых нужно найти исходные чеки";
            return;
        }

        try
        {
            if (!TryUseWorkspaceOfdReport() && _sourceOfdRows.Count == 0)
            {
                var archive = ReportsViewModel.TryReadOfdArchive(
                    ApplicationSettingsStore.Current.LastOfdReportPath);
                if (archive is not null)
                    SetSourceOfdReport(archive.ImportedFiles[0], archive.Rows);
            }

            var summary = MatchOriginalReceipts(selected);
            var onlineTargets = selected
                .Where(item => item.State.OriginalReceiptLookupState == OriginalReceiptLookupState.NotFound)
                .Select(item => (Item: item, FiscalSign: ParseFiscalSign(item.Record.PrimaryDocument.OriginalFiscalNumber)))
                .Where(x => x.FiscalSign.HasValue)
                .GroupBy(x => x.FiscalSign!.Value)
                .Select(group => CreateTaxcomRequest(group.First().Item, group.Key))
                .ToList();

            var onlineFound = 0;
            var onlineFailed = 0;
            if (onlineTargets.Count > 0)
            {
                Status = $"В локальном архиве нет {onlineTargets.Count} чеков. Открываем поиск Такскома...";
                var searchWindow = new TaxcomReceiptSearchWindow(onlineTargets)
                {
                    Owner = Application.Current?.MainWindow,
                };
                searchWindow.ShowDialog();
                onlineFound = searchWindow.FoundRows.Count;
                onlineFailed = searchWindow.Failures.Count;

                if (onlineFound > 0)
                {
                    if (OnlineOfdRowsImported is not null)
                        OnlineOfdRowsImported(searchWindow.FoundRows);
                    else
                        TaxcomReceiptCacheService.Upsert(searchWindow.FoundRows);

                    var mergedRows = ReportImportService.MergeOfdRows(
                        searchWindow.FoundRows.Concat(_sourceOfdRows));
                    SetSourceOfdReport(
                        string.IsNullOrWhiteSpace(_sourceOfdReportPath)
                            ? FileHelper.TaxcomReceiptCachePath
                            : _sourceOfdReportPath,
                        mergedRows);
                    summary = MatchOriginalReceipts(selected);
                }

                foreach (var failure in searchWindow.Failures)
                {
                    foreach (var item in selected.Where(item =>
                                 ParseFiscalSign(item.Record.PrimaryDocument.OriginalFiscalNumber) == failure.FiscalSign &&
                                 item.State.OriginalReceiptLookupState == OriginalReceiptLookupState.NotFound))
                    {
                        item.State.LastMessage = $"ФП {failure.FiscalSign} не найден онлайн в Такскоме: {failure.Message}";
                        item.State.UpdatedAt = DateTime.Now;
                        item.Refresh();
                    }
                }
            }

            SaveStates();
            Status = $"Исходные чеки: найдено {summary.Found}; не найдено {summary.NotFound}; " +
                     $"требуют проверки {summary.RequiresAttention}" +
                     (onlineTargets.Count > 0
                         ? $" · Такском: найдено {onlineFound}, не найдено {onlineFailed}"
                         : string.Empty);
            RefreshCounters();
        }
        catch (Exception ex)
        {
            Status = $"Не удалось найти исходные чеки: {ex.Message}";
        }
    }

    private ReceiptMatchSummary MatchOriginalReceipts(IReadOnlyList<ObsidianCaseItemViewModel> selected)
    {
        var found = 0;
        var notFound = 0;
        var requiresAttention = 0;
        foreach (var item in selected)
        {
            var fiscalSign = ParseFiscalSign(item.Record.PrimaryDocument.OriginalFiscalNumber);
            if (!fiscalSign.HasValue)
            {
                item.State.OriginalReceipt = null;
                item.State.OriginalReceiptLookupState = OriginalReceiptLookupState.MissingFiscalSign;
                item.State.LastMessage = "В 1С не заполнен ФП исходного чека";
                item.State.UpdatedAt = DateTime.Now;
                item.Refresh();
                requiresAttention++;
                continue;
            }

            var candidates = _sourceOfdRows.Where(x => x.FiscalSign == fiscalSign.Value).Take(2).ToList();
            if (candidates.Count != 1)
            {
                item.State.OriginalReceipt = null;
                item.State.OriginalReceiptLookupState = candidates.Count == 0
                    ? OriginalReceiptLookupState.NotFound
                    : OriginalReceiptLookupState.Ambiguous;
                item.State.LastMessage = candidates.Count == 0
                    ? $"ФП {fiscalSign.Value} не найден в загруженном архиве ОФД"
                    : $"ФП {fiscalSign.Value} встречается в архиве ОФД больше одного раза";
                item.State.UpdatedAt = DateTime.Now;
                item.Refresh();
                if (candidates.Count == 0) notFound++;
                else requiresAttention++;
                continue;
            }

            var row = candidates[0];
            item.State.OriginalReceipt = new ObsidianOriginalReceipt
            {
                Source = "ОФД / Такском",
                RegisteredAt = row.RegisteredAt,
                Document = row.Document,
                Operation = row.Operation,
                Amount = row.Amount,
                FiscalSign = row.FiscalSign,
                FiscalDocument = row.FiscalDocument,
                ReceiptUrl = row.ReceiptUrl,
            };
            item.State.OriginalReceiptLookupState = OriginalReceiptLookupState.Found;
            item.Record.PrimaryDocument.OriginalCheckAmount = Math.Abs(row.Amount);
            item.Record.PrimaryDocument.OriginalCheckDate = row.RegisteredAt;
            item.Record.PrimaryDocument.OriginalCheckOperation = row.Operation;
            item.State.LastMessage = CorrectionPlanService.IsCorrectionReceipt(item.State.OriginalReceipt)
                ? "Найден исходный чек коррекции: обработка отложена до этапа 3"
                : row.SourceFile.Equals("Онлайн-поиск Такском", StringComparison.OrdinalIgnoreCase)
                    ? "Исходный чек найден онлайн в Такскоме строго по ФП"
                    : "Исходный чек найден в архиве ОФД строго по ФП";
            item.State.UpdatedAt = DateTime.Now;
            item.Refresh();
            found++;
        }

        return new ReceiptMatchSummary(found, notFound, requiresAttention);
    }

    private static TaxcomReceiptSearchRequest CreateTaxcomRequest(
        ObsidianCaseItemViewModel item,
        long fiscalSign)
    {
        var source = item.Record.PrimaryDocument;
        var anchorDate = NormalizeReceiptDate(source.OneCCheckDate)
                         ?? ParseDocumentDate(source.OrderDate)
                         ?? ParseDocumentDate(source.SourceDocumentDate)
                         ?? DateTime.Today;
        var fromMonth = anchorDate.AddMonths(-1);
        var toMonth = anchorDate.AddMonths(1);
        return new TaxcomReceiptSearchRequest
        {
            FiscalSign = fiscalSign,
            PeriodFrom = new DateTime(fromMonth.Year, fromMonth.Month, 1),
            PeriodTo = new DateTime(toMonth.Year, toMonth.Month, 1).AddMonths(1).AddDays(-1),
            DocumentLabel = $"{item.DocumentType} {item.DocumentNumber}".Trim(),
        };
    }

    private static long? ParseFiscalSign(string value)
    {
        var normalized = value?.Replace(" ", string.Empty).Trim();
        return long.TryParse(normalized, out var result) ? result : null;
    }

    private static DateTime? ParseDocumentDate(string value)
    {
        string[] formats =
        {
            "dd.MM.yyyy HH:mm:ss", "dd.MM.yyyy H:mm:ss", "dd.MM.yyyy",
        };
        return DateTime.TryParseExact(value?.Trim(), formats, CultureInfo.InvariantCulture,
            DateTimeStyles.AllowWhiteSpaces, out var result)
            ? NormalizeReceiptDate(result)
            : null;
    }

    private static DateTime? NormalizeReceiptDate(DateTime? value)
    {
        if (!value.HasValue) return null;
        var date = value.Value;
        return date.Year is >= 2000 && date.Year <= 2100 ? date : null;
    }

    private sealed record ReceiptMatchSummary(int Found, int NotFound, int RequiresAttention);

    private bool TryUseWorkspaceOfdReport()
    {
        var report = OfdReportProvider?.Invoke();
        if (report is not { } value || string.IsNullOrWhiteSpace(value.Path) || value.Rows.Count == 0)
            return false;

        SetSourceOfdReport(value.Path, value.Rows);
        return true;
    }

    private void SetSourceOfdReport(string path, IReadOnlyList<OfdReportRow> rows)
    {
        _sourceOfdReportPath = path;
        _sourceOfdRows = rows.ToList();
        OnPropertyChanged(nameof(SourceOfdReportInfo));
    }

    private void VerifyReport()
    {
        var dialog = new OpenFileDialog
        {
            Title = "Выберите отчёт АТОЛ или ОФД",
            Filter = "Отчёты АТОЛ и ОФД (*.csv;*.xlsx)|*.csv;*.xlsx|CSV (*.csv)|*.csv|Excel (*.xlsx)|*.xlsx",
        };
        if (dialog.ShowDialog() != true) return;

        try
        {
            var selected = SelectedActive().ToList();
            var confirmed = dialog.FileName.EndsWith(".csv", StringComparison.OrdinalIgnoreCase)
                ? ConfirmByAtol(selected, ReportImportService.ReadAtolJournal(dialog.FileName))
                : ConfirmByOfd(selected, ReportImportService.ReadOfdReport(dialog.FileName));
            SaveStates();
            TryCompleteSelected(selected);
            Status = $"По отчёту подтверждено: {confirmed} из {selected.Count}";
        }
        catch (Exception ex)
        {
            Status = $"Ошибка проверки отчёта: {ex.Message}";
        }
    }

    private int ConfirmByAtol(IReadOnlyList<ObsidianCaseItemViewModel> selected, IReadOnlyList<AtolJournalReportRow> rows)
    {
        var confirmed = 0;
        foreach (var item in selected)
        {
            var expected = item.State.ExpectedChecks;
            if (expected.Count == 0)
            {
                item.State.LastMessage = "Нет идентификаторов сформированных чеков";
                item.Refresh();
                continue;
            }

            var allFound = true;
            foreach (var check in expected)
            {
                var row = rows.FirstOrDefault(x => x.ExternalId.Equals(check.ExternalId, StringComparison.OrdinalIgnoreCase));
                if (row is null || row.FiscalSign is null)
                {
                    allFound = false;
                    continue;
                }
                check.FiscalSign = row.FiscalSign;
                check.FiscalDocument = row.FiscalDocument;
            }
            item.State.CheckConfirmed = allFound;
            item.State.LastMessage = allFound ? "Все чеки найдены в отчёте АТОЛ" : "В отчёте найдены не все чеки";
            item.State.UpdatedAt = DateTime.Now;
            if (allFound) confirmed++;
            item.Refresh();
        }
        return confirmed;
    }

    private int ConfirmByOfd(IReadOnlyList<ObsidianCaseItemViewModel> selected, IReadOnlyList<OfdReportRow> rows)
    {
        var confirmed = 0;
        foreach (var item in selected)
        {
            var expected = item.State.ExpectedChecks;
            if (expected.Count == 0)
            {
                item.State.LastMessage = "Нет данных сформированных чеков для проверки ОФД";
                item.Refresh();
                continue;
            }

            var allFound = true;
            foreach (var check in expected)
            {
                OfdReportRow? match = check.FiscalSign is not null
                    ? rows.FirstOrDefault(x => x.FiscalSign == check.FiscalSign)
                    : FindUniqueOfdMatch(rows, check);
                if (match is null)
                {
                    allFound = false;
                    continue;
                }
                check.FiscalSign = match.FiscalSign;
                check.FiscalDocument = match.FiscalDocument;
            }
            item.State.CheckConfirmed = allFound;
            item.State.LastMessage = allFound ? "Все чеки найдены в отчёте ОФД" : "В ОФД нет однозначного совпадения";
            item.State.UpdatedAt = DateTime.Now;
            if (allFound) confirmed++;
            item.Refresh();
        }
        return confirmed;
    }

    private static OfdReportRow? FindUniqueOfdMatch(IEnumerable<OfdReportRow> rows, ObsidianExpectedCheck check)
    {
        var candidates = rows.Where(x =>
                Math.Abs(Math.Abs(x.Amount) - Math.Abs(check.Amount)) < 0.01 &&
                x.RegisteredAt >= check.GeneratedAt.AddMinutes(-5) &&
                OperationMatches(x.Operation, check.Operation))
            .Take(2)
            .ToList();
        return candidates.Count == 1 ? candidates[0] : null;
    }

    private static bool OperationMatches(string reportOperation, string expectedOperation)
    {
        if (string.IsNullOrWhiteSpace(expectedOperation)) return true;
        var value = reportOperation.Trim().ToLowerInvariant();
        if (value.Contains("возврат") && value.Contains("приход")) return expectedOperation == "sell_refund";
        if (value.Contains("коррек") && value.Contains("приход")) return expectedOperation == "sell_correction";
        if (value.Contains("коррек") && value.Contains("расход")) return expectedOperation == "buy_correction";
        if (value == "приход") return expectedOperation == "sell";
        return true;
    }

    private void MarkOneCRecorded()
    {
        var selected = SelectedActive().ToList();
        foreach (var item in selected)
        {
            item.State.OneCRecorded = true;
            item.State.LastMessage = "Запись пояснения в 1С подтверждена вручную";
            item.State.UpdatedAt = DateTime.Now;
            item.Refresh();
        }
        SaveStates();
        TryCompleteSelected(selected);
        Status = $"Запись в 1С подтверждена: {selected.Count}";
    }

    private void SaveProblem()
    {
        var item = SelectedCase;
        if (item is null || !item.CanEdit) return;
        _ignoreWatcherUntil = DateTime.Now.AddSeconds(1);
        if (!ObsidianSyncService.UpdateProblem(FilePath, item.CaseId,
                item.Record.PrimaryDocument.Notes, item.EditableProblem))
        {
            Status = "Не удалось найти строку задачи в Obsidian";
            return;
        }
        item.State.LastMessage = "Пояснение сохранено в Obsidian";
        item.State.UpdatedAt = DateTime.Now;
        SaveStates();
        Sync();
    }

    private void OpenServiceNote()
    {
        var path = SelectedCase?.State.ServiceNotePath;
        if (string.IsNullOrWhiteSpace(path) || !File.Exists(path)) return;
        Process.Start(new ProcessStartInfo(path) { UseShellExecute = true });
    }

    private void SelectAllVisible()
    {
        var visible = CasesView.Cast<ObsidianCaseItemViewModel>().Where(x => !x.IsCompleted).ToList();
        var select = visible.Any(x => !x.IsSelected);
        foreach (var item in visible) item.IsSelected = select;
        RefreshCounters();
    }

    private void SelectCase(object? parameter)
    {
        if (parameter is ObsidianCaseItemViewModel item)
            SelectedCase = item;
    }

    private bool ValidateMemo(ObsidianCaseState state, string documentNumber, double amount)
    {
        var result = ServiceNoteValidationService.Validate(documentNumber, amount, state.ServiceNotePath);
        state.ServiceNoteVerified = result.IsValid;
        state.ServiceNotePath = result.Path;
        state.LastMessage = result.Message;
        state.UpdatedAt = DateTime.Now;
        return result.IsValid;
    }

    private void TryCompleteSelected(IReadOnlyList<ObsidianCaseItemViewModel> selected)
    {
        var completed = selected.Where(x =>
            x.State.CheckConfirmed && x.State.ServiceNoteVerified && x.State.OneCRecorded).ToList();
        if (completed.Count == 0) return;

        _ignoreWatcherUntil = DateTime.Now.AddSeconds(1);
        foreach (var item in completed)
        {
            ObsidianSyncService.MarkCompleted(FilePath, item.CaseId);
            item.State.LastMessage = "Случай закрыт и отмечен в Obsidian";
            item.State.UpdatedAt = DateTime.Now;
        }
        SaveStates();
        Sync();
    }

    private IEnumerable<ObsidianCaseItemViewModel> SelectedActive() =>
        Cases.Where(x => x.IsSelected && !x.IsCompleted);

    private bool HasSelectedActive() => SelectedActive().Any();
    private bool HasSelectedReady() => SelectedActive().Any(x => x.PlanReady);

    private static OrderEntry CloneSnapshot(OrderEntry source) => new()
    {
        ObsidianCaseId = source.ObsidianCaseId,
        OrderNum = source.OrderNum,
        OrderDate = source.OrderDate,
        Amount = source.Amount,
        CustomerName = source.CustomerName,
        Items = source.Items.Select(x => new OrderItem
        {
            Name = x.Name,
            Quantity = x.Quantity,
            Sum = x.Sum,
            VatType = x.VatType,
        }).ToList(),
        OriginalItems = source.OriginalItems.Select(x => new OrderItem
        {
            Name = x.Name,
            Quantity = x.Quantity,
            Sum = x.Sum,
            VatType = x.VatType,
        }).ToList(),
        AgentInfo = source.AgentInfo is null ? null : new ServiceProvider
        {
            Service = source.AgentInfo.Service,
            City = source.AgentInfo.City,
            Name = source.AgentInfo.Name,
            Inn = source.AgentInfo.Inn,
            Phone = source.AgentInfo.Phone,
            VatType = source.AgentInfo.VatType,
        },
        CorrectionDate = source.CorrectionDate,
        CorrectionNumber = source.CorrectionNumber,
        IsService = source.IsService,
        IsOwnService = source.IsOwnService,
        ServiceType = source.ServiceType,
        City = source.City,
        Kind = source.Kind,
        DocumentType = source.DocumentType,
        CorrectionScenario = source.CorrectionScenario,
        OriginalFiscalNumber = source.OriginalFiscalNumber,
        OriginalCheckAmount = source.OriginalCheckAmount,
        CorrectAmount = source.CorrectAmount,
        Notes = source.Notes,
        OneCComment = source.OneCComment,
        OneCCheckNumber = source.OneCCheckNumber,
        OneCCheckDate = source.OneCCheckDate,
        OriginalPaymentWasCash = source.OriginalPaymentWasCash,
        CorrectPaymentIsCash = source.CorrectPaymentIsCash,
        PlannedReverseOperation = source.PlannedReverseOperation,
        PlannedCorrectOperation = source.PlannedCorrectOperation,
        OriginalVatType = source.OriginalVatType,
        CorrectVatType = source.CorrectVatType,
        PlannedVatType = source.PlannedVatType,
    };

    private static void ApplySnapshot(OrderEntry snapshot, OrderEntry target)
    {
        if (ParseDocumentDate(snapshot.OrderDate).HasValue)
            target.OrderDate = snapshot.OrderDate;
        target.Amount = snapshot.Amount;
        target.CustomerName = snapshot.CustomerName;
        target.Items = snapshot.Items.Select(x => new OrderItem
        {
            Name = x.Name,
            Quantity = x.Quantity,
            Sum = x.Sum,
            VatType = x.VatType,
        }).ToList();
        target.OriginalItems = snapshot.OriginalItems.Select(x => new OrderItem
        {
            Name = x.Name,
            Quantity = x.Quantity,
            Sum = x.Sum,
            VatType = x.VatType,
        }).ToList();
        target.AgentInfo = snapshot.AgentInfo;
        target.CorrectionDate = snapshot.CorrectionDate;
        target.CorrectionNumber = snapshot.CorrectionNumber;
        target.IsService = snapshot.IsService;
        target.IsOwnService = snapshot.IsOwnService;
        target.ServiceType = snapshot.ServiceType;
        target.City = snapshot.City;
        target.OriginalFiscalNumber = snapshot.OriginalFiscalNumber;
        target.OriginalCheckAmount = snapshot.OriginalCheckAmount;
        target.CorrectAmount = snapshot.CorrectAmount;
        target.OneCComment = snapshot.OneCComment;
        target.OneCCheckNumber = snapshot.OneCCheckNumber;
        target.OneCCheckDate = NormalizeReceiptDate(snapshot.OneCCheckDate);
        target.OriginalPaymentWasCash = snapshot.OriginalPaymentWasCash;
        target.CorrectPaymentIsCash = snapshot.CorrectPaymentIsCash;
        target.PlannedReverseOperation = snapshot.PlannedReverseOperation;
        target.PlannedCorrectOperation = snapshot.PlannedCorrectOperation;
        target.OriginalVatType = snapshot.OriginalVatType;
        target.CorrectVatType = snapshot.CorrectVatType;
        target.PlannedVatType = snapshot.PlannedVatType;
    }

    private bool FilterCase(object value)
    {
        if (value is not ObsidianCaseItemViewModel item) return false;
        if (!ShowClosed && item.IsCompleted) return false;
        if (!string.Equals(SelectedPeriod, "Все периоды", StringComparison.OrdinalIgnoreCase) &&
            !string.Equals(item.PeriodLabel, SelectedPeriod, StringComparison.OrdinalIgnoreCase))
            return false;
        if (string.IsNullOrWhiteSpace(SearchText)) return true;
        var query = SearchText.Trim();
        return item.DocumentNumber.Contains(query, StringComparison.OrdinalIgnoreCase) ||
               item.DocumentType.Contains(query, StringComparison.OrdinalIgnoreCase) ||
               item.Department.Contains(query, StringComparison.OrdinalIgnoreCase) ||
               item.EditableProblem.Contains(query, StringComparison.OrdinalIgnoreCase);
    }

    private void RefreshPeriodOptions()
    {
        var selected = SelectedPeriod;
        PeriodOptions.Clear();
        PeriodOptions.Add("Все периоды");
        foreach (var period in Cases.Select(x => x.PeriodLabel)
                     .Distinct(StringComparer.OrdinalIgnoreCase))
            PeriodOptions.Add(period);

        var resolved = PeriodOptions.Contains(selected) ? selected : "Все периоды";
        _selectedPeriod = resolved;
        _selectedPeriodIndex = Math.Max(0, PeriodOptions.IndexOf(resolved));
        OnPropertyChanged(nameof(SelectedPeriod));
        OnPropertyChanged(nameof(SelectedPeriodIndex));
        CasesView.Refresh();
        Application.Current?.Dispatcher.BeginInvoke(
            new Action(() =>
            {
                OnPropertyChanged(nameof(SelectedPeriodIndex));
            }),
            System.Windows.Threading.DispatcherPriority.Loaded);
    }

    private ObsidianCaseState GetState(string caseId)
    {
        if (string.IsNullOrWhiteSpace(caseId)) return new ObsidianCaseState();
        if (_states.TryGetValue(caseId, out var state)) return state;
        state = new ObsidianCaseState { CaseId = caseId };
        _states[caseId] = state;
        return state;
    }

    private void SaveStates() => ObsidianCaseStateStore.Save(_states.Values);

    private void RefreshCounters()
    {
        OnPropertyChanged(nameof(ActiveCount));
        OnPropertyChanged(nameof(ClosedCount));
        OnPropertyChanged(nameof(SelectedCount));
        CommandManager.InvalidateRequerySuggested();
    }

    private void ConfigureWatcher()
    {
        _watcher?.Dispose();
        _watcher = null;
        if (string.IsNullOrWhiteSpace(FilePath) || !File.Exists(FilePath)) return;

        _watcher = new FileSystemWatcher(Path.GetDirectoryName(FilePath)!, Path.GetFileName(FilePath))
        {
            NotifyFilter = NotifyFilters.LastWrite | NotifyFilters.Size | NotifyFilters.FileName,
            EnableRaisingEvents = true,
        };
        _watcher.Changed += OnFileChanged;
        _watcher.Created += OnFileChanged;
        _watcher.Renamed += OnFileChanged;
    }

    private void OnFileChanged(object sender, FileSystemEventArgs e)
    {
        if (DateTime.Now < _ignoreWatcherUntil) return;
        Application.Current?.Dispatcher.BeginInvoke(() =>
        {
            if (DateTime.Now < _ignoreWatcherUntil) return;
            Sync();
        });
    }

    public void Dispose() => _watcher?.Dispose();
}
