using System.Collections.ObjectModel;
using System.Windows;
using System.Windows.Input;
using AtolGenerator.Constants;
using AtolGenerator.Helpers;
using AtolGenerator.Models;
using AtolGenerator.Services;

namespace AtolGenerator.ViewModels;

public sealed class CorrectionWorkStepViewModel
{
    public int Sequence { get; init; }
    public string Title { get; init; } = string.Empty;
    public string Operation { get; init; } = string.Empty;
    public string OperationLabel => Operation switch
    {
        "sell" => "Приход",
        "sell_refund" => "Возврат прихода",
        "buy" => "Расход",
        "buy_refund" => "Возврат расхода",
        "sell_correction" => "Коррекция прихода",
        "buy_correction" => "Коррекция возврата",
        _ => Operation,
    };
    public double Amount { get; init; }
    public string AmountText => $"{Amount:N2} ₽";
    public string PaymentType { get; init; } = string.Empty;
    public string PaymentLabel => PaymentType == "cash" ? "Наличные" : "Безналичные";
    public string VatType { get; init; } = string.Empty;
    public string VatLabel => VatRateCatalog.LabelFor(VatType);
    public bool UsesTag1192 { get; init; }
    public string Tag1192Text { get; init; } = string.Empty;
    public string ItemsText { get; init; } = string.Empty;
    public bool IsCorrectionReceipt => Operation.EndsWith("_correction", StringComparison.OrdinalIgnoreCase);
}

public sealed class CorrectionWorkItemViewModel : BaseViewModel
{
    private readonly Action _selectionChanged;
    private bool _isSelected = true;
    private bool _isGenerated;

    public CorrectionWorkItemViewModel(OrderEntry entry, Action selectionChanged)
    {
        Entry = entry;
        _selectionChanged = selectionChanged;
        Refresh();
    }

    public OrderEntry Entry { get; }
    public ObservableCollection<CorrectionWorkStepViewModel> Steps { get; } = new();

    public bool IsSelected
    {
        get => _isSelected;
        set
        {
            if (!Set(ref _isSelected, value)) return;
            _selectionChanged();
        }
    }

    public bool IsGenerated
    {
        get => _isGenerated;
        set
        {
            if (!Set(ref _isGenerated, value)) return;
            OnPropertyChanged(nameof(StateLabel));
        }
    }

    public string DocumentNumber => Entry.OrderNum;
    public string DocumentDate => Entry.OrderDate;
    public string Department => Entry.City;
    public string ScenarioLabel => Entry.CorrectionScenario.ToDisplayString();
    public string DocumentTypeLabel => Entry.DocumentType switch
    {
        SourceDocumentType.Realization => "Реализация",
        SourceDocumentType.CardPayment => "Оплата картой",
        SourceDocumentType.CashPayment => "ПКО / наличные",
        SourceDocumentType.CashExpense => "РКО / возврат",
        SourceDocumentType.BuyerOrder => "Заказ покупателя",
        _ => "Документ 1С",
    };
    public string OriginalFiscalNumber => string.IsNullOrWhiteSpace(Entry.OriginalFiscalNumber)
        ? "Не заполнен"
        : Entry.OriginalFiscalNumber;
    public string Notes => Entry.Notes;
    public string StateLabel => IsGenerated ? "Сформирован" : IsReady ? "Готов" : "Нужна проверка";
    public bool IsReady { get; private set; }
    public string ReadinessMessage { get; private set; } = string.Empty;
    public string CheckCountText => Steps.Count switch
    {
        1 => "1 чек",
        _ => $"{Steps.Count} чека",
    };

    public void Refresh()
    {
        Steps.Clear();
        var hasReverse = !string.IsNullOrWhiteSpace(Entry.PlannedReverseOperation);
        var correctOperation = ResolveCorrectOperation(hasReverse);
        if (hasReverse)
        {
            Steps.Add(BuildStep(
                1,
                "Отмена исходного чека",
                Entry.PlannedReverseOperation,
                Entry.OriginalCheckAmount ?? Entry.Amount,
                ResolvePayment(Entry.OriginalPaymentWasCash),
                usesTag1192: true,
                vatType: VatRateCatalog.Normalize(Entry.OriginalVatType, Entry.PlannedVatType),
                items: Entry.OriginalItems.Count > 0 ? Entry.OriginalItems : Entry.Items));
        }

        if (!string.IsNullOrWhiteSpace(correctOperation))
        {
            Steps.Add(BuildStep(
                Steps.Count + 1,
                correctOperation.EndsWith("_correction", StringComparison.OrdinalIgnoreCase)
                    ? "Чек коррекции"
                    : "Правильный чек",
                correctOperation,
                Entry.CorrectAmount ?? Entry.Amount,
                ResolvePayment(Entry.CorrectPaymentIsCash),
                usesTag1192: false,
                vatType: VatRateCatalog.Normalize(Entry.CorrectVatType, Entry.PlannedVatType),
                items: Entry.Items));
        }

        IsReady = Validate(out var message);
        ReadinessMessage = message;
        OnPropertyChanged(nameof(DocumentNumber));
        OnPropertyChanged(nameof(DocumentDate));
        OnPropertyChanged(nameof(Department));
        OnPropertyChanged(nameof(ScenarioLabel));
        OnPropertyChanged(nameof(DocumentTypeLabel));
        OnPropertyChanged(nameof(OriginalFiscalNumber));
        OnPropertyChanged(nameof(Notes));
        OnPropertyChanged(nameof(StateLabel));
        OnPropertyChanged(nameof(IsReady));
        OnPropertyChanged(nameof(ReadinessMessage));
        OnPropertyChanged(nameof(CheckCountText));
    }

    private CorrectionWorkStepViewModel BuildStep(
        int sequence,
        string title,
        string operation,
        double amount,
        string payment,
        bool usesTag1192,
        string vatType,
        IReadOnlyCollection<OrderItem> items)
    {
        var isCorrection = operation.EndsWith("_correction", StringComparison.OrdinalIgnoreCase);
        var itemsTotal = items.Sum(x => x.Sum);
        return new CorrectionWorkStepViewModel
        {
            Sequence = sequence,
            Title = title,
            Operation = operation,
            Amount = amount,
            PaymentType = payment,
            VatType = VatRateCatalog.Normalize(vatType, "none"),
            UsesTag1192 = usesTag1192,
            Tag1192Text = usesTag1192
                ? $"Тег 1192: {OriginalFiscalNumber}"
                : isCorrection ? "Тег 1192 не допускается XSD ФФД 1.05" : "Без тега 1192",
            ItemsText = isCorrection
                ? "Без табличной части"
                : $"Позиций: {items.Count} · итог {itemsTotal:N2} ₽",
        };
    }

    private bool Validate(out string message)
    {
        if (Entry.CorrectionScenario is CorrectionScenario.Unknown or CorrectionScenario.RealRefund)
        {
            message = Entry.CorrectionScenario == CorrectionScenario.RealRefund
                ? "Реальный возврат перенесите в раздел «Возвраты по заказам»."
                : "Выберите сценарий исправления.";
            return false;
        }

        if (Steps.Count == 0)
        {
            message = "План чеков не рассчитан.";
            return false;
        }

        var hasReverse = !string.IsNullOrWhiteSpace(Entry.PlannedReverseOperation);
        if (hasReverse && string.IsNullOrWhiteSpace(Entry.OriginalFiscalNumber))
        {
            message = "Не заполнен ФП исходного чека для тега 1192.";
            return false;
        }

        if (hasReverse && (Entry.OriginalCheckAmount ?? 0) <= 0)
        {
            message = "Не заполнена сумма исходного чека.";
            return false;
        }

        if (Entry.DocumentType == SourceDocumentType.Realization && hasReverse)
        {
            var originalItems = Entry.OriginalItems.Count > 0 ? Entry.OriginalItems : Entry.Items;
            if (!ItemsMatch(originalItems, Entry.OriginalCheckAmount ?? 0))
            {
                message = "Позиции исходного чека не сходятся с его суммой.";
                return false;
            }
        }

        var correctOperation = ResolveCorrectOperation(hasReverse);
        var hasCorrectReceipt = !string.IsNullOrWhiteSpace(correctOperation) &&
                                !correctOperation.EndsWith(
                                    "_correction", StringComparison.OrdinalIgnoreCase);
        if (hasCorrectReceipt && (Entry.CorrectAmount ?? Entry.Amount) <= 0)
        {
            message = "Не заполнена сумма правильного чека.";
            return false;
        }

        if (Entry.DocumentType == SourceDocumentType.Realization && hasCorrectReceipt &&
            !ItemsMatch(Entry.Items, Entry.CorrectAmount ?? Entry.Amount))
        {
            message = "Позиции правильного чека не сходятся с исправленной суммой.";
            return false;
        }

            message = hasReverse
                ? "Готов: возврат через API, коррекция через XML."
                : "Готов к пробитию чека коррекции через XML.";
        return true;
    }

    private string ResolvePayment(bool? isCash) => isCash.HasValue
        ? isCash.Value ? "cash" : "card"
        : Entry.DocumentType is SourceDocumentType.CashPayment or SourceDocumentType.CashExpense
            ? "cash"
            : "card";

    private string ResolveCorrectOperation(bool hasReverse)
    {
        if (string.IsNullOrWhiteSpace(Entry.PlannedCorrectOperation))
            return string.Empty;

        return hasReverse && !Entry.PlannedCorrectOperation.EndsWith("_correction", StringComparison.OrdinalIgnoreCase)
            ? CorrectionPlanService.ToCorrectionOperation(Entry.PlannedCorrectOperation)
            : Entry.PlannedCorrectOperation;
    }

    private static bool ItemsMatch(IReadOnlyCollection<OrderItem> items, double amount) =>
        items.Count > 0 && Math.Abs(items.Sum(x => x.Sum) - amount) <= 0.01;
}

public sealed class CorrectionWorkViewModel : BaseViewModel
{
    private CorrectionWorkItemViewModel? _selectedItem;
    private CashierInfo _selectedCashier = AppConstants.DefaultCashier;
    private bool _mergeXml;
    private bool _isBusy;
    private string _statusText = "Добавьте подготовленные случаи из раздела «Исправление чеков».";
    private string _lastXmlPath = string.Empty;
    private CorrectionPunchPlan? _lastPunchPlan;
    private bool _allowCorrectionXml;

    public CorrectionWorkViewModel()
    {
        GenerateCommand = new AsyncRelayCommand(GenerateAsync);
        EditCommand = new RelayCommand(item =>
        {
            if (item is CorrectionWorkItemViewModel vm) EditRequested?.Invoke(vm);
        });
        RemoveCommand = new RelayCommand(item => Remove(item as CorrectionWorkItemViewModel));
        ClearCommand = new RelayCommand(Clear);
        BackCommand = new RelayCommand(() => BackRequested?.Invoke());
        OpenFolderCommand = new RelayCommand(() => FileHelper.OpenFolder(FileHelper.OutputDir));
        OpenXmlUploadCommand = new RelayCommand(OpenLastXmlUpload, () => HasCorrectionXml);
        SelectAllCommand = new RelayCommand(() => SetAllSelected(true));
        DeselectAllCommand = new RelayCommand(() => SetAllSelected(false));
    }

    public event Action<CorrectionWorkItemViewModel>? EditRequested;
    public event Action? BackRequested;
    public event Action<IReadOnlyList<GenerationResult>>? Generated;

    public ObservableCollection<CorrectionWorkItemViewModel> Items { get; } = new();
    public ObservableCollection<CashierInfo> AvailableCashiers { get; } = new();

    public CorrectionWorkItemViewModel? SelectedItem
    {
        get => _selectedItem;
        set => Set(ref _selectedItem, value);
    }

    public CashierInfo SelectedCashier
    {
        get => _selectedCashier;
        set => Set(ref _selectedCashier, value ?? AppConstants.DefaultCashier);
    }

    public bool MergeXml
    {
        get => _mergeXml;
        set => Set(ref _mergeXml, value);
    }

    public bool IsBusy
    {
        get => _isBusy;
        private set
        {
            if (!Set(ref _isBusy, value)) return;
            OnPropertyChanged(nameof(CanGenerate));
        }
    }

    public string StatusText
    {
        get => _statusText;
        private set => Set(ref _statusText, value);
    }

    public string LastXmlPath
    {
        get => _lastXmlPath;
        private set
        {
            if (!Set(ref _lastXmlPath, value)) return;
            OnPropertyChanged(nameof(HasGeneratedFiles));
        }
    }

    public int ItemCount => Items.Count;
    public int SelectedCount => Items.Count(x => x.IsSelected);
    public int SelectedCheckCount => Items.Where(x => x.IsSelected).Sum(x => x.Steps.Count);
    public int ReadyCount => Items.Count(x => x.IsReady);
    public bool CanGenerate => !IsBusy && SelectedCount > 0 &&
                               Items.Where(x => x.IsSelected).All(x => x.IsReady);
    public bool HasGeneratedFiles => !string.IsNullOrWhiteSpace(LastXmlPath);
    public bool HasCorrectionXml => _lastPunchPlan?.HasCorrectionXml == true;

    public ICommand GenerateCommand { get; }
    public ICommand EditCommand { get; }
    public ICommand RemoveCommand { get; }
    public ICommand ClearCommand { get; }
    public ICommand BackCommand { get; }
    public ICommand OpenFolderCommand { get; }
    public ICommand OpenXmlUploadCommand { get; }
    public ICommand SelectAllCommand { get; }
    public ICommand DeselectAllCommand { get; }

    public void SyncCashiers(IEnumerable<CashierInfo> cashiers, CashierInfo selected)
    {
        var selectedShortName = SelectedCashier?.ShortName;
        AvailableCashiers.Clear();
        foreach (var cashier in cashiers)
            AvailableCashiers.Add(cashier);

        SelectedCashier = AvailableCashiers.FirstOrDefault(x => string.Equals(
                              x.ShortName, selectedShortName, StringComparison.OrdinalIgnoreCase))
                          ?? AvailableCashiers.FirstOrDefault(x => string.Equals(
                              x.ShortName, selected.ShortName, StringComparison.OrdinalIgnoreCase))
                          ?? AvailableCashiers.FirstOrDefault()
                          ?? AppConstants.DefaultCashier;
    }

    public (int Added, int Updated) AddOrUpdate(IEnumerable<OrderEntry> entries)
    {
        var added = 0;
        var updated = 0;
        foreach (var entry in entries)
        {
            var existing = string.IsNullOrWhiteSpace(entry.ObsidianCaseId)
                ? null
                : Items.FirstOrDefault(x => x.Entry.ObsidianCaseId.Equals(
                    entry.ObsidianCaseId, StringComparison.OrdinalIgnoreCase));
            var wrapper = new CorrectionWorkItemViewModel(entry, RefreshCounters);
            if (existing is null)
            {
                Items.Add(wrapper);
                added++;
            }
            else
            {
                var index = Items.IndexOf(existing);
                Items[index] = wrapper;
                updated++;
            }

            SelectedItem = wrapper;
        }

        StatusText = updated > 0
            ? $"Добавлено: {added}; обновлено: {updated}. Проверьте стороны чеков перед генерацией."
            : $"Добавлено исправлений: {added}. Проверьте стороны чеков перед генерацией.";
        RefreshCounters();
        return (added, updated);
    }

    public void RefreshItem(CorrectionWorkItemViewModel item)
    {
        item.Refresh();
        SelectedItem = item;
        RefreshCounters();
    }

    private async Task GenerateAsync()
    {
        var selected = Items.Where(x => x.IsSelected).ToList();
        if (selected.Count == 0)
        {
            StatusText = "Выберите хотя бы одно исправление.";
            return;
        }

        var invalid = selected.Where(x => !x.IsReady).ToList();
        if (invalid.Count > 0)
        {
            StatusText = $"Не готовы: {string.Join(", ", invalid.Select(x => x.DocumentNumber))}. Откройте редактор.";
            return;
        }

        IsBusy = true;
        StatusText = $"Формирование чеков: {selected.Count} исправлений, {selected.Sum(x => x.Steps.Count)} чеков...";
        try
        {
            var parameters = new GenerationParams
            {
                Tab = "payment",
                CheckType = "sell",
                PaymentType = "card",
                MergeXml = MergeXml,
                Orders = selected.Select(x => x.Entry).ToList(),
                OutputDir = FileHelper.OutputDir,
                Cashier = SelectedCashier,
            };
            var results = await Task.Run(() => CheckGeneratorService.Generate(parameters));
            foreach (var item in selected)
                item.IsGenerated = results.Any(x => x.ObsidianCaseId.Equals(
                    item.Entry.ObsidianCaseId, StringComparison.OrdinalIgnoreCase));

            LastXmlPath = results.FirstOrDefault()?.XmlPath ?? string.Empty;
            _lastPunchPlan = CorrectionPunchPlanner.FromResults(results);
            OnPropertyChanged(nameof(HasCorrectionXml));
            CommandManager.InvalidateRequerySuggested();
            var xmlFiles = results
                .Select(x => x.XmlPath)
                .Where(path => !string.IsNullOrWhiteSpace(path))
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .ToList();
            StatusText = results.Count == 0
                ? "Чеки не сформированы: в выбранных случаях нет готовых документов."
                : $"Сформировано {results.Count} чеков, XML-файлов: {xmlFiles.Count}. Возвраты — API, коррекции — отдельный XML.";
            if (results.Count > 0)
            {
                Generated?.Invoke(results);
                var outcome = await CorrectionPunchCoordinator.AfterGenerateAsync(
                    results,
                    AtolCredentials.Load(),
                    CorrectionPunchCoordinator.FindOwner(),
                    text => StatusText = text);
                _allowCorrectionXml = outcome.AllowCorrectionXml;
                StatusText = outcome.StatusText;
            }
        }
        catch (Exception ex)
        {
            StatusText = $"Ошибка формирования: {ex.Message}";
        }
        finally
        {
            IsBusy = false;
        }
    }

    private void Remove(CorrectionWorkItemViewModel? item)
    {
        if (item is null) return;
        var index = Items.IndexOf(item);
        Items.Remove(item);
        SelectedItem = Items.Count == 0
            ? null
            : Items[Math.Min(Math.Max(index, 0), Items.Count - 1)];
        StatusText = $"Удалено из подготовки: {item.DocumentNumber}.";
        RefreshCounters();
    }

    private void Clear()
    {
        Items.Clear();
        SelectedItem = null;
        LastXmlPath = string.Empty;
        _lastPunchPlan = null;
        _allowCorrectionXml = false;
        OnPropertyChanged(nameof(HasCorrectionXml));
        CommandManager.InvalidateRequerySuggested();
        StatusText = "Список подготовки очищен.";
        RefreshCounters();
    }

    private void SetAllSelected(bool selected)
    {
        foreach (var item in Items)
            item.IsSelected = selected;
        RefreshCounters();
    }

    private void OpenLastXmlUpload()
    {
        if (_lastPunchPlan is not { HasCorrectionXml: true })
        {
            StatusText = "Сначала сформируйте чеки коррекции.";
            return;
        }

        if (!_allowCorrectionXml && _lastPunchPlan.HasApiReceipts)
        {
            var go = MessageBox.Show(
                CorrectionPunchCoordinator.FindOwner(),
                "Возвраты этой пачки не подтверждены через API.\n" +
                "Загрузка XML коррекций оставит в журнале коррекцию без отмены исходного чека.\n\n" +
                "Открыть кабинет всё равно?",
                "XML коррекций",
                MessageBoxButton.YesNo,
                MessageBoxImage.Warning,
                MessageBoxResult.No) == MessageBoxResult.Yes;
            if (!go) return;
        }

        CorrectionPunchCoordinator.OpenXmlUploadWindow(
            _lastPunchPlan, CorrectionPunchCoordinator.FindOwner());
    }

    private void RefreshCounters()
    {
        OnPropertyChanged(nameof(ItemCount));
        OnPropertyChanged(nameof(SelectedCount));
        OnPropertyChanged(nameof(SelectedCheckCount));
        OnPropertyChanged(nameof(ReadyCount));
        OnPropertyChanged(nameof(CanGenerate));
    }
}
