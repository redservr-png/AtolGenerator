using System.Collections.ObjectModel;
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
    public string VatLabel => VatType switch
    {
        "vat122" => "НДС 22/122",
        "vat22" => "НДС 22%",
        "vat105" => "НДС 5/105",
        "vat5" => "НДС 5%",
        "none" => "Без НДС",
        _ => VatType,
    };
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
    public string StateLabel => IsGenerated ? "XML сформирован" : IsReady ? "Готов" : "Нужна проверка";
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
        if (hasReverse)
        {
            Steps.Add(BuildStep(
                1,
                "Отмена исходного чека",
                Entry.PlannedReverseOperation,
                Entry.OriginalCheckAmount ?? Entry.Amount,
                ResolvePayment(Entry.OriginalPaymentWasCash),
                usesTag1192: true,
                Entry.OriginalItems.Count > 0 ? Entry.OriginalItems : Entry.Items));
        }

        if (!string.IsNullOrWhiteSpace(Entry.PlannedCorrectOperation))
        {
            Steps.Add(BuildStep(
                Steps.Count + 1,
                Entry.PlannedCorrectOperation.EndsWith("_correction", StringComparison.OrdinalIgnoreCase)
                    ? "Чек коррекции"
                    : "Правильный чек",
                Entry.PlannedCorrectOperation,
                Entry.CorrectAmount ?? Entry.Amount,
                ResolvePayment(Entry.CorrectPaymentIsCash),
                usesTag1192: hasReverse &&
                             !Entry.PlannedCorrectOperation.EndsWith("_correction", StringComparison.OrdinalIgnoreCase),
                Entry.Items));
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
            VatType = Entry.PlannedVatType,
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

        var hasCorrectReceipt = !string.IsNullOrWhiteSpace(Entry.PlannedCorrectOperation) &&
                                !Entry.PlannedCorrectOperation.EndsWith(
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
            ? "Официальный комплект ФФД 1.05 готов к формированию XML."
            : "Чек готов к формированию XML.";
        return true;
    }

    private string ResolvePayment(bool? isCash) => isCash.HasValue
        ? isCash.Value ? "cash" : "card"
        : Entry.DocumentType is SourceDocumentType.CashPayment or SourceDocumentType.CashExpense
            ? "cash"
            : "card";

    private static bool ItemsMatch(IReadOnlyCollection<OrderItem> items, double amount) =>
        items.Count > 0 && Math.Abs(items.Sum(x => x.Sum) - amount) <= 0.01;
}

public sealed class CorrectionWorkViewModel : BaseViewModel
{
    private CorrectionWorkItemViewModel? _selectedItem;
    private CashierInfo _selectedCashier = AppConstants.DefaultCashier;
    private bool _mergeXml = true;
    private bool _isBusy;
    private string _statusText = "Добавьте подготовленные случаи из раздела «Исправление чеков».";
    private string _lastXmlPath = string.Empty;

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

    public ICommand GenerateCommand { get; }
    public ICommand EditCommand { get; }
    public ICommand RemoveCommand { get; }
    public ICommand ClearCommand { get; }
    public ICommand BackCommand { get; }
    public ICommand OpenFolderCommand { get; }
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
        StatusText = $"Формирование XML: {selected.Count} исправлений, {selected.Sum(x => x.Steps.Count)} чеков...";
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
            StatusText = results.Count == 0
                ? "XML не сформирован: в выбранных случаях нет готовых чеков."
                : $"Готово: {results.Count} чеков и {selected.Count} служебных записок. Файлы сохранены в atol_output.";
            if (results.Count > 0)
                Generated?.Invoke(results);
        }
        catch (Exception ex)
        {
            StatusText = $"Ошибка формирования XML: {ex.Message}";
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
        StatusText = "Список подготовки очищен.";
        RefreshCounters();
    }

    private void SetAllSelected(bool selected)
    {
        foreach (var item in Items)
            item.IsSelected = selected;
        RefreshCounters();
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
