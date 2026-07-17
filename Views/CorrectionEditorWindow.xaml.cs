using System.Windows;
using System.Windows.Controls;
using System.Windows.Media;
using AtolGenerator.Models;
using AtolGenerator.Services;

namespace AtolGenerator.Views;

/// <summary>
/// Редактор исправительного комплекта. Для ФФД 1.05 исходная и правильная
/// стороны обычных чеков редактируются раздельно.
/// </summary>
public partial class CorrectionEditorWindow : Window
{
    private bool _initializing;
    private bool _refreshingLayout;

    public OrderEntry Entry { get; private set; }
    public IReadOnlyList<VatRateOption> AvailableVatRates => VatRateCatalog.All;

    public CorrectionScenario[] AllScenarios { get; } = Enum
        .GetValues<CorrectionScenario>()
        .Where(x => x != CorrectionScenario.RealRefund)
        .ToArray();

    public SourceDocumentType[] AllDocumentTypes { get; } =
        Enum.GetValues<SourceDocumentType>();

    public CorrectionEditorWindow(OrderEntry entry)
    {
        Entry = entry ?? throw new ArgumentNullException(nameof(entry));
        _initializing = true;
        InitializeComponent();
        DataContext = Entry;
        RebuildOfficialPlan();
        EnsureVatSelections();
        EnsureReceiptItems();
        RefreshLayoutMode();
        _initializing = false;
        RefreshTotals();
    }

    private void Save_Click(object sender, RoutedEventArgs e)
    {
        RebuildOfficialPlan();
        EnsureVatSelections();
        EnsureReceiptItems();
        if (!ValidateWorkflow())
            return;

        Entry.Kind = !string.IsNullOrWhiteSpace(Entry.PlannedReverseOperation) &&
                     !string.IsNullOrWhiteSpace(Entry.PlannedCorrectOperation)
            ? OrderKind.RefundCorrectionPair
            : !string.IsNullOrWhiteSpace(Entry.PlannedCorrectOperation) &&
              Entry.PlannedCorrectOperation.EndsWith("_correction", StringComparison.OrdinalIgnoreCase)
                ? OrderKind.SingleCorrection
                : OrderKind.SingleRefund;

        DialogResult = true;
        Close();
    }

    private void Cancel_Click(object sender, RoutedEventArgs e)
    {
        DialogResult = false;
        Close();
    }

    private void LayoutMode_SelectionChanged(object sender, SelectionChangedEventArgs e)
    {
        if (_initializing || _refreshingLayout)
            return;

        RebuildOfficialPlan();
        EnsureVatSelections();
        EnsureReceiptItems();
        RefreshLayoutMode();
    }

    private void AddOriginalItem_Click(object sender, RoutedEventArgs e)
    {
        AddItem(Entry.OriginalItems, Entry.OriginalCheckAmount ?? Entry.Amount, Entry.OriginalVatType);
        OriginalItemsGrid.Items.Refresh();
        RefreshTotals();
    }

    private void DeleteOriginalItem_Click(object sender, RoutedEventArgs e)
    {
        if (OriginalItemsGrid.SelectedItem is not OrderItem item)
            return;

        Entry.OriginalItems.Remove(item);
        OriginalItemsGrid.Items.Refresh();
        RefreshTotals();
    }

    private void AddCorrectItem_Click(object sender, RoutedEventArgs e)
    {
        AddItem(Entry.Items, Entry.CorrectAmount ?? Entry.Amount, Entry.CorrectVatType);
        CorrectItemsGrid.Items.Refresh();
        RefreshTotals();
    }

    private void DeleteCorrectItem_Click(object sender, RoutedEventArgs e)
    {
        if (CorrectItemsGrid.SelectedItem is not OrderItem item)
            return;

        Entry.Items.Remove(item);
        CorrectItemsGrid.Items.Refresh();
        RefreshTotals();
    }

    private void ItemsGrid_CurrentCellChanged(object sender, EventArgs e)
    {
        if (_initializing || _refreshingLayout) return;
        EnsureItemVatSelections();
        RefreshTotals();
    }

    private void ItemVat_SelectionChanged(object sender, SelectionChangedEventArgs e)
    {
        if (_initializing || _refreshingLayout)
            return;

        RefreshTotals();
    }

    private void OriginalVat_SelectionChanged(object sender, SelectionChangedEventArgs e)
    {
        if (_initializing || _refreshingLayout) return;
        ApplyVatToItems(Entry.OriginalItems, Entry.OriginalVatType);
        OriginalItemsGrid?.Items.Refresh();
        RefreshTotals();
    }

    private void CorrectVat_SelectionChanged(object sender, SelectionChangedEventArgs e)
    {
        if (_initializing || _refreshingLayout) return;
        Entry.PlannedVatType = VatRateCatalog.Normalize(Entry.CorrectVatType, Entry.PlannedVatType);
        ApplyVatToItems(Entry.Items, Entry.CorrectVatType);
        CorrectItemsGrid?.Items.Refresh();
        RefreshTotals();
    }

    private void RebuildOfficialPlan()
    {
        ObsidianOriginalReceipt? receipt = null;
        if (!string.IsNullOrWhiteSpace(Entry.OriginalFiscalNumber) &&
            long.TryParse(Entry.OriginalFiscalNumber, out var fiscalSign))
        {
            receipt = new ObsidianOriginalReceipt
            {
                FiscalSign = fiscalSign,
                Amount = Entry.OriginalCheckAmount ?? Entry.Amount,
                RegisteredAt = Entry.OriginalCheckDate,
                Operation = ResolveOriginalOperation(),
                Document = Entry.DocumentType.ToString(),
            };
        }

        var plan = CorrectionPlanService.Build(Entry, receipt, DateTime.Today);
        if (!plan.IsReady)
            return;

        Entry.PlannedReverseOperation = plan.Checks.Count > 1 ||
                                        Entry.CorrectionScenario == CorrectionScenario.FullCancel
            ? plan.Checks.First().Operation
            : string.Empty;
        Entry.PlannedCorrectOperation = Entry.CorrectionScenario == CorrectionScenario.FullCancel
            ? string.Empty
            : plan.Checks.Last().Operation;
        var suggestedVat = VatRateCatalog.Normalize(plan.Checks.Last().VatType, "vat122");
        if (string.IsNullOrWhiteSpace(Entry.PlannedVatType))
            Entry.PlannedVatType = suggestedVat;
        if (string.IsNullOrWhiteSpace(Entry.CorrectVatType))
            Entry.CorrectVatType = Entry.PlannedVatType;
        if (string.IsNullOrWhiteSpace(Entry.OriginalVatType))
            Entry.OriginalVatType = Entry.CorrectVatType;
    }

    private void EnsureVatSelections()
    {
        var fallback = VatRateCatalog.Normalize(Entry.PlannedVatType, "vat122");
        Entry.CorrectVatType = VatRateCatalog.Normalize(Entry.CorrectVatType, fallback);
        Entry.OriginalVatType = VatRateCatalog.Normalize(Entry.OriginalVatType, Entry.CorrectVatType);
        Entry.PlannedVatType = Entry.CorrectVatType;
        EnsureItemVatSelections();
    }

    private void EnsureItemVatSelections()
    {
        FillMissingVat(Entry.OriginalItems, Entry.OriginalVatType);
        FillMissingVat(Entry.Items, Entry.CorrectVatType);
    }

    private static void FillMissingVat(IEnumerable<OrderItem> items, string fallback)
    {
        foreach (var item in items)
            item.VatType = VatRateCatalog.Normalize(item.VatType, fallback);
    }

    private static void ApplyVatToItems(IEnumerable<OrderItem> items, string vatType)
    {
        var normalized = VatRateCatalog.Normalize(vatType, "none");
        foreach (var item in items)
            item.VatType = normalized;
    }

    private string ResolveOriginalOperation()
    {
        if (!string.IsNullOrWhiteSpace(Entry.OriginalCheckOperation))
            return Entry.OriginalCheckOperation;

        return Entry.PlannedReverseOperation switch
        {
            "sell_refund" => "sell",
            "sell" => "sell_refund",
            "buy_refund" => "buy",
            "buy" => "buy_refund",
            _ => Entry.DocumentType == SourceDocumentType.CashExpense ? "sell_refund" : "sell",
        };
    }

    private void EnsureReceiptItems()
    {
        if (!string.IsNullOrWhiteSpace(Entry.PlannedReverseOperation) && Entry.OriginalItems.Count == 0)
        {
            if (Entry.Items.Count > 0)
                Entry.OriginalItems = CloneItems(Entry.Items);
            else
                AddItem(Entry.OriginalItems, Entry.OriginalCheckAmount ?? Entry.Amount, Entry.OriginalVatType);
        }

        if (HasCorrectOrdinaryReceipt() && Entry.Items.Count == 0)
            AddItem(Entry.Items, Entry.CorrectAmount ?? Entry.Amount, Entry.CorrectVatType);

        EnsureItemVatSelections();
    }

    private void AddItem(ICollection<OrderItem> target, double amount, string vatType)
    {
        var current = target.Sum(i => i.Sum);
        var nextSum = Math.Round(amount - current, 2);
        if (nextSum <= 0)
            nextSum = Math.Round(amount > 0 ? amount : Entry.Amount, 2);

        target.Add(new OrderItem
        {
            Name = DefaultItemName(),
            Quantity = 1,
            Sum = nextSum,
            VatType = VatRateCatalog.Normalize(vatType, "none"),
        });
    }

    private void RefreshLayoutMode()
    {
        if (RepairPairPanel is null)
            return;

        var hasReverse = !string.IsNullOrWhiteSpace(Entry.PlannedReverseOperation);
        var hasCorrect = !string.IsNullOrWhiteSpace(Entry.PlannedCorrectOperation);
        var showWorkflow = hasReverse || hasCorrect;
        RepairPairPanel.Visibility = showWorkflow
            ? Visibility.Visible
            : Visibility.Collapsed;
        FallbackDataPanel.Visibility = showWorkflow
            ? Visibility.Collapsed
            : Visibility.Visible;
        ReverseSidePanel.Visibility = hasReverse ? Visibility.Visible : Visibility.Collapsed;
        CorrectSidePanel.Visibility = hasCorrect ? Visibility.Visible : Visibility.Collapsed;
        CorrectItemsPanel.Visibility = HasCorrectOrdinaryReceipt()
            ? Visibility.Visible
            : Visibility.Collapsed;
        CorrectCorrectionHint.Visibility = hasCorrect && !HasCorrectOrdinaryReceipt()
            ? Visibility.Visible
            : Visibility.Collapsed;

        _refreshingLayout = true;
        try
        {
            // OrderEntry is a persisted DTO without property notifications. Rebinding refreshes
            // calculated plan fields while the guard prevents SelectionChanged recursion.
            DataContext = null;
            DataContext = Entry;
        }
        finally
        {
            _refreshingLayout = false;
        }
        RefreshTotals();
    }

    private bool ValidateWorkflow()
    {
        var hasReverse = !string.IsNullOrWhiteSpace(Entry.PlannedReverseOperation);
        if (hasReverse)
        {
            if (!VatRateCatalog.IsKnown(Entry.OriginalVatType))
                return Warn("Выберите ставку НДС для отмены исходного чека.");

            if (string.IsNullOrWhiteSpace(Entry.OriginalFiscalNumber))
                return Warn("Для отмены исходного чека нужен ФП. Он будет записан в тег 1192.");

            if ((Entry.OriginalCheckAmount ?? 0) <= 0)
                return Warn("Заполните сумму исходного ошибочного чека.");

            if (!ValidateItems(Entry.OriginalItems, Entry.OriginalCheckAmount!.Value,
                    "исходного ошибочного чека"))
                return false;
        }

        if (HasCorrectOrdinaryReceipt())
        {
            if (!VatRateCatalog.IsKnown(Entry.CorrectVatType))
                return Warn("Выберите ставку НДС для правильного чека.");

            var correctAmount = Entry.CorrectAmount ?? Entry.Amount;
            if (correctAmount <= 0)
                return Warn("Заполните исправленную сумму правильного чека.");

            if (!ValidateItems(Entry.Items, correctAmount, "правильного чека"))
                return false;
        }

        if (!string.IsNullOrWhiteSpace(Entry.PlannedCorrectOperation) &&
            !VatRateCatalog.IsKnown(Entry.CorrectVatType))
            return Warn("Выберите ставку НДС для правильного чека.");

        return true;
    }

    private bool ValidateItems(IReadOnlyCollection<OrderItem> items, double target, string side)
    {
        if (items.Count == 0)
            return Warn($"Добавьте хотя бы одну позицию для {side}.");

        if (items.Any(item => !VatRateCatalog.IsKnown(item.VatType)))
            return Warn($"Выберите ставку НДС для каждой позиции {side}.");

        var total = Math.Round(items.Sum(i => i.Sum), 2);
        if (Math.Abs(total - target) <= 0.01)
            return true;

        return Warn($"Сумма позиций {side} ({total:N2} ₽) не равна итогу чека ({target:N2} ₽).");
    }

    private bool Warn(string message)
    {
        MessageBox.Show(message, "Проверка исправления", MessageBoxButton.OK, MessageBoxImage.Warning);
        return false;
    }

    private void RefreshTotals()
    {
        if (_initializing || _refreshingLayout ||
            OriginalItemsTotalText is null || CorrectItemsTotalText is null)
            return;

        SetTotal(OriginalItemsTotalText, Entry.OriginalItems.Sum(i => i.Sum),
            Entry.OriginalCheckAmount ?? Entry.Amount, "исходный чек");
        SetTotal(CorrectItemsTotalText, Entry.Items.Sum(i => i.Sum),
            Entry.CorrectAmount ?? Entry.Amount, "правильный чек");

        if (OriginalVatAmountText is not null)
            OriginalVatAmountText.Text = VatSummary(Entry.OriginalItems, Entry.OriginalCheckAmount ?? Entry.Amount,
                Entry.OriginalVatType);
        if (CorrectVatAmountText is not null)
            CorrectVatAmountText.Text = VatSummary(Entry.Items, Entry.CorrectAmount ?? Entry.Amount,
                Entry.CorrectVatType);
    }

    private static string VatSummary(IReadOnlyCollection<OrderItem> items, double amount, string fallbackVat)
    {
        var groups = items.Count == 0
            ? new[] { new { Vat = VatRateCatalog.Normalize(fallbackVat), Sum = VatRateCatalog.Calculate(amount, fallbackVat) } }
            : items
                .GroupBy(item => VatRateCatalog.Normalize(item.VatType, fallbackVat), StringComparer.OrdinalIgnoreCase)
                .Select(group => new
                {
                    Vat = group.Key,
                    Sum = group.Sum(item => VatRateCatalog.Calculate(item.Sum, group.Key)),
                });

        return string.Join(" · ", groups.Select(group =>
            $"{VatRateCatalog.LabelFor(group.Vat)}: {group.Sum:N2} ₽"));
    }

    private void SetTotal(TextBlock target, double total, double expected, string label)
    {
        target.Text = $"Итог позиций: {total:N2} ₽ / {label}: {expected:N2} ₽";
        var ok = Math.Abs(total - expected) <= 0.01;
        target.Foreground = TryFindResource(ok ? "BrushAccent2" : "BrushOrange") as Brush
                            ?? target.Foreground;
    }

    private bool HasCorrectOrdinaryReceipt() =>
        !string.IsNullOrWhiteSpace(Entry.PlannedCorrectOperation) &&
        !Entry.PlannedCorrectOperation.EndsWith("_correction", StringComparison.OrdinalIgnoreCase);

    private string DefaultItemName()
    {
        var docNum = !string.IsNullOrWhiteSpace(Entry.CorrectionNumber)
            ? Entry.CorrectionNumber
            : Entry.OrderNum;
        var prefix = Entry.IsService || Entry.AgentInfo is not null ? "Услуга" : "Товар";
        return Entry.DocumentType == SourceDocumentType.Realization
            ? $"{prefix} по реализации {docNum}"
            : $"Аванс по документу {docNum}";
    }

    private static List<OrderItem> CloneItems(IEnumerable<OrderItem> source) => source
        .Select(item => new OrderItem
        {
            Name = item.Name,
            Quantity = item.Quantity,
            Sum = item.Sum,
            VatType = item.VatType,
        })
        .ToList();
}
