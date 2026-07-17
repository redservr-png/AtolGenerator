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

    public OrderEntry Entry { get; private set; }

    public CorrectionScenario[] AllScenarios { get; } = Enum
        .GetValues<CorrectionScenario>()
        .Where(x => x != CorrectionScenario.RealRefund)
        .ToArray();

    public SourceDocumentType[] AllDocumentTypes { get; } =
        Enum.GetValues<SourceDocumentType>();

    public CorrectionEditorWindow(OrderEntry entry)
    {
        _initializing = true;
        InitializeComponent();
        Entry = entry;
        RebuildOfficialPlan();
        EnsureReceiptItems();
        DataContext = Entry;
        RefreshLayoutMode();
        _initializing = false;
    }

    private void Save_Click(object sender, RoutedEventArgs e)
    {
        RebuildOfficialPlan();
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
        if (_initializing || Entry is null)
            return;

        RebuildOfficialPlan();
        EnsureReceiptItems();
        RefreshLayoutMode();
    }

    private void AddOriginalItem_Click(object sender, RoutedEventArgs e)
    {
        AddItem(Entry.OriginalItems, Entry.OriginalCheckAmount ?? Entry.Amount);
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
        AddItem(Entry.Items, Entry.CorrectAmount ?? Entry.Amount);
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

    private void ItemsGrid_CurrentCellChanged(object sender, EventArgs e) => RefreshTotals();

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
        Entry.PlannedVatType = plan.Checks.Last().VatType;
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
                AddItem(Entry.OriginalItems, Entry.OriginalCheckAmount ?? Entry.Amount);
        }

        if (HasCorrectOrdinaryReceipt() && Entry.Items.Count == 0)
            AddItem(Entry.Items, Entry.CorrectAmount ?? Entry.Amount);
    }

    private void AddItem(ICollection<OrderItem> target, double amount)
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
        });
    }

    private void RefreshLayoutMode()
    {
        if (RepairPairPanel is null)
            return;

        var hasReverse = !string.IsNullOrWhiteSpace(Entry.PlannedReverseOperation);
        var hasCorrect = !string.IsNullOrWhiteSpace(Entry.PlannedCorrectOperation);
        RepairPairPanel.Visibility = hasReverse || HasCorrectOrdinaryReceipt()
            ? Visibility.Visible
            : Visibility.Collapsed;
        ReverseSidePanel.Visibility = hasReverse ? Visibility.Visible : Visibility.Collapsed;
        CorrectSidePanel.Visibility = hasCorrect ? Visibility.Visible : Visibility.Collapsed;
        CorrectItemsPanel.Visibility = HasCorrectOrdinaryReceipt()
            ? Visibility.Visible
            : Visibility.Collapsed;
        CorrectCorrectionHint.Visibility = hasCorrect && !HasCorrectOrdinaryReceipt()
            ? Visibility.Visible
            : Visibility.Collapsed;

        DataContext = null;
        DataContext = Entry;
        RefreshTotals();
    }

    private bool ValidateWorkflow()
    {
        var hasReverse = !string.IsNullOrWhiteSpace(Entry.PlannedReverseOperation);
        if (hasReverse)
        {
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
            var correctAmount = Entry.CorrectAmount ?? Entry.Amount;
            if (correctAmount <= 0)
                return Warn("Заполните исправленную сумму правильного чека.");

            if (!ValidateItems(Entry.Items, correctAmount, "правильного чека"))
                return false;
        }

        return true;
    }

    private bool ValidateItems(IReadOnlyCollection<OrderItem> items, double target, string side)
    {
        if (items.Count == 0)
            return Warn($"Добавьте хотя бы одну позицию для {side}.");

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
        if (OriginalItemsTotalText is null || CorrectItemsTotalText is null)
            return;

        SetTotal(OriginalItemsTotalText, Entry.OriginalItems.Sum(i => i.Sum),
            Entry.OriginalCheckAmount ?? Entry.Amount, "исходный чек");
        SetTotal(CorrectItemsTotalText, Entry.Items.Sum(i => i.Sum),
            Entry.CorrectAmount ?? Entry.Amount, "правильный чек");
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
        })
        .ToList();
}
