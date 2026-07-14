using System.Windows;
using System.Windows.Controls;
using System.Windows.Media;
using AtolGenerator.Models;

namespace AtolGenerator.Views;

/// <summary>
/// Модальное окно редактирования параметров коррекции для строки OrderEntry.
/// На вход — копия OrderEntry, чтобы отмена не затронула оригинал.
/// </summary>
public partial class CorrectionEditorWindow : Window
{
    /// <summary>Редактируемая (изменённая) копия записи.</summary>
    public OrderEntry Entry { get; private set; } = null!;

    /// <summary>Списки enum-значений для combobox'ов.</summary>
    public Array AllScenarios     { get; } = Enum.GetValues(typeof(CorrectionScenario));
    public Array AllDocumentTypes { get; } = Enum.GetValues(typeof(SourceDocumentType));

    public CorrectionEditorWindow(OrderEntry entry)
    {
        InitializeComponent();
        Entry       = entry;
        DataContext = entry;
        EnsureRepairPairItems();
        RefreshLayoutMode();
    }

    private void Save_Click(object sender, RoutedEventArgs e)
    {
        if (!ValidateRepairPair())
            return;

        // Для реализаций с уже пробитым чеком "другая дата" — это именно
        // исправительный комплект: sell_refund + sell_correction.
        Entry.Kind = IsRepairPair()
            ? OrderKind.RefundCorrectionPair
            : Entry.CorrectionScenario.ToOrderKind();

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
        if (Entry is null)
            return;

        EnsureRepairPairItems();
        RefreshLayoutMode();
    }

    private void AddRepairItem_Click(object sender, RoutedEventArgs e)
    {
        var target = Entry.OriginalCheckAmount ?? Entry.CorrectAmount ?? Entry.Amount;
        var current = Entry.Items.Sum(i => i.Sum);
        var nextSum = Math.Round(target - current, 2);
        if (nextSum <= 0)
            nextSum = Math.Round(target > 0 ? target : Entry.Amount, 2);

        Entry.Items.Add(new OrderItem
        {
            Name     = DefaultRepairItemName(),
            Quantity = 1,
            Sum      = nextSum,
        });

        RepairItemsGrid.Items.Refresh();
        RefreshRepairTotals();
    }

    private void DeleteRepairItem_Click(object sender, RoutedEventArgs e)
    {
        if (RepairItemsGrid.SelectedItem is not OrderItem item)
            return;

        Entry.Items.Remove(item);
        RepairItemsGrid.Items.Refresh();
        RefreshRepairTotals();
    }

    private void RepairItemsGrid_CurrentCellChanged(object sender, EventArgs e)
    {
        RefreshRepairTotals();
    }

    private void RefreshLayoutMode()
    {
        if (RepairPairPanel is null)
            return;

        var isRepair = IsRepairPair();
        RepairPairPanel.Visibility = isRepair ? Visibility.Visible : Visibility.Collapsed;
        RefreshRepairTotals();
    }

    private void EnsureRepairPairItems()
    {
        if (!IsRepairPair() || Entry.Items.Count > 0)
            return;

        var amount = Entry.OriginalCheckAmount ?? Entry.CorrectAmount ?? Entry.Amount;
        Entry.Items.Add(new OrderItem
        {
            Name     = DefaultRepairItemName(),
            Quantity = 1,
            Sum      = amount,
        });
    }

    private bool ValidateRepairPair()
    {
        if (!IsRepairPair())
            return true;

        if (string.IsNullOrWhiteSpace(Entry.OriginalFiscalNumber))
        {
            MessageBox.Show(
                "Для исправительного комплекта нужен ФП исходного чека из 1С или отчёта ОФД.",
                "Проверка исправительного комплекта",
                MessageBoxButton.OK,
                MessageBoxImage.Warning);
            return false;
        }

        if ((Entry.CorrectAmount ?? 0) <= 0)
        {
            MessageBox.Show(
                "Заполните исправленную сумму для чека коррекции.",
                "Проверка исправительного комплекта",
                MessageBoxButton.OK,
                MessageBoxImage.Warning);
            return false;
        }

        if ((Entry.OriginalCheckAmount ?? 0) <= 0)
        {
            MessageBox.Show(
                "Заполните сумму старого ошибочного чека для возврата.",
                "Проверка исправительного комплекта",
                MessageBoxButton.OK,
                MessageBoxImage.Warning);
            return false;
        }

        var itemsTotal = Entry.Items.Sum(i => i.Sum);
        if (Entry.Items.Count == 0 || itemsTotal <= 0)
        {
            MessageBox.Show(
                "Добавьте хотя бы одну позицию в табличную часть возврата.",
                "Проверка исправительного комплекта",
                MessageBoxButton.OK,
                MessageBoxImage.Warning);
            return false;
        }

        var target = Entry.OriginalCheckAmount!.Value;
        if (Math.Abs(itemsTotal - target) > 0.01)
        {
            MessageBox.Show(
                $"Сумма табличной части возврата ({itemsTotal:N2}) не равна сумме старого чека ({target:N2}).\n\n" +
                "Поправьте позиции или сумму старого чека, чтобы XML не разошёлся по итогам.",
                "Проверка исправительного комплекта",
                MessageBoxButton.OK,
                MessageBoxImage.Warning);
            return false;
        }

        return true;
    }

    private void RefreshRepairTotals()
    {
        if (RefundItemsTotalText is null || Entry is null)
            return;

        var itemsTotal = Entry.Items.Sum(i => i.Sum);
        var target = Entry.OriginalCheckAmount ?? Entry.CorrectAmount ?? Entry.Amount;
        RefundItemsTotalText.Text = $"Итог позиций: {itemsTotal:N2} ₽ / сумма старого чека: {target:N2} ₽";

        var ok = Math.Abs(itemsTotal - target) <= 0.01;
        RefundItemsTotalText.Foreground = TryFindResource(ok ? "BrushAccent2" : "BrushOrange") as Brush
                                          ?? RefundItemsTotalText.Foreground;
    }

    private bool IsRepairPair()
        => (!string.IsNullOrWhiteSpace(Entry.PlannedReverseOperation) &&
            !string.IsNullOrWhiteSpace(Entry.PlannedCorrectOperation)) ||
           (Entry.DocumentType == SourceDocumentType.Realization &&
            Entry.CorrectionScenario == CorrectionScenario.WrongDate);

    private string DefaultRepairItemName()
    {
        var docNum = !string.IsNullOrWhiteSpace(Entry.CorrectionNumber)
            ? Entry.CorrectionNumber
            : Entry.OrderNum;
        var prefix = Entry.IsService || Entry.AgentInfo is not null ? "Услуга" : "Товар";
        return Entry.DocumentType == SourceDocumentType.Realization
            ? $"{prefix} по реализации {docNum}"
            : $"{prefix} по документу {docNum}";
    }
}
