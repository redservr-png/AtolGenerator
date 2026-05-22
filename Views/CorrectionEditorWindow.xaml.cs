using System.Windows;
using AtolGenerator.Models;

namespace AtolGenerator.Views;

/// <summary>
/// Модальное окно редактирования параметров коррекции для строки OrderEntry.
/// На вход — копия OrderEntry, чтобы отмена не затронула оригинал.
/// </summary>
public partial class CorrectionEditorWindow : Window
{
    /// <summary>Редактируемая (изменённая) копия записи.</summary>
    public OrderEntry Entry { get; }

    /// <summary>Списки enum-значений для combobox'ов.</summary>
    public Array AllScenarios     { get; } = Enum.GetValues(typeof(CorrectionScenario));
    public Array AllDocumentTypes { get; } = Enum.GetValues(typeof(SourceDocumentType));

    public CorrectionEditorWindow(OrderEntry entry)
    {
        InitializeComponent();
        Entry       = entry;
        DataContext = entry;
    }

    private void Save_Click(object sender, RoutedEventArgs e)
    {
        // Пересчитываем Kind по выбранному сценарию — иначе UI не покажет новый цвет/чип
        Entry.Kind = Entry.CorrectionScenario.ToOrderKind();
        DialogResult = true;
        Close();
    }

    private void Cancel_Click(object sender, RoutedEventArgs e)
    {
        DialogResult = false;
        Close();
    }
}
