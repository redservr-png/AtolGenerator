using System.Windows.Controls;
using AtolGenerator.ViewModels;

namespace AtolGenerator.Views;

public partial class ObsidianCasesView : UserControl
{
    public ObsidianCasesView()
    {
        InitializeComponent();
    }

    private void CasesGrid_CurrentCellChanged(object sender, EventArgs e)
    {
        if (sender is not DataGrid { CurrentItem: ObsidianCaseItemViewModel item } ||
            DataContext is not ObsidianCasesViewModel viewModel)
            return;

        viewModel.SelectedCase = item;
    }
}
