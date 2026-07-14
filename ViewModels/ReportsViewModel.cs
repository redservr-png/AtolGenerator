using System.Collections.ObjectModel;
using System.ComponentModel;
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

public sealed class ReportsViewModel : BaseViewModel
{
    private int _selectedTabIndex;
    private string _localHistoryPath = AtolApiService.PunchedJsonPath;
    private string _atolReportPath = string.Empty;
    private string _ofdReportPath = string.Empty;
    private string _xmlPath = string.Empty;
    private string _exportPath = string.Empty;
    private string _localStatus = string.Empty;
    private string _atolStatus = "Отчёт не загружен";
    private string _ofdStatus = "Отчёт не загружен";
    private string _matchingStatus = "Загрузите XML и отчёт АТОЛ";
    private AtolJournalReportRow? _selectedAtolCheck;
    private DateTime? _atolDateFrom;
    private DateTime? _atolDateTo;
    private string _atolTypeFilter = "Все";
    private string _atolSourceFilter = "Все";
    private string _atolStatusFilter = "Все";
    private string _atolSearchText = string.Empty;
    private DateTime? _ofdDateFrom;
    private DateTime? _ofdDateTo;
    private string _ofdOperationFilter = "Все";
    private string _ofdTradingPointFilter = "Все";
    private string _ofdKktFilter = "Все";
    private string _ofdSearchText = string.Empty;
    private bool _startupReportsRestored;

    private List<XmlReportCheck> _xmlChecks = new();

    public ObservableCollection<LocalPunchReportRow> LocalHistory { get; } = new();
    public ObservableCollection<AtolJournalReportRow> AtolChecks { get; } = new();
    public ObservableCollection<OfdReportRow> OfdChecks { get; } = new();
    public ObservableCollection<OneCExportRow> ExportRows { get; } = new();
    public ObservableCollection<string> AtolTypeOptions { get; } = new() { "Все" };
    public ObservableCollection<string> AtolSourceOptions { get; } = new() { "Все" };
    public ObservableCollection<string> AtolStatusOptions { get; } = new() { "Все" };
    public ObservableCollection<string> OfdOperationOptions { get; } = new() { "Все" };
    public ObservableCollection<string> OfdTradingPointOptions { get; } = new() { "Все" };
    public ObservableCollection<string> OfdKktOptions { get; } = new() { "Все" };

    public ICollectionView AtolChecksView { get; }
    public ICollectionView OfdChecksView { get; }

    public int SelectedTabIndex
    {
        get => _selectedTabIndex;
        set => Set(ref _selectedTabIndex, value);
    }

    public string LocalHistoryPath
    {
        get => _localHistoryPath;
        private set => Set(ref _localHistoryPath, value);
    }

    public string AtolReportPath
    {
        get => _atolReportPath;
        private set => Set(ref _atolReportPath, value);
    }

    public string OfdReportPath
    {
        get => _ofdReportPath;
        private set => Set(ref _ofdReportPath, value);
    }

    public string XmlPath
    {
        get => _xmlPath;
        private set => Set(ref _xmlPath, value);
    }

    public string ExportPath
    {
        get => _exportPath;
        private set => Set(ref _exportPath, value);
    }

    public string LocalStatus { get => _localStatus; private set => Set(ref _localStatus, value); }
    public string AtolStatus { get => _atolStatus; private set => Set(ref _atolStatus, value); }
    public string OfdStatus { get => _ofdStatus; private set => Set(ref _ofdStatus, value); }
    public string MatchingStatus { get => _matchingStatus; private set => Set(ref _matchingStatus, value); }

    public string AtolFileName => FileNameOrPlaceholder(AtolReportPath, "CSV АТОЛ не выбран");
    public string OfdFileName => FileNameOrPlaceholder(OfdReportPath, "Отчёт Такскома не выбран");
    public string XmlFileName => FileNameOrPlaceholder(XmlPath, "XML не выбран");
    public string LocalFileName => FileNameOrPlaceholder(LocalHistoryPath, "Локальный журнал не выбран");

    public int AtolCorrectionCount => AtolChecks.Count(x =>
        x.Operation is "sell_correction" or "buy_correction");
    public int AtolXmlCount => AtolChecks.Count(x =>
        string.Equals(x.Source, "XML", StringComparison.OrdinalIgnoreCase));
    public int ReadyCount => ExportRows.Count(x => x.IsReady);
    public int ErrorCount => ExportRows.Count(x => !x.IsReady);
    public int OfdVerifiedCount => ExportRows.Count(x => x.IsReady && x.OfdStatus == "Проверено ОФД");
    public int AtolVisibleCount => AtolChecksView.Cast<object>().Count();
    public int OfdVisibleCount => OfdChecksView.Cast<object>().Count();
    public bool CanBuildMatches => _xmlChecks.Count > 0 && AtolChecks.Count > 0;
    public bool CanExport => ReadyCount > 0;

    public AtolJournalReportRow? SelectedAtolCheck
    {
        get => _selectedAtolCheck;
        set => Set(ref _selectedAtolCheck, value);
    }

    public DateTime? AtolDateFrom
    {
        get => _atolDateFrom;
        set { if (Set(ref _atolDateFrom, value)) RefreshAtolFilter(); }
    }

    public DateTime? AtolDateTo
    {
        get => _atolDateTo;
        set { if (Set(ref _atolDateTo, value)) RefreshAtolFilter(); }
    }

    public string AtolTypeFilter
    {
        get => _atolTypeFilter;
        set
        {
            if (!Set(ref _atolTypeFilter, value ?? "Все")) return;
            OnPropertyChanged(nameof(AtolTypeFilterIndex));
            RefreshAtolFilter();
        }
    }

    public string AtolSourceFilter
    {
        get => _atolSourceFilter;
        set
        {
            if (!Set(ref _atolSourceFilter, value ?? "Все")) return;
            OnPropertyChanged(nameof(AtolSourceFilterIndex));
            RefreshAtolFilter();
        }
    }

    public string AtolStatusFilter
    {
        get => _atolStatusFilter;
        set
        {
            if (!Set(ref _atolStatusFilter, value ?? "Все")) return;
            OnPropertyChanged(nameof(AtolStatusFilterIndex));
            RefreshAtolFilter();
        }
    }

    public string AtolSearchText
    {
        get => _atolSearchText;
        set { if (Set(ref _atolSearchText, value ?? string.Empty)) RefreshAtolFilter(); }
    }

    public DateTime? OfdDateFrom
    {
        get => _ofdDateFrom;
        set { if (Set(ref _ofdDateFrom, value)) RefreshOfdFilter(); }
    }

    public DateTime? OfdDateTo
    {
        get => _ofdDateTo;
        set { if (Set(ref _ofdDateTo, value)) RefreshOfdFilter(); }
    }

    public string OfdOperationFilter
    {
        get => _ofdOperationFilter;
        set
        {
            if (!Set(ref _ofdOperationFilter, value ?? "Все")) return;
            OnPropertyChanged(nameof(OfdOperationFilterIndex));
            RefreshOfdFilter();
        }
    }

    public string OfdTradingPointFilter
    {
        get => _ofdTradingPointFilter;
        set
        {
            if (!Set(ref _ofdTradingPointFilter, value ?? "Все")) return;
            OnPropertyChanged(nameof(OfdTradingPointFilterIndex));
            RefreshOfdFilter();
        }
    }

    public string OfdKktFilter
    {
        get => _ofdKktFilter;
        set
        {
            if (!Set(ref _ofdKktFilter, value ?? "Все")) return;
            OnPropertyChanged(nameof(OfdKktFilterIndex));
            RefreshOfdFilter();
        }
    }

    public string OfdSearchText
    {
        get => _ofdSearchText;
        set { if (Set(ref _ofdSearchText, value ?? string.Empty)) RefreshOfdFilter(); }
    }

    public int AtolTypeFilterIndex
    {
        get => OptionIndex(AtolTypeOptions, AtolTypeFilter);
        set => SelectOption(AtolTypeOptions, value, selected => AtolTypeFilter = selected);
    }

    public int AtolSourceFilterIndex
    {
        get => OptionIndex(AtolSourceOptions, AtolSourceFilter);
        set => SelectOption(AtolSourceOptions, value, selected => AtolSourceFilter = selected);
    }

    public int AtolStatusFilterIndex
    {
        get => OptionIndex(AtolStatusOptions, AtolStatusFilter);
        set => SelectOption(AtolStatusOptions, value, selected => AtolStatusFilter = selected);
    }

    public int OfdOperationFilterIndex
    {
        get => OptionIndex(OfdOperationOptions, OfdOperationFilter);
        set => SelectOption(OfdOperationOptions, value, selected => OfdOperationFilter = selected);
    }

    public int OfdTradingPointFilterIndex
    {
        get => OptionIndex(OfdTradingPointOptions, OfdTradingPointFilter);
        set => SelectOption(OfdTradingPointOptions, value, selected => OfdTradingPointFilter = selected);
    }

    public int OfdKktFilterIndex
    {
        get => OptionIndex(OfdKktOptions, OfdKktFilter);
        set => SelectOption(OfdKktOptions, value, selected => OfdKktFilter = selected);
    }

    public ICommand RefreshLocalHistoryCommand { get; }
    public ICommand BrowseLocalHistoryCommand { get; }
    public ICommand OpenAtolBrowserCommand { get; }
    public ICommand ImportLatestAtolReportCommand { get; }
    public ICommand OpenAtolReportFolderCommand { get; }
    public ICommand OpenTaxcomBrowserCommand { get; }
    public ICommand ImportLatestOfdReportCommand { get; }
    public ICommand OpenTaxcomReportFolderCommand { get; }
    public ICommand ImportAtolReportCommand { get; }
    public ICommand ImportOfdReportCommand { get; }
    public ICommand ImportXmlCommand { get; }
    public ICommand BuildMatchesCommand { get; }
    public ICommand ExportOneCCsvCommand { get; }
    public ICommand OpenReceiptCommand { get; }
    public ICommand ClearAtolFiltersCommand { get; }
    public ICommand ClearOfdFiltersCommand { get; }

    public ReportsViewModel()
    {
        AtolChecksView = CollectionViewSource.GetDefaultView(AtolChecks);
        AtolChecksView.Filter = FilterAtolCheck;
        OfdChecksView = CollectionViewSource.GetDefaultView(OfdChecks);
        OfdChecksView.Filter = FilterOfdCheck;

        RefreshLocalHistoryCommand = new RelayCommand(_ => LoadLocalHistory(LocalHistoryPath));
        BrowseLocalHistoryCommand = new RelayCommand(_ => BrowseLocalHistory());
        OpenAtolBrowserCommand = new RelayCommand(_ => OpenAtolBrowser());
        ImportLatestAtolReportCommand = new RelayCommand(_ => ImportLatestAtolReport());
        OpenAtolReportFolderCommand = new RelayCommand(_ => FileHelper.OpenFolder(FileHelper.AtolReportDir));
        OpenTaxcomBrowserCommand = new RelayCommand(_ => OpenTaxcomBrowser());
        ImportLatestOfdReportCommand = new RelayCommand(_ => ImportLatestOfdReport());
        OpenTaxcomReportFolderCommand = new RelayCommand(_ => FileHelper.OpenFolder(FileHelper.TaxcomReportDir));
        ImportAtolReportCommand = new RelayCommand(_ => BrowseAtolReport());
        ImportOfdReportCommand = new RelayCommand(_ => BrowseOfdReport());
        ImportXmlCommand = new RelayCommand(_ => BrowseXml());
        BuildMatchesCommand = new RelayCommand(_ => BuildMatches(true), _ => CanBuildMatches);
        ExportOneCCsvCommand = new RelayCommand(_ => ExportOneCCsv(), _ => CanExport);
        OpenReceiptCommand = new RelayCommand(_ => OpenSelectedReceipt(), _ =>
            !string.IsNullOrWhiteSpace(SelectedAtolCheck?.OfdUrl));
        ClearAtolFiltersCommand = new RelayCommand(_ => ClearAtolFilters());
        ClearOfdFiltersCommand = new RelayCommand(_ => ClearOfdFilters());

        LoadLocalHistory(LocalHistoryPath);
    }

    public void LoadLocalHistory(string path)
    {
        LocalHistoryPath = path;
        OnPropertyChanged(nameof(LocalFileName));
        LocalHistory.Clear();

        if (!File.Exists(path))
        {
            LocalStatus = "Локальный журнал ещё не создан";
            return;
        }

        try
        {
            foreach (var row in ReportImportService.ReadLocalPunchHistory(path))
                LocalHistory.Add(row);
            LocalStatus = $"Загружено чеков: {LocalHistory.Count}";
        }
        catch (Exception ex)
        {
            LocalStatus = $"Ошибка чтения: {ex.Message}";
        }
    }

    public void LoadAtolReport(string path)
    {
        ApplyAtolReport(path, ReportImportService.ReadAtolJournal(path), selectTab: true);
    }

    public void LoadOfdReport(string path)
    {
        ApplyOfdReport(path, ReportImportService.ReadOfdReport(path), selectTab: true);
    }

    public async Task RestoreLatestReportsAsync()
    {
        if (_startupReportsRestored) return;
        _startupReportsRestored = true;

        AtolStatus = "Ищем последний отчёт...";
        OfdStatus = "Ищем последний отчёт...";

        var settings = ApplicationSettingsStore.Current;
        var atolTask = Task.Run(() => TryReadLatestReport(
            FileHelper.AtolReportDir,
            "*.csv",
            name => name.Contains("report_checkjournal", StringComparison.OrdinalIgnoreCase) ||
                    name.Contains("атол", StringComparison.OrdinalIgnoreCase) ||
                    name.Contains("журнал", StringComparison.OrdinalIgnoreCase),
            settings.LastAtolReportPath,
            ReportImportService.ReadAtolJournal));
        var ofdTask = Task.Run(() => TryReadLatestReport(
            FileHelper.TaxcomReportDir,
            "*.xlsx",
            name => name.Contains("такском", StringComparison.OrdinalIgnoreCase) ||
                    name.Contains("фискальн", StringComparison.OrdinalIgnoreCase) ||
                    name.Contains("сводн", StringComparison.OrdinalIgnoreCase),
            settings.LastOfdReportPath,
            ReportImportService.ReadOfdReport));

        await Task.WhenAll(atolTask, ofdTask);

        var atolReport = await atolTask;
        if (atolReport is not null)
            ApplyAtolReport(atolReport.Value.Path, atolReport.Value.Rows, selectTab: false, automatic: true);
        else
            AtolStatus = "Последний отчёт не найден";

        var ofdReport = await ofdTask;
        if (ofdReport is not null)
            ApplyOfdReport(ofdReport.Value.Path, ofdReport.Value.Rows, selectTab: false, automatic: true);
        else
            OfdStatus = "Последний отчёт не найден";
    }

    private void ApplyAtolReport(
        string path,
        IReadOnlyCollection<AtolJournalReportRow> rows,
        bool selectTab,
        bool automatic = false)
    {
        AtolChecks.Clear();
        foreach (var row in rows) AtolChecks.Add(row);

        AtolReportPath = path;
        AtolStatus = $"{(automatic ? "Автозагрузка · " : string.Empty)}{AtolChecks.Count} чеков · XML: {AtolXmlCount} · коррекции: {AtolCorrectionCount}";
        OnPropertyChanged(nameof(AtolFileName));
        OnPropertyChanged(nameof(AtolCorrectionCount));
        OnPropertyChanged(nameof(AtolXmlCount));
        OnPropertyChanged(nameof(CanBuildMatches));
        RefreshAtolOptions();
        ClearAtolFilters();
        if (selectTab) SelectedTabIndex = 1;
        RememberReportPath(atolPath: path);
        BuildMatches(false);
        CommandManager.InvalidateRequerySuggested();
    }

    private void ApplyOfdReport(
        string path,
        IReadOnlyCollection<OfdReportRow> rows,
        bool selectTab,
        bool automatic = false)
    {
        OfdChecks.Clear();
        foreach (var row in rows) OfdChecks.Add(row);

        OfdReportPath = path;
        OfdStatus = $"{(automatic ? "Автозагрузка · " : string.Empty)}загружено документов: {OfdChecks.Count}";
        OnPropertyChanged(nameof(OfdFileName));
        RefreshOfdOptions();
        ClearOfdFilters();
        if (selectTab) SelectedTabIndex = 2;
        RememberReportPath(ofdPath: path);
        BuildMatches(false);
    }

    private bool FilterAtolCheck(object item)
    {
        if (item is not AtolJournalReportRow row) return false;
        if (!MatchesDate(row.RegisteredAt, AtolDateFrom, AtolDateTo)) return false;
        if (!MatchesOption(row.CheckType, AtolTypeFilter)) return false;
        if (!MatchesOption(row.Source, AtolSourceFilter)) return false;
        if (!MatchesOption(row.Status, AtolStatusFilter)) return false;

        return MatchesSearch(AtolSearchText,
            row.CheckType, row.Operation, row.Source, row.Status,
            row.FiscalDocument?.ToString(), row.FiscalSign?.ToString(),
            row.ExternalId, row.Uuid, row.BaseNumber);
    }

    private bool FilterOfdCheck(object item)
    {
        if (item is not OfdReportRow row) return false;
        if (!MatchesDate(row.RegisteredAt, OfdDateFrom, OfdDateTo)) return false;
        if (!MatchesOption(row.Operation, OfdOperationFilter)) return false;
        if (!MatchesOption(row.TradingPoint, OfdTradingPointFilter)) return false;
        if (!MatchesOption(row.KktName, OfdKktFilter)) return false;

        return MatchesSearch(OfdSearchText,
            row.Document, row.Operation, row.TradingPoint, row.KktName,
            row.FiscalDocument?.ToString(), row.FiscalSign?.ToString(), row.ReceiptUrl);
    }

    private void RefreshAtolOptions()
    {
        ResetOptions(AtolTypeOptions, AtolChecks.Select(row => row.CheckType));
        ResetOptions(AtolSourceOptions, AtolChecks.Select(row => row.Source));
        ResetOptions(AtolStatusOptions, AtolChecks.Select(row => row.Status));
        OnPropertyChanged(nameof(AtolTypeFilterIndex));
        OnPropertyChanged(nameof(AtolSourceFilterIndex));
        OnPropertyChanged(nameof(AtolStatusFilterIndex));
    }

    private void RefreshOfdOptions()
    {
        ResetOptions(OfdOperationOptions, OfdChecks.Select(row => row.Operation));
        ResetOptions(OfdTradingPointOptions, OfdChecks.Select(row => row.TradingPoint));
        ResetOptions(OfdKktOptions, OfdChecks.Select(row => row.KktName));
        OnPropertyChanged(nameof(OfdOperationFilterIndex));
        OnPropertyChanged(nameof(OfdTradingPointFilterIndex));
        OnPropertyChanged(nameof(OfdKktFilterIndex));
    }

    private void ClearAtolFilters()
    {
        _atolDateFrom = null;
        _atolDateTo = null;
        _atolTypeFilter = string.Empty;
        _atolSourceFilter = string.Empty;
        _atolStatusFilter = string.Empty;
        _atolSearchText = string.Empty;
        OnPropertyChanged(nameof(AtolDateFrom));
        OnPropertyChanged(nameof(AtolDateTo));
        OnPropertyChanged(nameof(AtolSearchText));
        AtolTypeFilter = "Все";
        AtolSourceFilter = "Все";
        AtolStatusFilter = "Все";
    }

    private void ClearOfdFilters()
    {
        _ofdDateFrom = null;
        _ofdDateTo = null;
        _ofdOperationFilter = string.Empty;
        _ofdTradingPointFilter = string.Empty;
        _ofdKktFilter = string.Empty;
        _ofdSearchText = string.Empty;
        OnPropertyChanged(nameof(OfdDateFrom));
        OnPropertyChanged(nameof(OfdDateTo));
        OnPropertyChanged(nameof(OfdSearchText));
        OfdOperationFilter = "Все";
        OfdTradingPointFilter = "Все";
        OfdKktFilter = "Все";
    }

    private void RefreshAtolFilter()
    {
        AtolChecksView.Refresh();
        OnPropertyChanged(nameof(AtolVisibleCount));
    }

    private void RefreshOfdFilter()
    {
        OfdChecksView.Refresh();
        OnPropertyChanged(nameof(OfdVisibleCount));
    }

    private static bool MatchesDate(DateTime? value, DateTime? from, DateTime? to)
    {
        if (!from.HasValue && !to.HasValue) return true;
        if (!value.HasValue) return false;
        if (from.HasValue && value.Value.Date < from.Value.Date) return false;
        return !to.HasValue || value.Value.Date <= to.Value.Date;
    }

    private static bool MatchesOption(string value, string filter) =>
        string.IsNullOrWhiteSpace(filter) ||
        string.Equals(filter, "Все", StringComparison.OrdinalIgnoreCase) ||
        string.Equals(value, filter, StringComparison.OrdinalIgnoreCase);

    private static bool MatchesSearch(string searchText, params string?[] values)
    {
        var search = searchText?.Trim();
        return string.IsNullOrWhiteSpace(search) || values.Any(value =>
            value?.Contains(search, StringComparison.OrdinalIgnoreCase) == true);
    }

    private static void ResetOptions(
        ObservableCollection<string> options,
        IEnumerable<string> values)
    {
        options.Clear();
        options.Add("Все");
        foreach (var value in values
                     .Where(value => !string.IsNullOrWhiteSpace(value))
                     .Distinct(StringComparer.OrdinalIgnoreCase)
                     .OrderBy(value => value, StringComparer.CurrentCultureIgnoreCase))
            options.Add(value);
    }

    private static int OptionIndex(ObservableCollection<string> options, string selected)
    {
        var index = options.IndexOf(selected);
        return index >= 0 ? index : 0;
    }

    private static void SelectOption(
        ObservableCollection<string> options,
        int index,
        Action<string> select)
    {
        if (index >= 0 && index < options.Count)
            select(options[index]);
    }

    public void LoadXml(string path)
    {
        _xmlChecks = ReportImportService.ReadXmlChecks(path);
        XmlPath = path;
        OnPropertyChanged(nameof(XmlFileName));
        OnPropertyChanged(nameof(CanBuildMatches));
        SelectedTabIndex = 3;
        MatchingStatus = $"XML: {_xmlChecks.Count} чеков";
        BuildMatches(false);
        CommandManager.InvalidateRequerySuggested();
    }

    public void BuildMatches(bool showErrors)
    {
        if (!CanBuildMatches)
        {
            ExportRows.Clear();
            MatchingStatus = _xmlChecks.Count == 0
                ? "Загрузите XML с пробитыми чеками"
                : "Загрузите CSV из журнала АТОЛ";
            RefreshExportCounters();
            return;
        }

        try
        {
            var rows = ReportReconciliationService.Build(_xmlChecks, AtolChecks, OfdChecks);
            ExportRows.Clear();
            foreach (var row in rows) ExportRows.Add(row);
            MatchingStatus = $"Готово: {ReadyCount} · требуют внимания: {ErrorCount} · проверено ОФД: {OfdVerifiedCount}";
            SelectedTabIndex = 3;
            RefreshExportCounters();
        }
        catch (Exception ex)
        {
            MatchingStatus = $"Ошибка сопоставления: {ex.Message}";
            if (showErrors)
                MessageBox.Show(MatchingStatus, "Работа с отчётами", MessageBoxButton.OK, MessageBoxImage.Error);
        }
    }

    private void BrowseLocalHistory()
    {
        var dialog = new OpenFileDialog
        {
            Title = "Выберите локальный журнал пробитых чеков",
            Filter = "Журнал JSONL|*.jsonl|Все файлы|*.*",
            FileName = Path.GetFileName(LocalHistoryPath),
            InitialDirectory = Path.GetDirectoryName(LocalHistoryPath),
        };
        if (dialog.ShowDialog() == true) LoadLocalHistory(dialog.FileName);
    }

    private void BrowseAtolReport()
    {
        var dialog = new OpenFileDialog
        {
            Title = "Выберите CSV из журнала чеков АТОЛ Online",
            Filter = "CSV АТОЛ|*.csv|Все файлы|*.*",
        };
        if (dialog.ShowDialog() != true) return;
        TryLoad(() => LoadAtolReport(dialog.FileName), "Не удалось загрузить отчёт АТОЛ");
    }

    private void BrowseOfdReport()
    {
        var dialog = new OpenFileDialog
        {
            Title = "Выберите сводный отчёт ОФД / Такскома",
            Filter = "Отчёт Excel|*.xlsx|Все файлы|*.*",
        };
        if (dialog.ShowDialog() != true) return;
        TryLoad(() => LoadOfdReport(dialog.FileName), "Не удалось загрузить отчёт ОФД");
    }

    private void OpenTaxcomBrowser()
    {
        try
        {
            OfdStatus = "Открываем личный кабинет Такскома...";
            var browser = new ReportBrowserWindow(ReportPortal.Taxcom)
            {
                Owner = Application.Current?.MainWindow,
            };

            if (browser.ShowDialog() != true || string.IsNullOrWhiteSpace(browser.DownloadedReportPath))
            {
                OfdStatus = "Получение отчёта отменено";
                return;
            }

            TryLoad(
                () => LoadOfdReport(browser.DownloadedReportPath),
                "Не удалось загрузить скачанный отчёт Такскома");
        }
        catch (Exception ex)
        {
            OfdStatus = $"Не удалось открыть Такском: {ex.Message}";
            MessageBox.Show(OfdStatus, "Такском-Касса", MessageBoxButton.OK, MessageBoxImage.Error);
        }
    }

    private void OpenAtolBrowser()
    {
        try
        {
            AtolStatus = "Открываем личный кабинет АТОЛ Online...";
            var browser = new ReportBrowserWindow(ReportPortal.AtolOnline)
            {
                Owner = Application.Current?.MainWindow,
            };

            if (browser.ShowDialog() != true || string.IsNullOrWhiteSpace(browser.DownloadedReportPath))
            {
                AtolStatus = "Получение отчёта отменено";
                return;
            }

            TryLoad(
                () => LoadAtolReport(browser.DownloadedReportPath),
                "Не удалось загрузить скачанный отчёт АТОЛ Online");
        }
        catch (Exception ex)
        {
            AtolStatus = $"Не удалось открыть АТОЛ Online: {ex.Message}";
            MessageBox.Show(AtolStatus, "АТОЛ Online", MessageBoxButton.OK, MessageBoxImage.Error);
        }
    }

    private void ImportLatestAtolReport()
    {
        var reportPath = FindLatestValidAtolReport();
        if (reportPath is null)
        {
            MessageBox.Show(
                "В папке отчётов программы и в папке «Загрузки» не найден подходящий CSV журнала чеков АТОЛ.",
                "Отчёт АТОЛ Online", MessageBoxButton.OK, MessageBoxImage.Information);
            return;
        }

        TryLoad(() => LoadAtolReport(reportPath), "Не удалось загрузить последний отчёт АТОЛ Online");
    }

    private void ImportLatestOfdReport()
    {
        var reportPath = FindLatestValidTaxcomReport();
        if (reportPath is null)
        {
            MessageBox.Show(
                "В папке отчётов программы и в папке «Загрузки» не найден подходящий отчёт Такскома XLSX.",
                "Отчёт Такскома", MessageBoxButton.OK, MessageBoxImage.Information);
            return;
        }

        TryLoad(() => LoadOfdReport(reportPath), "Не удалось загрузить последний отчёт Такскома");
    }

    internal static string? FindLatestValidTaxcomReport()
    {
        return FindLatestValidReport(
            FileHelper.TaxcomReportDir,
            "*.xlsx",
            name => name.Contains("такском", StringComparison.OrdinalIgnoreCase) ||
                    name.Contains("фискальн", StringComparison.OrdinalIgnoreCase) ||
                    name.Contains("сводн", StringComparison.OrdinalIgnoreCase),
            path => ReportImportService.ReadOfdReport(path));
    }

    private static string? FindLatestValidAtolReport()
    {
        return FindLatestValidReport(
            FileHelper.AtolReportDir,
            "*.csv",
            name => name.Contains("report_checkjournal", StringComparison.OrdinalIgnoreCase) ||
                    name.Contains("атол", StringComparison.OrdinalIgnoreCase) ||
                    name.Contains("журнал", StringComparison.OrdinalIgnoreCase),
            path => ReportImportService.ReadAtolJournal(path));
    }

    private static string? FindLatestValidReport(
        string managedDirectory,
        string searchPattern,
        Func<string, bool> isLikelyReport,
        Action<string> validate)
    {
        foreach (var path in GetReportCandidates(
                     managedDirectory, searchPattern, isLikelyReport, string.Empty))
        {
            try
            {
                validate(path);
                return path;
            }
            catch
            {
                // Перебираем только свежие кандидаты, пока не найдём отчёт нужного формата.
            }
        }

        return null;
    }

    private static (string Path, List<T> Rows)? TryReadLatestReport<T>(
        string managedDirectory,
        string searchPattern,
        Func<string, bool> isLikelyReport,
        string rememberedPath,
        Func<string, List<T>> read)
    {
        try
        {
            foreach (var path in GetReportCandidates(
                         managedDirectory, searchPattern, isLikelyReport, rememberedPath))
            {
                try
                {
                    return (path, read(path));
                }
                catch
                {
                    // Следующий файл может быть более старым, но иметь корректный формат отчёта.
                }
            }
        }
        catch
        {
            // Ошибка доступа к папке не должна мешать запуску программы.
        }

        return null;
    }

    private static IEnumerable<string> GetReportCandidates(
        string managedDirectory,
        string searchPattern,
        Func<string, bool> isLikelyReport,
        string rememberedPath)
    {
        var candidates = new List<string>();
        if (!string.IsNullOrWhiteSpace(rememberedPath) && File.Exists(rememberedPath))
            candidates.Add(rememberedPath);

        AddCandidates(candidates, managedDirectory, searchPattern, includeEveryFile: true, isLikelyReport);

        var downloads = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.UserProfile), "Downloads");
        AddCandidates(candidates, downloads, searchPattern, includeEveryFile: false, isLikelyReport);

        return candidates
            .Distinct(StringComparer.OrdinalIgnoreCase)
            .OrderByDescending(File.GetLastWriteTimeUtc)
            .ToList();
    }

    private static void AddCandidates(
        List<string> target,
        string directory,
        string searchPattern,
        bool includeEveryFile,
        Func<string, bool> isLikelyReport)
    {
        if (!Directory.Exists(directory)) return;

        foreach (var path in Directory.EnumerateFiles(directory, searchPattern, SearchOption.TopDirectoryOnly))
        {
            var name = Path.GetFileNameWithoutExtension(path);
            if (includeEveryFile || isLikelyReport(name)) target.Add(path);
        }
    }

    private static void RememberReportPath(string? atolPath = null, string? ofdPath = null)
    {
        try
        {
            var settings = ApplicationSettingsStore.Current;
            if (!string.IsNullOrWhiteSpace(atolPath)) settings.LastAtolReportPath = atolPath;
            if (!string.IsNullOrWhiteSpace(ofdPath)) settings.LastOfdReportPath = ofdPath;
            ApplicationSettingsStore.Save(settings);
        }
        catch
        {
            // Сам отчёт уже загружен; сбой записи пути не должен отменять результат.
        }
    }

    private void BrowseXml()
    {
        var dialog = new OpenFileDialog
        {
            Title = "Выберите XML, загруженный в АТОЛ Online",
            Filter = "XML|*.xml|Все файлы|*.*",
        };
        if (dialog.ShowDialog() != true) return;
        TryLoad(() => LoadXml(dialog.FileName), "Не удалось загрузить XML");
    }

    private void ExportOneCCsv()
    {
        var dialog = new SaveFileDialog
        {
            Title = "Сохранить CSV для обработки 1С",
            Filter = "CSV для 1С|*.csv",
            FileName = $"atol_to_1c_{DateTime.Now:yyyyMMdd_HHmmss}.csv",
            InitialDirectory = FileHelper.OutputDir,
        };
        if (dialog.ShowDialog() != true) return;

        try
        {
            ReportReconciliationService.ExportOneCCsv(dialog.FileName, ExportRows);
            ExportPath = dialog.FileName;
            MatchingStatus = $"CSV для 1С сохранён: {Path.GetFileName(dialog.FileName)}";
            MessageBox.Show($"Экспортировано строк: {ReadyCount}\n\n{dialog.FileName}",
                "CSV для 1С готов", MessageBoxButton.OK, MessageBoxImage.Information);
        }
        catch (Exception ex)
        {
            MessageBox.Show($"Не удалось сохранить CSV: {ex.Message}",
                "Ошибка экспорта", MessageBoxButton.OK, MessageBoxImage.Error);
        }
    }

    private void OpenSelectedReceipt()
    {
        var url = SelectedAtolCheck?.OfdUrl;
        if (string.IsNullOrWhiteSpace(url)) return;
        System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo(url)
        {
            UseShellExecute = true,
        });
    }

    private static void TryLoad(Action action, string title)
    {
        try
        {
            action();
        }
        catch (Exception ex)
        {
            MessageBox.Show($"{title}:\n{ex.Message}", "Работа с отчётами",
                MessageBoxButton.OK, MessageBoxImage.Error);
        }
    }

    private void RefreshExportCounters()
    {
        OnPropertyChanged(nameof(ReadyCount));
        OnPropertyChanged(nameof(ErrorCount));
        OnPropertyChanged(nameof(OfdVerifiedCount));
        OnPropertyChanged(nameof(CanExport));
        CommandManager.InvalidateRequerySuggested();
    }

    private static string FileNameOrPlaceholder(string path, string placeholder) =>
        string.IsNullOrWhiteSpace(path) ? placeholder : Path.GetFileName(path);
}
