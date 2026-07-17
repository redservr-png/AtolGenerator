using System.Collections.ObjectModel;
using System.IO;
using System.Windows.Input;
using AtolGenerator.Constants;
using AtolGenerator.Helpers;
using AtolGenerator.Models;
using AtolGenerator.Services;
using Microsoft.Win32;

namespace AtolGenerator.ViewModels;

public sealed class ThemeOption
{
    public required string Key         { get; init; }
    public required string Name        { get; init; }
    public required string Description { get; init; }
    public required string Background  { get; init; }
    public required string Surface     { get; init; }
    public required string Accent      { get; init; }
}

public sealed class SettingsViewModel : BaseViewModel
{
    private readonly string _originalThemeKey;
    private string _section = "connections";
    private ThemeOption? _selectedTheme;
    private CashierInfo? _selectedCashier;
    private CashierInfo? _defaultCashier;
    private ServiceProvider? _selectedAgent;
    private string _atolStatus = string.Empty;
    private string _oneCStatus = string.Empty;
    private string _validationMessage = string.Empty;
    private string _obsidianFilePath = string.Empty;
    private bool _autoValidateServiceNotes;
    private string _updateStatus = "Обновления ещё не проверялись";
    private string _availableUpdateUrl = string.Empty;

    public SettingsViewModel()
    {
        var appSettings = ApplicationSettingsStore.Current;
        _originalThemeKey = appSettings.ThemeKey;

        ThemeOptions = new ObservableCollection<ThemeOption>
        {
            new()
            {
                Key = "light", Name = "Светлая", Description = "Чистый рабочий интерфейс",
                Background = "#F5F7FA", Surface = "#FFFFFF", Accent = "#155EEF",
            },
            new()
            {
                Key = "dark", Name = "Тёмная", Description = "Спокойная работа вечером",
                Background = "#0C111D", Surface = "#1D2939", Accent = "#84ADFF",
            },
            new()
            {
                Key = "warm", Name = "Тёплая", Description = "Мягкая палитра в стиле Claude",
                Background = "#F7F5F2", Surface = "#FFFFFF", Accent = "#C55232",
            },
        };

        Cashiers = new ObservableCollection<CashierInfo>(appSettings.Cashiers.Select(CloneCashier));
        Agents = new ObservableCollection<ServiceProvider>(appSettings.Agents.Select(CloneAgent));
        DefaultCashier = Cashiers.FirstOrDefault(c => string.Equals(
            c.ShortName, appSettings.SelectedCashierShortName, StringComparison.OrdinalIgnoreCase))
            ?? Cashiers.FirstOrDefault();
        SelectedTheme = ThemeOptions.FirstOrDefault(t => t.Key == appSettings.ThemeKey)
                        ?? ThemeOptions[0];

        var atol = AtolCredentials.Load();
        AtolLogin = atol.Login;
        AtolPassword = atol.Password;
        AtolGroupCode = atol.GroupCode;

        var oneC = OneCConnectionSettings.Load();
        OneCServer = oneC.Server;
        OneCDatabase = oneC.Database;
        OneCUser = oneC.User;
        OneCPassword = oneC.Password;

        ObsidianFilePath = !string.IsNullOrWhiteSpace(appSettings.ObsidianFilePath)
            ? appSettings.ObsidianFilePath
            : ObsidianSettings.Load().MdFilePath;
        AutoValidateServiceNotes = appSettings.AutoValidateServiceNotes;

        SelectSectionCommand = new RelayCommand(value => Section = value?.ToString() ?? "connections");
        AddCashierCommand = new RelayCommand(AddCashier);
        DeleteCashierCommand = new RelayCommand(DeleteCashier);
        AddAgentCommand = new RelayCommand(AddAgent);
        DeleteAgentCommand = new RelayCommand(DeleteAgent);
        SaveCommand = new RelayCommand(Save);
        CancelCommand = new RelayCommand(Cancel);
        TestAtolCommand = new AsyncRelayCommand(TestAtolAsync);
        TestOneCCommand = new AsyncRelayCommand(TestOneCAsync);
        BrowseObsidianFileCommand = new RelayCommand(BrowseObsidianFile);
        CheckForUpdatesCommand = new AsyncRelayCommand(CheckForUpdatesAsync);
        OpenUpdateCommand = new RelayCommand(OpenUpdate, () => HasUpdateAvailable);
        OpenRepositoryCommand = new RelayCommand(() =>
            ApplicationUpdateService.OpenUrl(ApplicationUpdateService.RepositoryUrl));

        if (ApplicationUpdateService.LastResult is { } lastUpdateResult)
            ApplyUpdateResult(lastUpdateResult);
    }

    public event Action<bool>? RequestClose;

    public ObservableCollection<ThemeOption> ThemeOptions { get; }
    public ObservableCollection<CashierInfo> Cashiers { get; }
    public ObservableCollection<ServiceProvider> Agents { get; }
    public IReadOnlyList<string> ServiceTypes { get; } = new[] { "Доставка", "Сборка" };
    public IReadOnlyList<VatRateOption> VatTypes { get; } = VatRateCatalog.All;

    public ICommand SelectSectionCommand { get; }
    public ICommand AddCashierCommand { get; }
    public ICommand DeleteCashierCommand { get; }
    public ICommand AddAgentCommand { get; }
    public ICommand DeleteAgentCommand { get; }
    public ICommand SaveCommand { get; }
    public ICommand CancelCommand { get; }
    public ICommand TestAtolCommand { get; }
    public ICommand TestOneCCommand { get; }
    public ICommand BrowseObsidianFileCommand { get; }
    public ICommand CheckForUpdatesCommand { get; }
    public ICommand OpenUpdateCommand { get; }
    public ICommand OpenRepositoryCommand { get; }

    public string Section
    {
        get => _section;
        set
        {
            if (!Set(ref _section, value)) return;
            OnPropertyChanged(nameof(ShowConnections));
            OnPropertyChanged(nameof(ShowAppearance));
            OnPropertyChanged(nameof(ShowCashiers));
            OnPropertyChanged(nameof(ShowAgents));
            OnPropertyChanged(nameof(ShowObsidian));
            OnPropertyChanged(nameof(ShowAbout));
        }
    }

    public bool ShowConnections => Section == "connections";
    public bool ShowAppearance  => Section == "appearance";
    public bool ShowCashiers    => Section == "cashiers";
    public bool ShowAgents      => Section == "agents";
    public bool ShowObsidian    => Section == "obsidian";
    public bool ShowAbout       => Section == "about";

    public string ApplicationName => "АТОЛ Чек-генератор";
    public string VersionText => $"Версия {ApplicationUpdateService.CurrentVersionText}";
    public string RepositoryText => "github.com/redservr-png/AtolGenerator";

    public ThemeOption? SelectedTheme
    {
        get => _selectedTheme;
        set
        {
            if (!Set(ref _selectedTheme, value) || value is null) return;
            ThemeService.ApplyTheme(value.Key);
        }
    }

    public CashierInfo? SelectedCashier
    {
        get => _selectedCashier;
        set => Set(ref _selectedCashier, value);
    }

    public CashierInfo? DefaultCashier
    {
        get => _defaultCashier;
        set => Set(ref _defaultCashier, value);
    }

    public ServiceProvider? SelectedAgent
    {
        get => _selectedAgent;
        set => Set(ref _selectedAgent, value);
    }

    public string AtolLogin { get; set; } = string.Empty;
    public string AtolPassword { get; set; } = string.Empty;
    public string AtolGroupCode { get; set; } = string.Empty;
    public string OneCServer { get; set; } = string.Empty;
    public string OneCDatabase { get; set; } = string.Empty;
    public string OneCUser { get; set; } = string.Empty;
    public string OneCPassword { get; set; } = string.Empty;
    public string ObsidianFilePath { get => _obsidianFilePath; set => Set(ref _obsidianFilePath, value); }
    public bool AutoValidateServiceNotes
    {
        get => _autoValidateServiceNotes;
        set => Set(ref _autoValidateServiceNotes, value);
    }

    public string AtolStatus { get => _atolStatus; set => Set(ref _atolStatus, value); }
    public string OneCStatus { get => _oneCStatus; set => Set(ref _oneCStatus, value); }
    public string UpdateStatus { get => _updateStatus; private set => Set(ref _updateStatus, value); }
    public bool HasUpdateAvailable => !string.IsNullOrWhiteSpace(_availableUpdateUrl);
    public string ValidationMessage
    {
        get => _validationMessage;
        set
        {
            Set(ref _validationMessage, value);
            OnPropertyChanged(nameof(HasValidationMessage));
        }
    }
    public bool HasValidationMessage => !string.IsNullOrWhiteSpace(ValidationMessage);

    public void CancelPreview() => ThemeService.ApplyTheme(_originalThemeKey);

    private void AddCashier()
    {
        var cashier = new CashierInfo
        {
            FullName = "Новый кассир",
            ShortName = "Фамилия И.О.",
            Position = "Должности",
            NameGenitive = "Фамилии Имени Отчества",
            Display = "Фамилия И.О. — должность",
        };
        Cashiers.Add(cashier);
        SelectedCashier = cashier;
        DefaultCashier ??= cashier;
    }

    private void DeleteCashier()
    {
        if (SelectedCashier is null) return;
        if (Cashiers.Count == 1)
        {
            ValidationMessage = "В программе должен остаться хотя бы один кассир.";
            return;
        }

        var removed = SelectedCashier;
        Cashiers.Remove(removed);
        if (ReferenceEquals(DefaultCashier, removed)) DefaultCashier = Cashiers[0];
        SelectedCashier = Cashiers.FirstOrDefault();
    }

    private void AddAgent()
    {
        var agent = new ServiceProvider("Доставка", "Новое подразделение", "Новый агент", "", "", "none");
        Agents.Add(agent);
        SelectedAgent = agent;
    }

    private void DeleteAgent()
    {
        if (SelectedAgent is null) return;
        Agents.Remove(SelectedAgent);
        SelectedAgent = Agents.FirstOrDefault();
    }

    private async Task TestAtolAsync()
    {
        AtolStatus = "Проверяем подключение...";
        var creds = new AtolCredentials
        {
            Login = AtolLogin.Trim(),
            Password = AtolPassword,
            GroupCode = AtolGroupCode.Trim(),
        };
        var (token, error) = await AtolApiService.GetTokenAsync(creds);
        AtolStatus = token is not null ? "Подключение установлено" : $"Ошибка: {error}";
    }

    private async Task TestOneCAsync()
    {
        OneCStatus = "Проверяем подключение...";
        var settings = BuildOneCSettings();
        OneCStatus = await Task.Run(() => OneCService.TestConnection(settings));
    }

    private void Save()
    {
        ValidationMessage = ValidateSettings();
        if (HasValidationMessage) return;

        foreach (var cashier in Cashiers)
        {
            cashier.FullName = cashier.FullName.Trim();
            cashier.ShortName = cashier.ShortName.Trim();
            cashier.Position = cashier.Position.Trim();
            cashier.NameGenitive = cashier.NameGenitive.Trim();
            cashier.Display = string.IsNullOrWhiteSpace(cashier.Display)
                ? cashier.ShortName
                : cashier.Display.Trim();
        }

        foreach (var agent in Agents)
        {
            agent.Service = agent.Service.Trim();
            agent.City = agent.City.Trim();
            agent.Name = agent.Name.Trim();
            agent.Inn = agent.Inn.Trim();
            agent.Phone = agent.Phone.Trim();
        }

        new AtolCredentials
        {
            Login = AtolLogin.Trim(),
            Password = AtolPassword,
            GroupCode = AtolGroupCode.Trim(),
        }.Save();
        BuildOneCSettings().Save();

        ApplicationSettingsStore.Save(new ApplicationSettings
        {
            ThemeKey = SelectedTheme?.Key ?? "light",
            SelectedCashierShortName = DefaultCashier!.ShortName,
            ObsidianFilePath = ObsidianFilePath.Trim(),
            LastAtolReportPath = ApplicationSettingsStore.Current.LastAtolReportPath,
            LastOfdReportPath = ApplicationSettingsStore.Current.LastOfdReportPath,
            AutoValidateServiceNotes = AutoValidateServiceNotes,
            Cashiers = Cashiers.Select(CloneCashier).ToList(),
            Agents = Agents.Select(CloneAgent).ToList(),
        });

        ThemeService.ApplyTheme(ApplicationSettingsStore.Current.ThemeKey);
        new ObsidianSettings { MdFilePath = ObsidianFilePath.Trim() }.Save();
        RequestClose?.Invoke(true);
    }

    private string ValidateSettings()
    {
        if (Cashiers.Count == 0 || DefaultCashier is null)
            return "Добавьте хотя бы одного кассира и выберите кассира по умолчанию.";

        var invalidCashier = Cashiers.FirstOrDefault(c =>
            string.IsNullOrWhiteSpace(c.FullName) || string.IsNullOrWhiteSpace(c.ShortName));
        if (invalidCashier is not null)
            return "У каждого кассира должны быть заполнены полное и короткое имя.";

        if (Agents.Count == 0)
            return "Добавьте хотя бы одно правило определения агента.";

        var invalidAgent = Agents.FirstOrDefault(a =>
            string.IsNullOrWhiteSpace(a.Service) || string.IsNullOrWhiteSpace(a.City) ||
            string.IsNullOrWhiteSpace(a.Name) || string.IsNullOrWhiteSpace(a.Inn) ||
            string.IsNullOrWhiteSpace(a.Phone));
        if (invalidAgent is not null)
            return "В каждом правиле агента заполните услугу, подразделение, агента, ИНН и телефон.";

        if (Agents.Any(a => !ServiceTypes.Contains(a.Service.Trim())))
            return "Для агента можно выбрать только услугу «Доставка» или «Сборка».";

        if (Agents.Any(a => !VatRateCatalog.IsKnown(a.VatType)))
            return "Выберите ставку НДС агента из справочника.";

        if (Agents.Any(a =>
                a.Inn.Trim().Length is not (10 or 12) ||
                a.Inn.Trim().Any(ch => !char.IsDigit(ch))))
            return "ИНН агента должен содержать 10 или 12 цифр.";

        if (Agents.Any(a =>
                !a.Phone.Trim().StartsWith('+') ||
                a.Phone.Trim().Skip(1).Any(ch => !char.IsDigit(ch))))
            return "Телефон агента укажите в формате +79991234567.";

        var duplicate = Agents
            .GroupBy(a => $"{a.Service.Trim().ToUpperInvariant()}|{a.City.Trim().ToUpperInvariant()}")
            .FirstOrDefault(g => g.Count() > 1);
        if (duplicate is not null)
            return "Для одной услуги и подразделения можно назначить только одного агента.";

        return string.Empty;
    }

    private void Cancel()
    {
        CancelPreview();
        RequestClose?.Invoke(false);
    }

    private void BrowseObsidianFile()
    {
        var dialog = new OpenFileDialog
        {
            Title = "Выберите рабочий файл Obsidian",
            Filter = "Markdown (*.md)|*.md|Все файлы (*.*)|*.*",
        };
        if (File.Exists(ObsidianFilePath))
            dialog.InitialDirectory = Path.GetDirectoryName(ObsidianFilePath);
        if (dialog.ShowDialog() == true)
            ObsidianFilePath = dialog.FileName;
    }

    private async Task CheckForUpdatesAsync()
    {
        UpdateStatus = "Проверяем GitHub Releases...";
        _availableUpdateUrl = string.Empty;
        OnPropertyChanged(nameof(HasUpdateAvailable));
        CommandManager.InvalidateRequerySuggested();

        ApplyUpdateResult(await ApplicationUpdateService.CheckForUpdateAsync());
    }

    private void ApplyUpdateResult(ApplicationUpdateResult result)
    {
        _availableUpdateUrl = string.Empty;
        if (!result.CheckSucceeded)
        {
            UpdateStatus = result.ErrorMessage;
            OnPropertyChanged(nameof(HasUpdateAvailable));
            CommandManager.InvalidateRequerySuggested();
            return;
        }

        if (!result.IsUpdateAvailable)
        {
            UpdateStatus = $"Установлена актуальная версия {result.CurrentVersion}";
            OnPropertyChanged(nameof(HasUpdateAvailable));
            CommandManager.InvalidateRequerySuggested();
            return;
        }

        _availableUpdateUrl = string.IsNullOrWhiteSpace(result.DownloadUrl)
            ? result.ReleaseUrl
            : result.DownloadUrl;
        UpdateStatus = $"Доступна версия {result.LatestVersion}";
        OnPropertyChanged(nameof(HasUpdateAvailable));
        CommandManager.InvalidateRequerySuggested();
    }

    private void OpenUpdate() => ApplicationUpdateService.OpenUrl(_availableUpdateUrl);

    private OneCConnectionSettings BuildOneCSettings() => new()
    {
        Server = OneCServer.Trim(),
        Database = OneCDatabase.Trim(),
        User = OneCUser.Trim(),
        Password = OneCPassword,
    };

    private static CashierInfo CloneCashier(CashierInfo source) => new()
    {
        FullName = source.FullName,
        ShortName = source.ShortName,
        Position = source.Position,
        NameGenitive = source.NameGenitive,
        Display = source.Display,
    };

    private static ServiceProvider CloneAgent(ServiceProvider source) => new(
        source.Service, source.City, source.Name, source.Inn, source.Phone, source.VatType);
}
