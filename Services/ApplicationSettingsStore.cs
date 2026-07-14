using System.IO;
using System.Text.Encodings.Web;
using System.Text.Json;
using AtolGenerator.Constants;
using AtolGenerator.Models;

namespace AtolGenerator.Services;

public static class ApplicationSettingsStore
{
    private static readonly JsonSerializerOptions JsonOptions = new()
    {
        WriteIndented = true,
        Encoder = JavaScriptEncoder.UnsafeRelaxedJsonEscaping,
    };

    private static string SettingsPath => Path.Combine(
        AppDomain.CurrentDomain.BaseDirectory, "app_settings.json");

    public static ApplicationSettings Current { get; private set; } = LoadCore();

    public static ApplicationSettings Reload()
    {
        Current = LoadCore();
        return Current;
    }

    public static void Save(ApplicationSettings settings)
    {
        Current = Normalize(settings);
        var json = JsonSerializer.Serialize(Current, JsonOptions);
        File.WriteAllText(SettingsPath, json);
    }

    private static ApplicationSettings LoadCore()
    {
        try
        {
            if (File.Exists(SettingsPath))
            {
                var json = File.ReadAllText(SettingsPath);
                var loaded = JsonSerializer.Deserialize<ApplicationSettings>(json);
                if (loaded is not null)
                    return Normalize(loaded);
            }
        }
        catch
        {
            // Повреждённые настройки не должны мешать запуску программы.
        }

        return Normalize(new ApplicationSettings());
    }

    private static ApplicationSettings Normalize(ApplicationSettings settings)
    {
        settings.ThemeKey = settings.ThemeKey is "dark" or "warm" ? settings.ThemeKey : "light";
        settings.ObsidianFilePath = settings.ObsidianFilePath?.Trim() ?? string.Empty;
        settings.Cashiers ??= new List<CashierInfo>();
        settings.Agents ??= new List<ServiceProvider>();

        settings.Cashiers = settings.Cashiers
            .Where(c => !string.IsNullOrWhiteSpace(c.FullName) || !string.IsNullOrWhiteSpace(c.ShortName))
            .ToList();
        if (settings.Cashiers.Count == 0)
            settings.Cashiers = AppConstants.CreateDefaultCashiers();

        settings.Agents = settings.Agents
            .Where(a => !string.IsNullOrWhiteSpace(a.Name) || !string.IsNullOrWhiteSpace(a.City))
            .ToList();
        if (settings.Agents.Count == 0)
            settings.Agents = AppConstants.CreateDefaultServiceProviders();

        if (string.IsNullOrWhiteSpace(settings.SelectedCashierShortName) ||
            settings.Cashiers.All(c => !string.Equals(
                c.ShortName, settings.SelectedCashierShortName, StringComparison.OrdinalIgnoreCase)))
            settings.SelectedCashierShortName = settings.Cashiers[0].ShortName;

        return settings;
    }
}
