using System.IO;
using System.Text.Json;

namespace AtolGenerator.Models;

/// <summary>
/// Настройки для загрузки кейсов коррекций из Obsidian-файла.
/// Хранятся в obsidian_settings.json рядом с .exe — путь к .md файлу.
/// </summary>
public class ObsidianSettings
{
    public string MdFilePath { get; set; } = string.Empty;

    private static string SettingsPath => Path.Combine(
        AppDomain.CurrentDomain.BaseDirectory, "obsidian_settings.json");

    public void Save()
    {
        try
        {
            var json = JsonSerializer.Serialize(this,
                new JsonSerializerOptions
                {
                    WriteIndented = true,
                    Encoder = System.Text.Encodings.Web.JavaScriptEncoder.UnsafeRelaxedJsonEscaping,
                });
            File.WriteAllText(SettingsPath, json);
        }
        catch { /* проглатываем ошибки записи настроек */ }
    }

    public static ObsidianSettings Load()
    {
        try
        {
            if (!File.Exists(SettingsPath)) return new();
            var json = File.ReadAllText(SettingsPath);
            return JsonSerializer.Deserialize<ObsidianSettings>(json) ?? new();
        }
        catch { return new(); }
    }
}
