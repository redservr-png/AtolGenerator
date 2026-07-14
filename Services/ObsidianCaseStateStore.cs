using System.IO;
using System.Text.Encodings.Web;
using System.Text.Json;
using AtolGenerator.Models;

namespace AtolGenerator.Services;

public static class ObsidianCaseStateStore
{
    private static readonly JsonSerializerOptions JsonOptions = new()
    {
        WriteIndented = true,
        Encoder = JavaScriptEncoder.UnsafeRelaxedJsonEscaping,
    };

    private static string StatePath => Path.Combine(
        AppDomain.CurrentDomain.BaseDirectory, "obsidian_case_state.json");

    public static Dictionary<string, ObsidianCaseState> Load()
    {
        try
        {
            if (!File.Exists(StatePath))
                return new(StringComparer.OrdinalIgnoreCase);

            var json = File.ReadAllText(StatePath);
            var states = JsonSerializer.Deserialize<List<ObsidianCaseState>>(json) ?? new();
            foreach (var state in states)
                state.ExpectedChecks ??= new();
            return states
                .Where(x => !string.IsNullOrWhiteSpace(x.CaseId))
                .GroupBy(x => x.CaseId, StringComparer.OrdinalIgnoreCase)
                .ToDictionary(x => x.Key, x => x.Last(), StringComparer.OrdinalIgnoreCase);
        }
        catch
        {
            return new(StringComparer.OrdinalIgnoreCase);
        }
    }

    public static void Save(IEnumerable<ObsidianCaseState> states)
    {
        var json = JsonSerializer.Serialize(states.OrderBy(x => x.CaseId).ToList(), JsonOptions);
        File.WriteAllText(StatePath, json);
    }
}
