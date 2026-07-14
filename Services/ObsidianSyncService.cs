using System.IO;
using System.Text;
using System.Text.RegularExpressions;
using AtolGenerator.Models;

namespace AtolGenerator.Services;

public static class ObsidianSyncService
{
    private static readonly Regex TaskRegex = new(
        @"^(?<prefix>\s*-\s*\[(?<done>[ xX])\]\s*)(?<body>.*)$",
        RegexOptions.Compiled);

    private static readonly Regex IdRegex = new(
        @"<!--\s*atol-case:\s*(?<id>[a-zA-Z0-9-]+)\s*-->",
        RegexOptions.Compiled | RegexOptions.IgnoreCase);

    private static readonly Regex ScenarioRegex = new(
        @"<!--\s*atol-scenario:\s*(?<scenario>[a-zA-Z]+)\s*-->",
        RegexOptions.Compiled | RegexOptions.IgnoreCase);

    public static List<ObsidianCaseRecord> LoadAndEnsureIds(string path)
    {
        if (string.IsNullOrWhiteSpace(path) || !File.Exists(path))
            return new();

        var snapshot = ReadSnapshot(path);
        var changed = false;
        for (var index = 0; index < snapshot.Lines.Length; index++)
        {
            var match = TaskRegex.Match(snapshot.Lines[index]);
            if (!match.Success) continue;
            if (match.Groups["done"].Value.Equals("x", StringComparison.OrdinalIgnoreCase)) continue;
            if (IdRegex.IsMatch(match.Groups["body"].Value)) continue;

            snapshot.Lines[index] = snapshot.Lines[index].TrimEnd() +
                                    $" <!-- atol-case: {Guid.NewGuid():N} -->";
            changed = true;
        }

        if (changed) WriteSnapshot(path, snapshot);
        return ObsidianParserService.ParseDocument(string.Join(snapshot.NewLine, snapshot.Lines));
    }

    public static bool MarkCompleted(string path, string caseId)
    {
        return UpdateTaskLine(path, caseId, line =>
        {
            var match = TaskRegex.Match(line);
            if (!match.Success) return line;
            return match.Groups["done"].Value.Equals("x", StringComparison.OrdinalIgnoreCase)
                ? line
                : Regex.Replace(line, @"\[(?: |X)\]", "[x]", RegexOptions.IgnoreCase, TimeSpan.FromSeconds(1));
        });
    }

    public static bool UpdateProblem(string path, string caseId, string oldProblem, string newProblem)
    {
        return UpdateTaskLine(path, caseId, line =>
        {
            var idMatch = IdRegex.Match(line);
            var idComment = idMatch.Success ? idMatch.Value : string.Empty;
            var scenarioMatch = ScenarioRegex.Match(line);
            var scenarioComment = scenarioMatch.Success ? scenarioMatch.Value : string.Empty;
            var withoutId = ScenarioRegex.Replace(
                IdRegex.Replace(line, string.Empty), string.Empty).TrimEnd();
            if (!string.IsNullOrWhiteSpace(oldProblem))
            {
                var position = withoutId.LastIndexOf(oldProblem, StringComparison.Ordinal);
                if (position >= 0)
                    withoutId = withoutId[..position].TrimEnd() + " " + newProblem.Trim();
                else if (!string.IsNullOrWhiteSpace(newProblem))
                    withoutId += " " + newProblem.Trim();
            }
            else if (!string.IsNullOrWhiteSpace(newProblem))
            {
                withoutId += " " + newProblem.Trim();
            }

            return JoinTaskParts(withoutId, scenarioComment, idComment);
        });
    }

    public static bool UpdateScenario(
        string path,
        string caseId,
        CorrectionScenario scenario)
    {
        return UpdateTaskLine(path, caseId, line =>
        {
            var idMatch = IdRegex.Match(line);
            var idComment = idMatch.Success ? idMatch.Value : string.Empty;
            var body = ScenarioRegex.Replace(
                IdRegex.Replace(line, string.Empty), string.Empty).TrimEnd();
            var scenarioComment = $"<!-- atol-scenario: {scenario} -->";
            return JoinTaskParts(body, scenarioComment, idComment);
        });
    }

    private static string JoinTaskParts(params string[] parts)
    {
        var values = parts.Where(part => !string.IsNullOrWhiteSpace(part)).ToList();
        if (values.Count == 0) return string.Empty;
        var result = values[0].TrimEnd();
        foreach (var value in values.Skip(1))
            result += " " + value.Trim();
        return result;
    }

    private static bool UpdateTaskLine(string path, string caseId, Func<string, string> update)
    {
        if (!File.Exists(path) || string.IsNullOrWhiteSpace(caseId)) return false;
        var snapshot = ReadSnapshot(path);
        for (var index = 0; index < snapshot.Lines.Length; index++)
        {
            var id = IdRegex.Match(snapshot.Lines[index]);
            if (!id.Success || !id.Groups["id"].Value.Equals(caseId, StringComparison.OrdinalIgnoreCase)) continue;
            var updated = update(snapshot.Lines[index]);
            if (updated == snapshot.Lines[index]) return true;
            snapshot.Lines[index] = updated;
            WriteSnapshot(path, snapshot);
            return true;
        }
        return false;
    }

    private static TextSnapshot ReadSnapshot(string path)
    {
        var bytes = File.ReadAllBytes(path);
        var hasBom = bytes.Length >= 3 && bytes[0] == 0xEF && bytes[1] == 0xBB && bytes[2] == 0xBF;
        var text = new UTF8Encoding(hasBom, true).GetString(bytes);
        if (text.Length > 0 && text[0] == '\uFEFF') text = text[1..];
        var newLine = text.Contains("\r\n", StringComparison.Ordinal) ? "\r\n" : "\n";
        return new TextSnapshot(text.Replace("\r\n", "\n").Replace('\r', '\n').Split('\n'), newLine, hasBom);
    }

    private static void WriteSnapshot(string path, TextSnapshot snapshot)
    {
        var directory = Path.GetDirectoryName(Path.GetFullPath(path))
                        ?? throw new InvalidOperationException("Не удалось определить папку Obsidian-файла.");
        var tempPath = Path.Combine(directory, $".{Path.GetFileName(path)}.{Guid.NewGuid():N}.tmp");
        try
        {
            File.WriteAllText(tempPath, string.Join(snapshot.NewLine, snapshot.Lines),
                new UTF8Encoding(snapshot.HasBom));
            File.Move(tempPath, path, true);
        }
        finally
        {
            if (File.Exists(tempPath)) File.Delete(tempPath);
        }
    }

    private sealed record TextSnapshot(string[] Lines, string NewLine, bool HasBom);
}
