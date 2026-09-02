using System.Diagnostics;
using System.IO;
using System.Text.RegularExpressions;
using AtolGenerator.Models;

namespace AtolGenerator.Helpers;

public static class FileHelper
{
    private static readonly Regex SafePattern = new(@"[^\w\-]", RegexOptions.Compiled);

    public static string SafeFilename(string? s) => SafePattern.Replace(s ?? string.Empty, "_");

    public static void OpenFolder(string path)
    {
        Directory.CreateDirectory(path);
        Process.Start(new ProcessStartInfo
        {
            FileName        = path,
            UseShellExecute = true
        });
    }

    public static void RevealInExplorer(string path)
    {
        if (File.Exists(path))
        {
            Process.Start(new ProcessStartInfo
            {
                FileName = "explorer.exe",
                Arguments = $"/select,\"{path}\"",
                UseShellExecute = true,
            });
            return;
        }

        var directory = Directory.Exists(path) ? path : Path.GetDirectoryName(path);
        if (!string.IsNullOrWhiteSpace(directory))
            OpenFolder(directory);
    }

    public static string OutputDir =>
        Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "atol_output");

    public static string TaxcomReportDir =>
        Path.Combine(OutputDir, "reports", "taxcom");

    public static string TaxcomReceiptCachePath =>
        Path.Combine(TaxcomReportDir, "online_receipts.json");

    public static string AtolReportDir =>
        Path.Combine(OutputDir, "reports", "atol");

    public static string PendingXmlRoot(string? outputDir = null) =>
        Path.Combine(outputDir ?? OutputDir, "xml");

    public static string ServiceNotesRoot(string? outputDir = null) =>
        Path.Combine(outputDir ?? OutputDir, "служебные_записки");

    public static string ProcessedXmlRoot(string? outputDir = null) =>
        Path.Combine(outputDir ?? OutputDir, "xml_в_1с");

    public static string YearMonthDirectory(string root, DateTime date)
    {
        var dir = Path.Combine(root, date.ToString("yyyy"), date.ToString("MM"));
        Directory.CreateDirectory(dir);
        return dir;
    }

    public static string GetPendingXmlDirectory(DateTime? date = null, string? outputDir = null) =>
        YearMonthDirectory(PendingXmlRoot(outputDir), date ?? DateTime.Now);

    public static string GetPendingXmlPath(string fileName, DateTime? date = null, string? outputDir = null) =>
        Path.Combine(GetPendingXmlDirectory(date, outputDir), fileName);

    public static string GetServiceNotePath(string fileName, DateTime? date = null, string? outputDir = null)
    {
        var dir = YearMonthDirectory(ServiceNotesRoot(outputDir), date ?? DateTime.Now);
        return Path.Combine(dir, fileName);
    }

    public static string GetProcessedXmlDirectory(DateTime? date = null, string? outputDir = null) =>
        YearMonthDirectory(ProcessedXmlRoot(outputDir), date ?? DateTime.Now);

    public static bool IsUnderDirectory(string path, string directory)
    {
        if (string.IsNullOrWhiteSpace(path) || string.IsNullOrWhiteSpace(directory))
            return false;
        var fullPath = Path.GetFullPath(path);
        var fullDir = Path.GetFullPath(directory).TrimEnd(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar)
                      + Path.DirectorySeparatorChar;
        return fullPath.StartsWith(fullDir, StringComparison.OrdinalIgnoreCase);
    }

    public static string MoveFileToDirectory(string sourcePath, string destDir)
    {
        Directory.CreateDirectory(destDir);
        var dest = Path.Combine(destDir, Path.GetFileName(sourcePath));
        if (string.Equals(Path.GetFullPath(sourcePath), Path.GetFullPath(dest), StringComparison.OrdinalIgnoreCase))
            return dest;
        if (File.Exists(dest))
        {
            var name = Path.GetFileNameWithoutExtension(sourcePath);
            var ext = Path.GetExtension(sourcePath);
            dest = Path.Combine(destDir, $"{name}_{DateTime.Now:HHmmss}{ext}");
        }
        File.Move(sourcePath, dest);
        return dest;
    }

    public static int ArchiveProcessedXml(
        IReadOnlyCollection<XmlReportCheck> xmlChecks,
        IReadOnlyCollection<OneCExportRow> readyRows,
        IReadOnlyCollection<string> processedNumbers)
    {
        var processed = new HashSet<string>(processedNumbers, StringComparer.OrdinalIgnoreCase);
        if (processed.Count == 0) return 0;

        var moved = 0;
        var files = readyRows
            .Where(x => !string.IsNullOrWhiteSpace(x.SourcePath))
            .Select(x => x.SourcePath)
            .Distinct(StringComparer.OrdinalIgnoreCase);

        foreach (var file in files)
        {
            if (!File.Exists(file)) continue;
            if (IsUnderDirectory(file, ProcessedXmlRoot())) continue;

            var checksInFile = xmlChecks
                .Where(x => string.Equals(x.SourcePath, file, StringComparison.OrdinalIgnoreCase))
                .ToList();
            if (checksInFile.Count == 0) continue;

            var readyFromFile = readyRows
                .Where(x => string.Equals(x.SourcePath, file, StringComparison.OrdinalIgnoreCase))
                .ToList();
            if (readyFromFile.Count != checksInFile.Count) continue;
            if (readyFromFile.Any(x => !processed.Contains(x.RealizationNumber))) continue;

            var when = readyFromFile
                .Select(x => x.RegisteredAt)
                .Where(x => x.HasValue)
                .Select(x => x!.Value)
                .DefaultIfEmpty(DateTime.Now)
                .Max();
            try
            {
                MoveFileToDirectory(file, GetProcessedXmlDirectory(when));
                moved++;
            }
            catch
            {
                // Перенос не должен откатывать уже записанные в 1С данные.
            }
        }

        return moved;
    }
}
