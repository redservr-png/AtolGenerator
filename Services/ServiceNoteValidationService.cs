using System.Globalization;
using System.IO;
using AtolGenerator.Helpers;
using DocumentFormat.OpenXml.Packaging;

namespace AtolGenerator.Services;

public static class ServiceNoteValidationService
{
    public static (bool IsValid, string Path, string Message) Validate(
        string documentNumber, double expectedAmount, string preferredPath)
    {
        var path = FindPath(documentNumber, preferredPath);
        if (string.IsNullOrWhiteSpace(path))
            return (false, string.Empty, "Служебная записка не найдена");

        try
        {
            using var document = WordprocessingDocument.Open(path, false);
            var text = document.MainDocumentPart?.Document?.Body?.InnerText ?? string.Empty;
            if (!string.IsNullOrWhiteSpace(documentNumber) &&
                !text.Contains(documentNumber, StringComparison.OrdinalIgnoreCase))
                return (false, path, "В служебке не найден номер документа");

            if (expectedAmount > 0 && !ContainsAmount(text, expectedAmount))
                return (false, path, "Сумма в служебке не совпадает с исправленной суммой");

            return (true, path, "Служебка проверена");
        }
        catch (Exception ex)
        {
            return (false, path, $"Не удалось проверить служебку: {ex.Message}");
        }
    }

    private static string FindPath(string documentNumber, string preferredPath)
    {
        if (!string.IsNullOrWhiteSpace(preferredPath) && File.Exists(preferredPath))
            return preferredPath;
        if (string.IsNullOrWhiteSpace(documentNumber) || !Directory.Exists(FileHelper.OutputDir))
            return string.Empty;

        var safeNumber = FileHelper.SafeFilename(documentNumber);
        return Directory.EnumerateFiles(FileHelper.OutputDir, $"*{safeNumber}*служебка.docx")
            .OrderByDescending(File.GetLastWriteTime)
            .FirstOrDefault() ?? string.Empty;
    }

    private static bool ContainsAmount(string text, double amount)
    {
        var variants = new[]
        {
            amount.ToString("F2", CultureInfo.GetCultureInfo("ru-RU")),
            amount.ToString("N2", CultureInfo.GetCultureInfo("ru-RU")),
            amount.ToString("F2", CultureInfo.InvariantCulture),
        };
        var compactText = text.Replace("\u00A0", string.Empty).Replace(" ", string.Empty);
        return variants.Any(value => compactText.Contains(
            value.Replace("\u00A0", string.Empty).Replace(" ", string.Empty),
            StringComparison.OrdinalIgnoreCase));
    }
}
