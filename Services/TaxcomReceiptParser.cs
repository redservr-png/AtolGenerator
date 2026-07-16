using System.Globalization;
using System.Text.RegularExpressions;
using AtolGenerator.Models;

namespace AtolGenerator.Services;

public static partial class TaxcomReceiptParser
{
    public static bool TryParse(string text, long expectedFiscalSign, out OfdReportRow row, out string error)
    {
        row = new OfdReportRow();
        error = string.Empty;
        var normalized = Normalize(text);
        if (normalized.Length == 0)
        {
            error = "Карточка чека не содержит текста";
            return false;
        }

        var fiscalSign = ReadLong(FiscalSignRegex(), normalized);
        if (fiscalSign != expectedFiscalSign)
        {
            error = fiscalSign.HasValue
                ? $"Открыт другой чек: ФП {fiscalSign.Value}"
                : "В карточке чека не найден ФП";
            return false;
        }

        var registeredAt = ReadDate(normalized);
        var fiscalDocument = ReadLong(FiscalDocumentRegex(), normalized);
        var amount = ReadAmount(normalized);
        var operation = ReadOperation(normalized);
        if (!registeredAt.HasValue || !fiscalDocument.HasValue || amount <= 0 || operation.Length == 0)
        {
            error = "Не удалось прочитать дату, операцию, сумму или ФД из карточки чека";
            return false;
        }

        var checkNumber = ReadText(CheckNumberRegex(), normalized);
        row = new OfdReportRow
        {
            RegisteredAt = registeredAt,
            Document = checkNumber.Length > 0 ? $"Кассовый чек № {checkNumber}" : "Кассовый чек",
            Operation = operation,
            CalculationMethod = ReadCalculationMethod(normalized),
            Amount = amount,
            FiscalDocument = fiscalDocument,
            FiscalSign = fiscalSign,
            FiscalDriveNumber = ReadText(FiscalDriveRegex(), normalized),
            KktRegistrationNumber = ReadText(KktRegex(), normalized),
            ReceiptUrl = "https://lk-ofd.taxcom.ru/#receipts",
            SourceFile = "Онлайн-поиск Такском",
        };
        return true;
    }

    private static string Normalize(string value) =>
        value.Replace('\u00A0', ' ').Replace("\r", string.Empty).Trim();

    private static DateTime? ReadDate(string text)
    {
        var match = ReceiptDateRegex().Match(text);
        if (!match.Success) return null;
        var value = $"{match.Groups[1].Value} {match.Groups[2].Value}";
        string[] formats =
        {
            "dd.MM.yy HH:mm", "dd.MM.yy HH:mm:ss",
            "dd.MM.yyyy HH:mm", "dd.MM.yyyy HH:mm:ss",
        };
        return DateTime.TryParseExact(value, formats, CultureInfo.InvariantCulture,
            DateTimeStyles.AllowWhiteSpaces, out var result)
            ? result
            : null;
    }

    private static double ReadAmount(string text)
    {
        var match = TotalRegex().Match(text);
        if (!match.Success) return 0;
        var value = match.Groups[1].Value.Replace(" ", string.Empty).Replace(',', '.');
        return double.TryParse(value, NumberStyles.Any, CultureInfo.InvariantCulture, out var result)
            ? Math.Abs(result)
            : 0;
    }

    private static long? ReadLong(Regex regex, string text)
    {
        var value = ReadText(regex, text).Replace(" ", string.Empty);
        return long.TryParse(value, NumberStyles.Integer, CultureInfo.InvariantCulture, out var result)
            ? result
            : null;
    }

    private static string ReadText(Regex regex, string text)
    {
        var match = regex.Match(text);
        return match.Success ? match.Groups[1].Value.Trim() : string.Empty;
    }

    private static string ReadOperation(string text)
    {
        if (text.Contains("ВОЗВРАТ ПРИХОДА", StringComparison.OrdinalIgnoreCase)) return "Возврат прихода";
        if (text.Contains("КОРРЕКЦИЯ ПРИХОДА", StringComparison.OrdinalIgnoreCase)) return "Коррекция прихода";
        if (text.Contains("КОРРЕКЦИЯ РАСХОДА", StringComparison.OrdinalIgnoreCase)) return "Коррекция расхода";
        if (OperationLineRegex().IsMatch(text)) return "Приход";
        return string.Empty;
    }

    private static string ReadCalculationMethod(string text)
    {
        var match = CalculationMethodRegex().Match(text);
        if (!match.Success) return string.Empty;
        return match.Groups[1].Value.Trim() switch
        {
            "ПРЕДОПЛАТА 100%" => "Предоплата 100%",
            "ПРЕДОПЛАТА" => "Частичная предоплата",
            "АВАНС" => "Аванс расчет",
            "ПОЛНЫЙ РАСЧЕТ" => "Полный расчет",
            "ЧАСТИЧНЫЙ РАСЧЕТ И КРЕДИТ" => "Частичный расчет и кредит",
            "ПЕРЕДАЧА В КРЕДИТ" => "Передача в кредит",
            "ОПЛАТА КРЕДИТА" => "Оплата кредита",
            var value => value,
        };
    }

    [GeneratedRegex(@"КАССОВЫЙ\s+ЧЕК\s*№\s*:?\s*(\d+)", RegexOptions.IgnoreCase)]
    private static partial Regex CheckNumberRegex();

    [GeneratedRegex(@"(\d{2}\.\d{2}\.(?:\d{2}|\d{4}))\s+(\d{2}:\d{2}(?::\d{2})?)")]
    private static partial Regex ReceiptDateRegex();

    [GeneratedRegex(@"ИТОГО\s*:?\s*([\d\s]+[.,]\d{2})", RegexOptions.IgnoreCase)]
    private static partial Regex TotalRegex();

    [GeneratedRegex(@"№\s*ККТ\s*:?\s*(\d+)", RegexOptions.IgnoreCase)]
    private static partial Regex KktRegex();

    [GeneratedRegex(@"№\s*ФН\s*:?\s*(\d+)", RegexOptions.IgnoreCase)]
    private static partial Regex FiscalDriveRegex();

    [GeneratedRegex(@"№\s*ФД\s*:?\s*(\d+)", RegexOptions.IgnoreCase)]
    private static partial Regex FiscalDocumentRegex();

    [GeneratedRegex(@"(?:^|\n)\s*ФП(?:Д)?\s*:?\s*(\d+)", RegexOptions.IgnoreCase | RegexOptions.Multiline)]
    private static partial Regex FiscalSignRegex();

    [GeneratedRegex(@"(?:^|\n)\s*ПРИХОД\s*(?:\n|$)", RegexOptions.IgnoreCase | RegexOptions.Multiline)]
    private static partial Regex OperationLineRegex();

    [GeneratedRegex(@"Признак\s+способа\s+расчета\s*\n\s*([^\r\n]+)", RegexOptions.IgnoreCase)]
    private static partial Regex CalculationMethodRegex();
}
