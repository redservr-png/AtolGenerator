using System.Globalization;
using System.IO;
using System.Text;
using AtolGenerator.Models;

namespace AtolGenerator.Services;

public static class ReportReconciliationService
{
    private sealed class FiscalMatch
    {
        public DateTime RegisteredAt { get; init; }
        public long FiscalSign { get; init; }
        public long FiscalDocument { get; init; }
        public string Operation { get; init; } = string.Empty;
        public double Amount { get; init; }
        public string Source { get; init; } = string.Empty;
    }

    public static List<OneCExportRow> Build(
        IReadOnlyCollection<XmlReportCheck> xmlChecks,
        IReadOnlyCollection<AtolJournalReportRow> atolChecks,
        IReadOnlyCollection<OfdReportRow> ofdRows)
    {
        var atolByExternalId = atolChecks
            .Where(x => !string.IsNullOrWhiteSpace(x.ExternalId))
            .GroupBy(x => x.ExternalId, StringComparer.OrdinalIgnoreCase)
            .ToDictionary(g => g.Key, g => g.OrderByDescending(x => x.RegisteredAt).First(), StringComparer.OrdinalIgnoreCase);

        var ofdByRealization = ofdRows
            .Where(x => !string.IsNullOrWhiteSpace(x.AdditionalUserPropValue))
            .GroupBy(x => x.AdditionalUserPropValue, StringComparer.OrdinalIgnoreCase)
            .ToDictionary(g => g.Key, g => g.ToList(), StringComparer.OrdinalIgnoreCase);

        var ofdByFiscalSign = ofdRows
            .Where(x => x.FiscalSign.HasValue)
            .GroupBy(x => x.FiscalSign!.Value)
            .ToDictionary(g => g.Key, g => g.ToList());

        var usedOfdKeys = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

        var pairRealizations = xmlChecks
            .Where(IsSupportedOperation)
            .Where(x => !string.IsNullOrWhiteSpace(x.RealizationNumber))
            .GroupBy(x => x.RealizationNumber, StringComparer.OrdinalIgnoreCase)
            .Where(g => g.Any(x => x.Operation == "sell_refund") &&
                        g.Any(x => x.Operation is "sell" or "sell_correction"))
            .Select(g => g.Key)
            .ToHashSet(StringComparer.OrdinalIgnoreCase);

        var result = new List<OneCExportRow>();
        foreach (var xml in xmlChecks.OrderBy(x => x.Index))
        {
            if (!IsSupportedOperation(xml))
            {
                result.Add(ErrorRow(xml, "Тип чека не поддерживается для загрузки реализаций"));
                continue;
            }

            if (string.IsNullOrWhiteSpace(xml.RealizationNumber))
            {
                result.Add(ErrorRow(xml, "Не указан номер реализации"));
                continue;
            }

            if (string.IsNullOrWhiteSpace(xml.ExternalId))
            {
                result.Add(ErrorRow(xml, "В XML отсутствует External Id"));
                continue;
            }

            if (!TryResolveFiscalMatch(xml, atolByExternalId, ofdByRealization, ofdRows, usedOfdKeys, out var match, out var matchError))
            {
                result.Add(ErrorRow(xml, matchError));
                continue;
            }

            if (!string.IsNullOrWhiteSpace(match.Operation) &&
                !string.Equals(match.Operation, xml.Operation, StringComparison.OrdinalIgnoreCase))
            {
                result.Add(ErrorRow(xml, $"Тип операции не совпадает: XML {xml.Operation}, источник {match.Operation}"));
                continue;
            }

            var isPair = pairRealizations.Contains(xml.RealizationNumber);
            if (!isPair && xml.Operation is "sell_refund" or "sell")
            {
                result.Add(ErrorRow(xml, "Одиночный обычный чек не относится к загрузке коррекции реализации"));
                continue;
            }

            var writeMode = isPair ? "comment_only" : "update_fields";
            var comment = BuildComment(xml.Operation, match.RegisteredAt, match.FiscalSign, isPair);
            var ofdStatus = BuildOfdStatus(match, ofdRows.Count, ofdByFiscalSign);
            var status = string.Equals(match.Source, "taxcom", StringComparison.OrdinalIgnoreCase)
                ? "Готово · Такском"
                : "Готово";

            result.Add(new OneCExportRow
            {
                RealizationNumber = xml.RealizationNumber,
                CheckType = xml.Operation,
                WriteMode = writeMode,
                ExternalId = xml.ExternalId,
                FiscalSign = match.FiscalSign,
                FiscalDocument = match.FiscalDocument,
                RegisteredAt = match.RegisteredAt,
                Comment = comment,
                OfdStatus = ofdStatus,
                Status = status,
                IsReady = true,
            });
        }

        return result;
    }

    public static string? GetAtolCoverageWarning(
        IReadOnlyCollection<XmlReportCheck> xmlChecks,
        IReadOnlyCollection<AtolJournalReportRow> atolChecks)
    {
        if (xmlChecks.Count == 0) return null;

        var xmlDates = xmlChecks
            .Select(x => x.GeneratedAt)
            .Where(x => x.HasValue)
            .Select(x => x!.Value)
            .ToList();
        if (xmlDates.Count == 0) return null;

        var xmlMin = xmlDates.Min();
        var warnings = new List<string>();

        if (atolChecks.Count == 0)
        {
            warnings.Add("Журнал АТОЛ не загружен — для сопоставления используется архив Такском.");
            return string.Join(" ", warnings);
        }

        var atolDates = atolChecks
            .Select(x => x.RegisteredAt)
            .Where(x => x.HasValue)
            .Select(x => x!.Value)
            .ToList();
        if (atolDates.Count == 0)
        {
            warnings.Add("В CSV АТОЛ нет дат чеков. Журнал доступен только ~15 дней — обновите отчёт или используйте архив Такском.");
            return string.Join(" ", warnings);
        }

        var atolMin = atolDates.Min();
        var atolMax = atolDates.Max();
        if (xmlMin.Date < atolMin.Date)
        {
            warnings.Add(
                $"XML содержит чеки от {xmlMin:dd.MM.yyyy}, а журнал АТОЛ — с {atolMin:dd.MM.yyyy}. Обновите CSV АТОЛ или загрузите архив Такском.");
        }

        if ((atolMax - atolMin).TotalDays >= 13.5)
        {
            warnings.Add("Журнал АТОЛ охватывает примерно 15 дней — для более старых чеков нужен отчёт Такском.");
        }

        return warnings.Count > 0 ? string.Join(" ", warnings) : null;
    }

    public static void ExportOneCCsv(string path, IEnumerable<OneCExportRow> rows)
    {
        var readyRows = rows.Where(x => x.IsReady).ToList();
        if (readyRows.Count == 0)
            throw new InvalidOperationException("Нет готовых строк для экспорта в 1С.");

        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);
        var encoding = Encoding.GetEncoding(1251);
        using var writer = new StreamWriter(path, false, encoding);
        writer.WriteLine("НомерРеализации;ТипЧека;РежимЗаписи;ExternalId;ФПД;НомерФД;ДатаЧека;Комментарий");

        foreach (var row in readyRows
                     .OrderBy(x => x.RealizationNumber, StringComparer.OrdinalIgnoreCase)
                     .ThenBy(x => x.WriteMode == "comment_only" && x.CheckType == "sell_refund" ? 0 : 1))
        {
            writer.WriteLine(string.Join(";", new[]
            {
                Clean(row.RealizationNumber),
                Clean(row.CheckType),
                Clean(row.WriteMode),
                Clean(row.ExternalId),
                row.FiscalSign?.ToString(CultureInfo.InvariantCulture) ?? string.Empty,
                row.FiscalDocument?.ToString(CultureInfo.InvariantCulture) ?? string.Empty,
                row.RegisteredAt?.ToString("dd.MM.yyyy HH:mm:ss", CultureInfo.InvariantCulture) ?? string.Empty,
                Clean(row.Comment),
            }));
        }
    }

    private static bool TryResolveFiscalMatch(
        XmlReportCheck xml,
        IReadOnlyDictionary<string, AtolJournalReportRow> atolByExternalId,
        IReadOnlyDictionary<string, List<OfdReportRow>> ofdByRealization,
        IReadOnlyCollection<OfdReportRow> ofdRows,
        HashSet<string> usedOfdKeys,
        out FiscalMatch match,
        out string error)
    {
        match = null!;
        error = string.Empty;

        if (atolByExternalId.TryGetValue(xml.ExternalId, out var atol))
        {
            if (!atol.FiscalSign.HasValue || !atol.FiscalDocument.HasValue || !atol.RegisteredAt.HasValue)
            {
                error = "В отчёте АТОЛ не заполнены ФПД, ФД или дата чека";
                return false;
            }

            match = new FiscalMatch
            {
                RegisteredAt = atol.RegisteredAt.Value,
                FiscalSign = atol.FiscalSign.Value,
                FiscalDocument = atol.FiscalDocument.Value,
                Operation = atol.Operation,
                Amount = atol.Amount,
                Source = "atol",
            };
            return true;
        }

        if (ofdRows.Count == 0)
        {
            error = atolByExternalId.Count > 0
                ? "Чек не найден в отчёте АТОЛ"
                : "Чек не найден: загрузите CSV АТОЛ или архив Такском";
            return false;
        }

        var ofd = FindOfdMatch(xml, ofdByRealization, ofdRows, usedOfdKeys);
        if (ofd is null)
        {
            error = atolByExternalId.Count > 0
                ? "Чек не найден в отчёте АТОЛ и в архиве Такском"
                : "Чек не найден в архиве Такском";
            return false;
        }

        if (!ofd.FiscalSign.HasValue || !ofd.FiscalDocument.HasValue || !ofd.RegisteredAt.HasValue)
        {
            error = "В отчёте Такском не заполнены ФПД, ФД или дата чека";
            return false;
        }

        usedOfdKeys.Add(BuildOfdKey(ofd));
        match = new FiscalMatch
        {
            RegisteredAt = ofd.RegisteredAt.Value,
            FiscalSign = ofd.FiscalSign.Value,
            FiscalDocument = ofd.FiscalDocument.Value,
            Operation = ResolveOfdOperation(ofd),
            Amount = ofd.Amount,
            Source = "taxcom",
        };
        return true;
    }

    private static OfdReportRow? FindOfdMatch(
        XmlReportCheck xml,
        IReadOnlyDictionary<string, List<OfdReportRow>> ofdByRealization,
        IReadOnlyCollection<OfdReportRow> ofdRows,
        HashSet<string> usedOfdKeys)
    {
        var candidates = new List<OfdReportRow>();
        if (ofdByRealization.TryGetValue(xml.RealizationNumber, out var byRealization))
            candidates.AddRange(byRealization);
        else if (xml.Operation == "sell_correction")
            candidates.AddRange(ofdRows);

        foreach (var ofd in candidates
                     .OrderBy(x => x.RegisteredAt ?? DateTime.MaxValue)
                     .ThenBy(x => x.FiscalDocument ?? long.MaxValue))
        {
            if (usedOfdKeys.Contains(BuildOfdKey(ofd))) continue;
            if (!AmountMatches(ofd.Amount, xml.Amount)) continue;
            if (!OperationMatches(ofd, xml.Operation)) continue;

            if (xml.Operation == "sell_correction")
            {
                if (!ofd.Document.Contains("коррекции", StringComparison.OrdinalIgnoreCase))
                    continue;
                if (!string.IsNullOrWhiteSpace(xml.OriginalFiscalSign) &&
                    !string.IsNullOrWhiteSpace(ofd.AdditionalCheckProps) &&
                    !string.Equals(
                        NormalizeDigits(xml.OriginalFiscalSign),
                        NormalizeDigits(ofd.AdditionalCheckProps),
                        StringComparison.OrdinalIgnoreCase))
                    continue;
            }
            else if (ofdByRealization.ContainsKey(xml.RealizationNumber) &&
                     !string.Equals(ofd.AdditionalUserPropValue, xml.RealizationNumber, StringComparison.OrdinalIgnoreCase))
            {
                continue;
            }

            return ofd;
        }

        return null;
    }

    private static bool IsSupportedOperation(XmlReportCheck check) =>
        check.Operation is "sell" or "sell_correction" or "sell_refund";

    private static string BuildComment(string operation, DateTime registeredAt, long fiscalSign, bool isPair)
    {
        var date = registeredAt.ToString("dd.MM.yyyy", CultureInfo.InvariantCulture);
        return operation switch
        {
            "sell_refund" when isPair =>
                $"{date} Пробит исправительный чек \"Возврат прихода\" ФП: {fiscalSign}",
            "sell" when isPair =>
                $"{date} Пробит исправительный чек \"Приход\" ФП: {fiscalSign}",
            "sell_correction" =>
                $"{date} Пробит чек коррекции \"Приход\" ФП: {fiscalSign}",
            _ => string.Empty,
        };
    }

    private static string BuildOfdStatus(
        FiscalMatch match,
        int ofdCount,
        IReadOnlyDictionary<long, List<OfdReportRow>> ofdByFiscalSign)
    {
        if (string.Equals(match.Source, "taxcom", StringComparison.OrdinalIgnoreCase))
            return "Источник: Такском";

        if (ofdCount == 0) return "ОФД не загружен";
        if (!ofdByFiscalSign.TryGetValue(match.FiscalSign, out var candidates))
            return "Не найден в ОФД";

        var exact = candidates.Any(x =>
            x.FiscalDocument == match.FiscalDocument &&
            Math.Abs(Math.Abs(x.Amount) - Math.Abs(match.Amount)) < 0.01);
        return exact ? "Проверено ОФД" : "Расхождение с ОФД";
    }

    private static bool OperationMatches(OfdReportRow ofd, string xmlOperation)
    {
        var ofdOperation = ResolveOfdOperation(ofd);
        return string.Equals(ofdOperation, xmlOperation, StringComparison.OrdinalIgnoreCase);
    }

    private static string ResolveOfdOperation(OfdReportRow ofd)
    {
        var operation = ofd.Operation.Trim().ToLowerInvariant();
        if (operation.Length > 0) return operation;

        if (ofd.Document.Contains("коррекции", StringComparison.OrdinalIgnoreCase))
            return "sell_correction";
        if (ofd.Document.Contains("возврат", StringComparison.OrdinalIgnoreCase))
            return "sell_refund";
        return "sell";
    }

    private static bool AmountMatches(double left, double right) =>
        Math.Abs(Math.Abs(left) - Math.Abs(right)) < 0.01;

    private static string BuildOfdKey(OfdReportRow row) =>
        $"{row.FiscalSign}:{row.FiscalDocument}:{row.RegisteredAt:O}";

    private static string NormalizeDigits(string value) =>
        new(value.Where(char.IsDigit).ToArray());

    private static OneCExportRow ErrorRow(XmlReportCheck xml, string status) => new()
    {
        RealizationNumber = xml.RealizationNumber,
        CheckType = xml.Operation,
        ExternalId = xml.ExternalId,
        Status = status,
        OfdStatus = string.Empty,
        IsReady = false,
    };

    private static string Clean(string value) =>
        value.Replace(';', ',').Replace('\r', ' ').Replace('\n', ' ').Trim();
}
