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
        public string Search { get; init; } = string.Empty;
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
            .Select(x => (Key: CanonRealization(x.AdditionalUserPropValue), Row: x))
            .Where(x => x.Key.Length > 0)
            .GroupBy(x => x.Key, StringComparer.OrdinalIgnoreCase)
            .ToDictionary(
                g => g.Key,
                g => g.Select(x => x.Row).ToList(),
                StringComparer.OrdinalIgnoreCase);

        var ofdByFiscalSign = ofdRows
            .Where(x => x.FiscalSign.HasValue)
            .GroupBy(x => x.FiscalSign!.Value)
            .ToDictionary(g => g.Key, g => g.ToList());

        var usedOfdKeys = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        var pairOriginalFiscalSign = xmlChecks
            .Where(x => x.Operation == "sell_refund" &&
                        !string.IsNullOrWhiteSpace(x.RealizationNumber) &&
                        !string.IsNullOrWhiteSpace(x.OriginalFiscalSign))
            .GroupBy(x => x.RealizationNumber, StringComparer.OrdinalIgnoreCase)
            .ToDictionary(g => g.Key, g => g.First().OriginalFiscalSign, StringComparer.OrdinalIgnoreCase);

        var uniqueDocumentDateByNumber = xmlChecks
            .Select(x => (x.RealizationNumber, Date: ParseDocumentDate(x.BaseDate)))
            .Where(x => !string.IsNullOrWhiteSpace(x.RealizationNumber) && x.Date.HasValue)
            .GroupBy(x => x.RealizationNumber, StringComparer.OrdinalIgnoreCase)
            .Select(g => (
                Number: g.Key,
                Dates: g.Select(x => x.Date!.Value.Date).Distinct().ToList()))
            .Where(x => x.Dates.Count == 1)
            .ToDictionary(x => x.Number, x => x.Dates[0], StringComparer.OrdinalIgnoreCase);

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

            var originalFiscalSign = xml.OriginalFiscalSign;
            if (string.IsNullOrWhiteSpace(originalFiscalSign) &&
                xml.Operation == "sell_correction" &&
                pairOriginalFiscalSign.TryGetValue(xml.RealizationNumber, out var siblingFiscalSign))
                originalFiscalSign = siblingFiscalSign;

            if (!TryResolveFiscalMatch(
                    xml, originalFiscalSign, atolChecks, atolByExternalId, ofdByRealization, ofdRows, usedOfdKeys,
                    out var match, out var matchError))
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
            var status = $"Готово · {match.Search}";
            var documentDate = ParseDocumentDate(xml.BaseDate)
                ?? (uniqueDocumentDateByNumber.TryGetValue(xml.RealizationNumber, out var sharedDate)
                    ? sharedDate
                    : null);

            result.Add(new OneCExportRow
            {
                RealizationNumber = xml.RealizationNumber,
                DocumentDate = documentDate,
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
                SourcePath = xml.SourcePath,
            });
        }

        return CollapseRetryDuplicates(result);
    }

    public static string? GetPairXmlWarning(IReadOnlyCollection<XmlReportCheck> xmlChecks)
    {
        var hasRefund = xmlChecks.Any(x => x.Operation == "sell_refund");
        var hasCorrection = xmlChecks.Any(x => x.Operation is "sell" or "sell_correction");
        if (hasCorrection && !hasRefund)
            return "Загружена только коррекция. Для пары исправления выберите оба XML (возврат и коррекцию) — иначе запись пойдёт как одиночная коррекция.";
        if (hasRefund && !hasCorrection)
            return "Загружен только возврат. Для пары исправления выберите оба XML: возврат и коррекцию.";
        return null;
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
        writer.WriteLine("НомерРеализации;ДатаРеализации;ТипЧека;РежимЗаписи;ExternalId;ФПД;НомерФД;ДатаЧека;Комментарий");

        foreach (var row in readyRows
                     .OrderBy(x => x.RealizationNumber, StringComparer.OrdinalIgnoreCase)
                     .ThenBy(x => x.WriteMode == "comment_only" && x.CheckType == "sell_refund" ? 0 : 1))
        {
            writer.WriteLine(string.Join(";", new[]
            {
                Clean(row.RealizationNumber),
                row.DocumentDate?.ToString("dd.MM.yyyy", CultureInfo.InvariantCulture) ?? string.Empty,
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
        string originalFiscalSign,
        IReadOnlyCollection<AtolJournalReportRow> atolChecks,
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

            match = FromAtol(atol, "АТОЛ по External Id");
            return true;
        }

        var atolByNumber = FindAtolByNumber(xml, atolChecks);
        if (atolByNumber is not null)
        {
            match = FromAtol(atolByNumber, "АТОЛ по номеру и сумме");
            return true;
        }

        if (ofdRows.Count == 0)
        {
            error = atolByExternalId.Count > 0
                ? "Чек не найден в отчёте АТОЛ"
                : "Чек не найден: загрузите CSV АТОЛ или архив Такском";
            return false;
        }

        var ofd = FindOfdMatch(xml, originalFiscalSign, ofdByRealization, ofdRows, usedOfdKeys, out var search);
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
        var operation = ResolveOfdOperation(ofd);
        if (string.IsNullOrWhiteSpace(operation) ||
            search.StartsWith("Такском: один", StringComparison.Ordinal))
            operation = xml.Operation;
        match = new FiscalMatch
        {
            RegisteredAt = ofd.RegisteredAt.Value,
            FiscalSign = ofd.FiscalSign.Value,
            FiscalDocument = ofd.FiscalDocument.Value,
            Operation = operation,
            Amount = ofd.Amount,
            Source = "taxcom",
            Search = search,
        };
        return true;
    }

    private static FiscalMatch FromAtol(AtolJournalReportRow atol, string search) => new()
    {
        RegisteredAt = atol.RegisteredAt!.Value,
        FiscalSign = atol.FiscalSign!.Value,
        FiscalDocument = atol.FiscalDocument!.Value,
        Operation = atol.Operation,
        Amount = atol.Amount,
        Source = "atol",
        Search = search,
    };

    private static AtolJournalReportRow? FindAtolByNumber(
        XmlReportCheck xml,
        IReadOnlyCollection<AtolJournalReportRow> atolChecks)
    {
        var number = CanonRealization(xml.RealizationNumber);
        if (number.Length == 0) return null;

        return atolChecks
            .Where(x => x.FiscalSign.HasValue && x.FiscalDocument.HasValue && x.RegisteredAt.HasValue)
            .Where(x => AmountMatches(x.Amount, xml.Amount))
            .Where(x => string.Equals(x.Operation, xml.Operation, StringComparison.OrdinalIgnoreCase))
            .Where(x => CanonRealization(x.BaseNumber) == number ||
                        TextHasRealization(x.ExternalId, number) ||
                        TextHasRealization(x.IncomingJson, number))
            .OrderBy(x => DateDistance(x.RegisteredAt, xml.GeneratedAt))
            .FirstOrDefault();
    }

    private static OfdReportRow? FindOfdMatch(
        XmlReportCheck xml,
        string originalFiscalSign,
        IReadOnlyDictionary<string, List<OfdReportRow>> ofdByRealization,
        IReadOnlyCollection<OfdReportRow> ofdRows,
        HashSet<string> usedOfdKeys,
        out string search)
    {
        search = string.Empty;
        var foundBy = string.Empty;
        var unused = ofdRows.Where(row => !usedOfdKeys.Contains(BuildOfdKey(row))).ToList();
        if (unused.Count == 0) return null;

        var number = CanonRealization(xml.RealizationNumber);
        OfdReportRow? Take(IEnumerable<OfdReportRow> rows, string label)
        {
            var hit = rows
                .OrderBy(row => DateDistance(row.RegisteredAt, xml.GeneratedAt))
                .FirstOrDefault();
            if (hit is null) return null;
            foundBy = label;
            return hit;
        }

        var byNumber = new List<OfdReportRow>();
        if (number.Length > 0 &&
            ofdByRealization.TryGetValue(number, out var indexed))
            byNumber.AddRange(indexed.Where(row => unused.Contains(row)));

        var hitByNumber = Take(
            byNumber.Where(row => OperationMatches(row, xml.Operation) && AmountMatches(row.Amount, xml.Amount)),
            "Такском по номеру и сумме");
        if (hitByNumber is not null)
        {
            search = foundBy;
            return hitByNumber;
        }

        if (!string.IsNullOrWhiteSpace(originalFiscalSign))
        {
            var byFiscalSign = Take(
                unused.Where(row =>
                    OperationMatches(row, xml.Operation) &&
                    SameDigits(row.AdditionalCheckProps, originalFiscalSign)),
                "Такском по ФП исходного чека");
            if (byFiscalSign is not null)
            {
                search = foundBy;
                return byFiscalSign;
            }
        }

        if (number.Length > 0)
        {
            var byText = Take(
                unused.Where(row =>
                    OperationMatches(row, xml.Operation) &&
                    AmountMatches(row.Amount, xml.Amount) &&
                    RowMentionsRealization(row, number)),
                "Такском по номеру в тексте чека");
            if (byText is not null)
            {
                search = foundBy;
                return byText;
            }
        }

        if (xml.GeneratedAt.HasValue)
        {
            var byDate = Take(
                unused.Where(row =>
                    OperationMatches(row, xml.Operation) &&
                    AmountMatches(row.Amount, xml.Amount) &&
                    row.RegisteredAt.HasValue &&
                    Math.Abs((row.RegisteredAt.Value - xml.GeneratedAt.Value).TotalDays) <= 7),
                "Такском по дате и сумме");
            if (byDate is not null)
            {
                search = foundBy;
                return byDate;
            }
        }

        var sameOperation = unused
            .Where(row => OperationMatches(row, xml.Operation) && AmountMatches(row.Amount, xml.Amount))
            .ToList();
        if (sameOperation.Count == 1)
        {
            search = "Такском: одна сумма и операция";
            return sameOperation[0];
        }

        if (number.Length > 0)
        {
            var sameNumber = unused
                .Where(row => AmountMatches(row.Amount, xml.Amount) &&
                              (CanonRealization(row.AdditionalUserPropValue) == number ||
                               RowMentionsRealization(row, number)))
                .ToList();
            if (sameNumber.Count == 1)
            {
                search = "Такском: один номер и сумма";
                return sameNumber[0];
            }
        }

        if (!string.IsNullOrWhiteSpace(originalFiscalSign))
        {
            var sameFiscalSign = unused
                .Where(row => AmountMatches(row.Amount, xml.Amount) &&
                              SameDigits(row.AdditionalCheckProps, originalFiscalSign))
                .ToList();
            if (sameFiscalSign.Count == 1)
            {
                search = "Такском: один ФП и сумма";
                return sameFiscalSign[0];
            }
        }

        return null;
    }

    private static bool RowMentionsRealization(OfdReportRow row, string number) =>
        TextHasRealization(row.Document, number) ||
        TextHasRealization(row.AdditionalUserPropValue, number) ||
        TextHasRealization(row.AdditionalUserPropName, number) ||
        TextHasRealization(row.AdditionalCheckProps, number) ||
        TextHasRealization(row.CalculationMethod, number);

    private static bool TextHasRealization(string? text, string number)
    {
        if (string.IsNullOrWhiteSpace(text) || number.Length == 0) return false;
        return CanonRealization(text).Contains(number, StringComparison.OrdinalIgnoreCase) ||
               text.Contains(number, StringComparison.OrdinalIgnoreCase);
    }

    private static string CanonRealization(string? value)
    {
        if (string.IsNullOrWhiteSpace(value)) return string.Empty;
        var text = ReportImportService.NormalizeRealizationNumber(value)
            .Replace(" ", string.Empty)
            .Replace('T', 'т')
            .Replace('t', 'т');
        return text;
    }

    private static bool SameDigits(string? left, string? right)
    {
        var a = NormalizeDigits(left ?? string.Empty);
        var b = NormalizeDigits(right ?? string.Empty);
        return a.Length > 0 && a == b;
    }

    private static double DateDistance(DateTime? registeredAt, DateTime? generatedAt)
    {
        if (!registeredAt.HasValue || !generatedAt.HasValue) return 10_000;
        return Math.Abs((registeredAt.Value - generatedAt.Value).TotalHours);
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
        if (operation is "sell" or "sell_refund" or "sell_correction" or "buy" or "buy_refund" or "buy_correction")
            return operation;

        var text = $"{operation} {ofd.Document}".ToLowerInvariant();
        var isCorrection = text.Contains("коррек", StringComparison.Ordinal);
        var isRefund = text.Contains("возврат", StringComparison.Ordinal);
        var isExpense = text.Contains("расход", StringComparison.Ordinal) &&
                        !text.Contains("приход", StringComparison.Ordinal);

        if (isCorrection && isExpense) return "buy_correction";
        if (isCorrection && isRefund) return "buy_refund";
        if (isCorrection) return "sell_correction";
        if (isRefund && isExpense) return "buy_refund";
        if (isRefund) return "sell_refund";
        if (isExpense) return "buy";
        if (text.Contains("приход", StringComparison.Ordinal) || operation == "sell") return "sell";
        return operation;
    }

    private static bool AmountMatches(double left, double right) =>
        Math.Abs(Math.Abs(left) - Math.Abs(right)) < 0.01;

    private static string BuildOfdKey(OfdReportRow row) =>
        $"{row.FiscalSign}:{row.FiscalDocument}:{row.RegisteredAt:O}";

    private static string NormalizeDigits(string value) =>
        new(value.Where(char.IsDigit).ToArray());

    private static List<OneCExportRow> CollapseRetryDuplicates(List<OneCExportRow> rows)
    {
        if (rows.Count <= 1) return rows;

        var selected = new HashSet<int>();
        var groups = rows
            .Select((row, index) => (row, index))
            .GroupBy(x => string.IsNullOrWhiteSpace(x.row.RealizationNumber)
                ? $"#{x.index}"
                : $"{x.row.RealizationNumber}|{x.row.CheckType}",
                StringComparer.OrdinalIgnoreCase);

        foreach (var group in groups)
        {
            var items = group.ToList();
            var ready = items.Where(x => x.row.IsReady).ToList();
            if (ready.Count > 0)
            {
                var best = ready
                    .OrderByDescending(x => x.row.RegisteredAt ?? DateTime.MinValue)
                    .ThenByDescending(x => x.index)
                    .First();
                selected.Add(best.index);
            }
            else
            {
                selected.Add(items[^1].index);
            }
        }

        return rows.Where((_, index) => selected.Contains(index)).ToList();
    }

    private static OneCExportRow ErrorRow(XmlReportCheck xml, string status) => new()
    {
        RealizationNumber = xml.RealizationNumber,
        DocumentDate = ParseDocumentDate(xml.BaseDate),
        CheckType = xml.Operation,
        ExternalId = xml.ExternalId,
        Status = status,
        OfdStatus = string.Empty,
        IsReady = false,
        SourcePath = xml.SourcePath,
    };

    private static DateTime? ParseDocumentDate(string value)
    {
        if (string.IsNullOrWhiteSpace(value)) return null;
        var text = value.Trim();
        string[] formats = ["yyyy-MM-dd", "dd.MM.yyyy", "yyyy-MM-ddTHH:mm:ss"];
        if (DateTime.TryParseExact(text, formats, CultureInfo.InvariantCulture,
                DateTimeStyles.None, out var exact))
            return exact.Date;
        return DateTime.TryParse(text, CultureInfo.GetCultureInfo("ru-RU"),
            DateTimeStyles.None, out var parsed)
            ? parsed.Date
            : null;
    }

    private static string Clean(string value) =>
        value.Replace(';', ',').Replace('\r', ' ').Replace('\n', ' ').Trim();
}
