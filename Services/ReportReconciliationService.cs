using System.Globalization;
using System.IO;
using System.Text;
using AtolGenerator.Models;

namespace AtolGenerator.Services;

public static class ReportReconciliationService
{
    public static List<OneCExportRow> Build(
        IReadOnlyCollection<XmlReportCheck> xmlChecks,
        IReadOnlyCollection<AtolJournalReportRow> atolChecks,
        IReadOnlyCollection<OfdReportRow> ofdRows)
    {
        var atolByExternalId = atolChecks
            .Where(x => !string.IsNullOrWhiteSpace(x.ExternalId))
            .GroupBy(x => x.ExternalId, StringComparer.OrdinalIgnoreCase)
            .ToDictionary(g => g.Key, g => g.OrderByDescending(x => x.RegisteredAt).First(), StringComparer.OrdinalIgnoreCase);

        var ofdByFiscalSign = ofdRows
            .Where(x => x.FiscalSign.HasValue)
            .GroupBy(x => x.FiscalSign!.Value)
            .ToDictionary(g => g.Key, g => g.ToList());

        var pairRealizations = xmlChecks
            .Where(IsSupportedOperation)
            .Where(x => !string.IsNullOrWhiteSpace(x.RealizationNumber))
            .GroupBy(x => x.RealizationNumber, StringComparer.OrdinalIgnoreCase)
            .Where(g => g.Any(x => x.Operation == "sell_refund") && g.Any(x => x.Operation == "sell_correction"))
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

            if (!atolByExternalId.TryGetValue(xml.ExternalId, out var atol))
            {
                result.Add(ErrorRow(xml, "Чек не найден в отчёте АТОЛ"));
                continue;
            }

            if (!atol.FiscalSign.HasValue || !atol.FiscalDocument.HasValue || !atol.RegisteredAt.HasValue)
            {
                result.Add(ErrorRow(xml, "В отчёте АТОЛ не заполнены ФПД, ФД или дата чека"));
                continue;
            }

            if (!string.IsNullOrWhiteSpace(atol.Operation) &&
                !string.Equals(atol.Operation, xml.Operation, StringComparison.OrdinalIgnoreCase))
            {
                result.Add(ErrorRow(xml, $"Тип операции не совпадает: XML {xml.Operation}, АТОЛ {atol.Operation}"));
                continue;
            }

            var isPair = pairRealizations.Contains(xml.RealizationNumber);
            if (!isPair && xml.Operation == "sell_refund")
            {
                result.Add(ErrorRow(xml, "Одиночный возврат не относится к загрузке коррекции реализации"));
                continue;
            }

            var writeMode = isPair ? "comment_only" : "update_fields";
            var comment = BuildComment(xml.Operation, atol.RegisteredAt.Value, atol.FiscalSign.Value, isPair);
            var ofdStatus = BuildOfdStatus(atol, ofdRows.Count, ofdByFiscalSign);

            result.Add(new OneCExportRow
            {
                RealizationNumber = xml.RealizationNumber,
                CheckType = xml.Operation,
                WriteMode = writeMode,
                ExternalId = xml.ExternalId,
                FiscalSign = atol.FiscalSign,
                FiscalDocument = atol.FiscalDocument,
                RegisteredAt = atol.RegisteredAt,
                Comment = comment,
                OfdStatus = ofdStatus,
                Status = "Готово",
                IsReady = true,
            });
        }

        return result;
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

    private static bool IsSupportedOperation(XmlReportCheck check) =>
        check.Operation is "sell_correction" or "sell_refund";

    private static string BuildComment(string operation, DateTime registeredAt, long fiscalSign, bool isPair)
    {
        var date = registeredAt.ToString("dd.MM.yyyy", CultureInfo.InvariantCulture);
        return operation switch
        {
            "sell_refund" when isPair =>
                $"{date} Пробит исправительный чек \"Возврат прихода\" ФП: {fiscalSign}",
            "sell_correction" =>
                $"{date} Пробит чек коррекции \"Приход\" ФП: {fiscalSign}",
            _ => string.Empty,
        };
    }

    private static string BuildOfdStatus(
        AtolJournalReportRow atol,
        int ofdCount,
        IReadOnlyDictionary<long, List<OfdReportRow>> ofdByFiscalSign)
    {
        if (ofdCount == 0) return "ОФД не загружен";
        if (!atol.FiscalSign.HasValue || !ofdByFiscalSign.TryGetValue(atol.FiscalSign.Value, out var candidates))
            return "Не найден в ОФД";

        var exact = candidates.Any(x =>
            (!atol.FiscalDocument.HasValue || x.FiscalDocument == atol.FiscalDocument) &&
            Math.Abs(Math.Abs(x.Amount) - Math.Abs(atol.Amount)) < 0.01);
        return exact ? "Проверено ОФД" : "Расхождение с ОФД";
    }

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
