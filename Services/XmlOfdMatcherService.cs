using System.IO;
using System.Xml.Linq;
using ClosedXML.Excel;

namespace AtolGenerator.Services;

/// <summary>
/// Связывает старый XML-файл коррекций (без тега 1086) с отчётом ОФД и
/// формирует список записей PunchedRecord для записи в 1С.
///
/// Используется когда чеки уже пробиты, но в XML не было номера реализации.
/// Алгоритм: парсим XML в порядке следования, парсим OFD-чеки коррекции
/// в хронологическом порядке, сопоставляем по сумме (sequence + sum match).
/// </summary>
public static class XmlOfdMatcherService
{
    public class XmlCheck
    {
        public int    Index         { get; set; }
        public string BaseNumber    { get; set; } = string.Empty;
        public double Sum           { get; set; }
        public string ExternalId    { get; set; } = string.Empty;
    }

    public class OfdRow
    {
        public DateTime Date          { get; set; }
        public int      ShiftNum      { get; set; }
        public long     FiscalDoc     { get; set; }
        public long     FiscalSign    { get; set; }
        public double   Sum           { get; set; }
        public string   DocType       { get; set; } = string.Empty;
    }

    public class MatchReport
    {
        public int Total             { get; set; }
        public int Matched           { get; set; }
        public int Unmatched         { get; set; }
        public List<string> Warnings { get; set; } = new();
        public List<OneCService.PunchedRecord> Records { get; set; } = new();
    }

    /// <summary>
    /// Парсит XML-файл коррекций и возвращает список (порядок, № реализации, сумма).
    /// </summary>
    public static List<XmlCheck> ReadXmlChecks(string xmlPath)
    {
        var result = new List<XmlCheck>();
        var doc = XDocument.Load(xmlPath);
        int idx = 0;
        foreach (var check in doc.Root!.Elements("check"))
        {
            var ext       = check.Element("external_id")?.Value ?? string.Empty;
            var correction = check.Element("correction");
            var receipt   = check.Element("receipt");

            // sell_correction
            string baseNumber = string.Empty;
            double sum        = 0;
            if (correction is not null)
            {
                baseNumber = correction.Element("correction_info")?
                                       .Element("base_number")?.Value ?? string.Empty;
                var firstPay = correction.Element("payments")?
                                         .Element("payment")?.Element("sum")?.Value;
                double.TryParse(firstPay,
                    System.Globalization.NumberStyles.Any,
                    System.Globalization.CultureInfo.InvariantCulture, out sum);
            }
            else if (receipt is not null)
            {
                // sell_refund — номер реализации в additional_user_props.
                // Старое имя оставлено как fallback для ранее созданных XML.
                var userProps = receipt.Element("additional_user_props")
                                ?? receipt.Element("additional_user_attribute");
                baseNumber = userProps?.Element("value")?.Value ?? string.Empty;
                if (string.IsNullOrEmpty(baseNumber))
                {
                    // fallback: парсим из имени позиции (если совсем старый XML)
                    var name = receipt.Element("items")?.Element("item")?.Element("name")?.Value ?? "";
                    var m = System.Text.RegularExpressions.Regex.Match(name, @"т\d{7}");
                    if (m.Success) baseNumber = m.Value;
                }
                var totalStr = receipt.Element("total")?.Value;
                double.TryParse(totalStr,
                    System.Globalization.NumberStyles.Any,
                    System.Globalization.CultureInfo.InvariantCulture, out sum);
            }

            if (!string.IsNullOrEmpty(baseNumber))
            {
                result.Add(new XmlCheck
                {
                    Index      = idx++,
                    BaseNumber = baseNumber,
                    Sum        = sum,
                    ExternalId = ext,
                });
            }
        }
        return result;
    }

    /// <summary>
    /// Читает из отчёта ОФД строки коррекций (Документ = «Кассовый чек коррекции»),
    /// упорядоченные по (Дата, № за смену).
    /// </summary>
    public static List<OfdRow> ReadOfdCorrections(string ofdReportPath)
    {
        var result = new List<OfdRow>();
        using var wb = new XLWorkbook(ofdReportPath);
        var ws = wb.Worksheets.First();
        const int headerRow = 11;
        var lastRow = ws.LastRowUsed()?.RowNumber() ?? headerRow;

        var colMap = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
        for (int c = 1; c <= ws.LastColumnUsed()!.ColumnNumber(); c++)
        {
            var name = ws.Cell(headerRow, c).GetString().Trim();
            if (!string.IsNullOrEmpty(name)) colMap[name] = c;
        }
        int colDate  = colMap.GetValueOrDefault("Дата и время",  1);
        int colDoc   = colMap.GetValueOrDefault("Документ",      2);
        int colShift = colMap.GetValueOrDefault("№ за смену",    4);
        int colSum   = colMap.GetValueOrDefault("Сумма",        12);
        int colFd    = colMap.GetValueOrDefault("№ ФД",         28);
        int colFp    = colMap.GetValueOrDefault("ФПД",          29);

        for (int r = headerRow + 1; r <= lastRow; r++)
        {
            var docType = ws.Cell(r, colDoc).GetString().Trim();
            if (!docType.Contains("коррекции", StringComparison.OrdinalIgnoreCase))
                continue;

            DateTime dt = DateTime.MinValue;
            var dc = ws.Cell(r, colDate);
            try
            {
                if (dc.DataType == XLDataType.DateTime) dt = dc.GetDateTime();
                else DateTime.TryParse(dc.GetString(), out dt);
            }
            catch { }

            double sum = 0;
            try
            {
                if (ws.Cell(r, colSum).DataType == XLDataType.Number)
                    sum = ws.Cell(r, colSum).GetDouble();
                else
                    double.TryParse(ws.Cell(r, colSum).GetString().Replace(',', '.'),
                        System.Globalization.NumberStyles.Any,
                        System.Globalization.CultureInfo.InvariantCulture, out sum);
            }
            catch { }

            if (!long.TryParse(ws.Cell(r, colFp).GetString().Trim(), out var fp)) continue;
            if (!long.TryParse(ws.Cell(r, colFd).GetString().Trim(), out var fd)) continue;
            int.TryParse(ws.Cell(r, colShift).GetString().Trim(), out var shiftNum);

            result.Add(new OfdRow
            {
                Date       = dt,
                ShiftNum   = shiftNum,
                FiscalDoc  = fd,
                FiscalSign = fp,
                Sum        = sum,
                DocType    = docType,
            });
        }

        return result.OrderBy(o => o.Date).ThenBy(o => o.ShiftNum).ToList();
    }

    /// <summary>
    /// Сопоставляет XML-чеки с OFD-строками. Каждый XML-чек берём по очереди и
    /// ищем ПЕРВУЮ неиспользованную OFD-строку с такой же суммой (±0.01).
    /// </summary>
    public static MatchReport Match(List<XmlCheck> xmlChecks, List<OfdRow> ofdRows)
    {
        var report = new MatchReport { Total = xmlChecks.Count };
        var used = new bool[ofdRows.Count];

        foreach (var x in xmlChecks)
        {
            int foundAt = -1;
            for (int i = 0; i < ofdRows.Count; i++)
            {
                if (used[i]) continue;
                if (Math.Abs(ofdRows[i].Sum - x.Sum) < 0.01)
                {
                    foundAt = i;
                    break;
                }
            }

            if (foundAt < 0)
            {
                report.Unmatched++;
                report.Warnings.Add($"{x.BaseNumber} (сумма {x.Sum:F2}) — не найдено в отчёте ОФД");
                continue;
            }

            used[foundAt] = true;
            var ofd = ofdRows[foundAt];
            report.Matched++;
            report.Records.Add(new OneCService.PunchedRecord
            {
                RealizationNum = x.BaseNumber,
                FiscalDoc      = ofd.FiscalDoc,
                FiscalSign     = ofd.FiscalSign,
                ReceiptDt      = ofd.Date.ToString("dd.MM.yyyy HH:mm:ss"),
            });
        }

        return report;
    }
}
