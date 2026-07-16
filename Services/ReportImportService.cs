using System.Globalization;
using System.IO;
using System.Text;
using System.Text.Json;
using System.Text.RegularExpressions;
using System.Xml.Linq;
using AtolGenerator.Models;
using ClosedXML.Excel;
using Microsoft.VisualBasic.FileIO;

namespace AtolGenerator.Services;

public static partial class ReportImportService
{
    private static readonly string[] DateFormats =
    {
        "yyyy-MM-ddTHH:mm:ss.fff",
        "yyyy-MM-ddTHH:mm:ss",
        "dd.MM.yyyy HH:mm:ss",
        "dd.MM.yyyy",
    };

    public static List<LocalPunchReportRow> ReadLocalPunchHistory(string path)
    {
        var rows = new List<LocalPunchReportRow>();
        if (!File.Exists(path)) return rows;

        foreach (var line in File.ReadLines(path))
        {
            if (string.IsNullOrWhiteSpace(line)) continue;
            try
            {
                using var document = JsonDocument.Parse(line);
                var root = document.RootElement;
                rows.Add(new LocalPunchReportRow
                {
                    RegisteredAt = ParseDate(GetString(root, "receipt_dt")) ?? ParseDate(GetString(root, "ts")),
                    Operation = GetString(root, "operation"),
                    OrderNumber = GetString(root, "order_num"),
                    RealizationNumber = GetString(root, "realization_num"),
                    Amount = GetDouble(root, "amount"),
                    FiscalDocument = GetInt64(root, "fiscal_doc"),
                    FiscalSign = GetInt64(root, "fiscal_sign"),
                    Uuid = GetString(root, "uuid"),
                    Cashier = GetString(root, "cashier"),
                    OfdUrl = GetString(root, "ofd_url"),
                });
            }
            catch (JsonException)
            {
                // Повреждённая строка журнала не должна мешать чтению остальных.
            }
        }

        return rows.OrderByDescending(r => r.RegisteredAt).ToList();
    }

    public static List<AtolJournalReportRow> ReadAtolJournal(string path)
    {
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);
        using var parser = new TextFieldParser(path, Encoding.GetEncoding(1251))
        {
            TextFieldType = FieldType.Delimited,
            HasFieldsEnclosedInQuotes = true,
            TrimWhiteSpace = false,
        };
        parser.SetDelimiters(";");

        var header = parser.ReadFields() ?? throw new InvalidDataException("CSV АТОЛ не содержит заголовок.");
        var columns = header
            .Select((name, index) => (Name: name.Trim(), Index: index))
            .Where(x => !string.IsNullOrWhiteSpace(x.Name))
            .ToDictionary(x => x.Name, x => x.Index, StringComparer.OrdinalIgnoreCase);

        RequireColumns(columns, "ФПД", "Номер ФД", "Тип чека", "Зарегистрирован на кассе", "External Id");

        var rows = new List<AtolJournalReportRow>();
        while (!parser.EndOfData)
        {
            string[]? fields;
            try
            {
                fields = parser.ReadFields();
            }
            catch (MalformedLineException ex)
            {
                throw new InvalidDataException($"Не удалось разобрать строку {ex.Message}", ex);
            }

            if (fields is null || fields.All(string.IsNullOrWhiteSpace)) continue;

            string Field(string name) =>
                columns.TryGetValue(name, out var index) && index < fields.Length ? fields[index].Trim() : string.Empty;

            var incomingJson = Field("JSON входящего чека");
            var resultJson = Field("JSON результата обработки");
            var jsonInfo = ReadAtolJson(incomingJson, resultJson);
            var russianType = Field("Тип чека");

            rows.Add(new AtolJournalReportRow
            {
                RegisteredAt = ParseDate(Field("Зарегистрирован на кассе")) ?? jsonInfo.RegisteredAt,
                ReceivedAt = ParseDate(Field("Поступил в систему")),
                CheckType = russianType,
                Operation = NormalizeOperation(jsonInfo.DocumentType, russianType),
                Status = Field("Статус чека"),
                Source = Field("Источник данных"),
                Amount = ParseDouble(Field("Сумма чека, руб.")),
                FiscalDocument = ParseInt64(Field("Номер ФД")) ?? jsonInfo.FiscalDocument,
                FiscalSign = ParseInt64(Field("ФПД")) ?? jsonInfo.FiscalSign,
                FiscalDriveNumber = Field("Номер ФН"),
                Uuid = Field("UUID чека"),
                ExternalId = Field("External Id"),
                BaseNumber = jsonInfo.BaseNumber,
                BaseDate = jsonInfo.BaseDate,
                OfdUrl = jsonInfo.OfdUrl,
                IncomingJson = incomingJson,
                ResultJson = resultJson,
            });
        }

        return rows.OrderByDescending(r => r.RegisteredAt).ToList();
    }

    public static List<OfdReportRow> ReadOfdReport(string path)
    {
        return ReadOfdReportDetails(path).Rows;
    }

    public static OfdArchiveResult ReadOfdArchive(IEnumerable<string> paths)
    {
        var importedFiles = new List<string>();
        var truncatedFiles = new List<string>();
        var failedFiles = new List<string>();
        var rowsByKey = new Dictionary<string, OfdReportRow>(StringComparer.OrdinalIgnoreCase);

        foreach (var path in paths
                     .Where(File.Exists)
                     .Distinct(StringComparer.OrdinalIgnoreCase))
        {
            try
            {
                var report = ReadOfdReportDetails(path);
                importedFiles.Add(path);
                if (report.IsTruncated) truncatedFiles.Add(path);

                foreach (var row in report.Rows)
                {
                    var key = GetOfdReceiptKey(row);
                    if (!rowsByKey.ContainsKey(key)) rowsByKey.Add(key, row);
                }
            }
            catch
            {
                failedFiles.Add(path);
            }
        }

        return new OfdArchiveResult
        {
            Rows = MergeOfdRows(rowsByKey.Values),
            ImportedFiles = importedFiles,
            TruncatedFiles = truncatedFiles,
            FailedFiles = failedFiles,
        };
    }

    public static List<OfdReportRow> MergeOfdRows(IEnumerable<OfdReportRow> rows)
    {
        var rowsByKey = new Dictionary<string, OfdReportRow>(StringComparer.OrdinalIgnoreCase);
        foreach (var row in rows)
        {
            var key = GetOfdReceiptKey(row);
            if (!rowsByKey.ContainsKey(key)) rowsByKey.Add(key, row);
        }

        return rowsByKey.Values.OrderByDescending(row => row.RegisteredAt).ToList();
    }

    public static OfdReportReadResult ReadOfdReportDetails(string path)
    {
        using var workbook = new XLWorkbook(path);
        var sheet = workbook.Worksheets.First();
        var headerRow = FindHeaderRow(sheet, "Дата и время");
        if (headerRow == 0)
            throw new InvalidDataException("В отчёте ОФД не найдена строка заголовков с полем «Дата и время».");

        var lastColumn = sheet.LastColumnUsed()?.ColumnNumber() ?? 0;
        var columns = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
        for (var column = 1; column <= lastColumn; column++)
        {
            var name = sheet.Cell(headerRow, column).GetString().Trim();
            if (!string.IsNullOrWhiteSpace(name)) columns[name] = column;
        }
        RequireColumns(columns, "Дата и время", "Документ", "Сумма", "№ ФД", "ФПД");

        var result = new List<OfdReportRow>();
        var lastRow = sheet.LastRowUsed()?.RowNumber() ?? headerRow;
        for (var row = headerRow + 1; row <= lastRow; row++)
        {
            var fiscalSign = ReadLong(sheet.Cell(row, columns["ФПД"]));
            var fiscalDocument = ReadLong(sheet.Cell(row, columns["№ ФД"]));
            if (fiscalSign is null && fiscalDocument is null) continue;

            result.Add(new OfdReportRow
            {
                RegisteredAt = ReadDate(sheet.Cell(row, columns["Дата и время"])),
                Document = ReadText(sheet.Cell(row, columns["Документ"])),
                Operation = columns.TryGetValue("Тип операции", out var operationColumn)
                    ? ReadText(sheet.Cell(row, operationColumn))
                    : string.Empty,
                CalculationMethod = ReadCalculationMethod(sheet, row, columns),
                Amount = ReadDouble(sheet.Cell(row, columns["Сумма"])),
                FiscalDocument = fiscalDocument,
                FiscalSign = fiscalSign,
                FiscalDriveNumber = ReadOptionalText(sheet, row, columns, "Зав. № ФН"),
                KktRegistrationNumber = ReadOptionalText(sheet, row, columns, "Рег. № ККТ"),
                TradingPoint = columns.TryGetValue("Торговая точка", out var tradingPointColumn)
                    ? ReadText(sheet.Cell(row, tradingPointColumn))
                    : string.Empty,
                KktName = columns.TryGetValue("Название ККТ", out var kktNameColumn)
                    ? ReadText(sheet.Cell(row, kktNameColumn))
                    : string.Empty,
                ReceiptUrl = columns.TryGetValue("Ссылка на просмотр чека", out var urlColumn)
                    ? ReadText(sheet.Cell(row, urlColumn))
                    : string.Empty,
                SourceFile = Path.GetFileName(path),
            });
        }

        return new OfdReportReadResult
        {
            Rows = result.OrderByDescending(r => r.RegisteredAt).ToList(),
            IsTruncated = IsTruncatedOfdReport(sheet, headerRow, result.Count),
        };
    }

    public static List<XmlReportCheck> ReadXmlChecks(string path)
    {
        var document = XDocument.Load(path);
        var checks = document.Root?.Name.LocalName == "check"
            ? new[] { document.Root }
            : document.Root?.Elements().Where(e => e.Name.LocalName == "check") ?? Enumerable.Empty<XElement>();

        var result = new List<XmlReportCheck>();
        var index = 0;
        foreach (var check in checks)
        {
            var receipt = Child(check, "receipt");
            var correction = Child(check, "correction");
            if (receipt is null && correction is null) continue;

            var operation = Value(Child(receipt ?? correction!, "operation"));
            var externalId = Value(Child(check, "external_id"));
            var generatedAt = ParseDate(Value(Child(check, "timestamp")));

            string realizationNumber;
            string baseDate;
            double amount;
            string originalFiscalSign;

            if (correction is not null)
            {
                var correctionInfo = Child(correction, "correction_info");
                realizationNumber = Value(Child(correctionInfo, "base_number"));
                baseDate = Value(Child(correctionInfo, "base_date"));
                amount = ParseDouble(Value(Child(Child(Child(correction, "payments"), "payment"), "sum")));
                originalFiscalSign = string.Empty;
            }
            else
            {
                var userProps = Child(receipt!, "additional_user_props") ?? Child(receipt!, "additional_user_attribute");
                realizationNumber = Value(Child(userProps, "value"));
                if (string.IsNullOrWhiteSpace(realizationNumber))
                {
                    var itemName = Value(Child(Child(Child(receipt!, "items"), "item"), "name"));
                    realizationNumber = RealizationNumberRegex().Match(itemName).Value;
                }
                baseDate = string.Empty;
                amount = ParseDouble(Value(Child(receipt!, "total")));
                originalFiscalSign = Value(Child(receipt!, "additional_check_props"));
            }

            result.Add(new XmlReportCheck
            {
                Index = index++,
                GeneratedAt = generatedAt,
                ExternalId = externalId,
                Operation = operation,
                RealizationNumber = NormalizeRealizationNumber(realizationNumber),
                BaseDate = baseDate,
                Amount = amount,
                OriginalFiscalSign = originalFiscalSign,
            });
        }

        return result;
    }

    public static string NormalizeRealizationNumber(string value)
    {
        var result = value.Trim();
        const string prefix = "Основание_";
        if (result.StartsWith(prefix, StringComparison.OrdinalIgnoreCase))
            result = result[prefix.Length..];
        return result;
    }

    private static AtolJsonInfo ReadAtolJson(string incomingJson, string resultJson)
    {
        var info = new AtolJsonInfo();
        if (!string.IsNullOrWhiteSpace(incomingJson))
        {
            try
            {
                using var incoming = JsonDocument.Parse(incomingJson);
                var root = incoming.RootElement;
                info.DocumentType = GetString(root, "document_type");
                if (root.TryGetProperty("correction", out var correction) &&
                    correction.TryGetProperty("correction_info", out var correctionInfo))
                {
                    info.BaseNumber = GetString(correctionInfo, "base_number");
                    info.BaseDate = GetString(correctionInfo, "base_date");
                }
            }
            catch (JsonException)
            {
                // Основные колонки отчёта остаются пригодны даже при повреждённом JSON.
            }
        }

        if (!string.IsNullOrWhiteSpace(resultJson))
        {
            try
            {
                using var response = JsonDocument.Parse(resultJson);
                var root = response.RootElement;
                if (root.TryGetProperty("payload", out var payload))
                {
                    info.FiscalDocument = GetInt64(payload, "fiscal_document_number");
                    info.FiscalSign = GetInt64(payload, "fiscal_document_attribute");
                    info.RegisteredAt = ParseDate(GetString(payload, "receipt_datetime"));
                    info.OfdUrl = GetString(payload, "ofd_receipt_url");
                }
            }
            catch (JsonException)
            {
                // Основные колонки отчёта остаются пригодны даже при повреждённом JSON.
            }
        }

        return info;
    }

    private static string NormalizeOperation(string documentType, string russianType)
    {
        var normalized = documentType.Replace("_", string.Empty).Replace("-", string.Empty).Trim().ToLowerInvariant();
        if (normalized.Length > 0)
        {
            return normalized switch
            {
                "sell" => "sell",
                "sellrefund" => "sell_refund",
                "sellcorrection" => "sell_correction",
                "buyrefund" => "buy_refund",
                "buycorrection" => "buy_correction",
                _ => documentType.Trim().ToLowerInvariant(),
            };
        }

        return russianType.Trim().ToLowerInvariant() switch
        {
            "приход" => "sell",
            "возврат прихода" => "sell_refund",
            "коррекция прихода" => "sell_correction",
            "корректировка прихода" => "sell_correction",
            "коррекция расхода" => "buy_correction",
            "корректировка расхода" => "buy_correction",
            _ => string.Empty,
        };
    }

    private static void RequireColumns<T>(Dictionary<string, T> columns, params string[] required)
    {
        var missing = required.Where(name => !columns.ContainsKey(name)).ToArray();
        if (missing.Length > 0)
            throw new InvalidDataException($"Отсутствуют обязательные колонки: {string.Join(", ", missing)}.");
    }

    private static int FindHeaderRow(IXLWorksheet sheet, string marker)
    {
        var maxRow = Math.Min(sheet.LastRowUsed()?.RowNumber() ?? 0, 40);
        var maxColumn = Math.Min(sheet.LastColumnUsed()?.ColumnNumber() ?? 0, 60);
        for (var row = 1; row <= maxRow; row++)
        for (var column = 1; column <= maxColumn; column++)
        {
            if (string.Equals(sheet.Cell(row, column).GetString().Trim(), marker, StringComparison.OrdinalIgnoreCase))
                return row;
        }
        return 0;
    }

    private static bool IsTruncatedOfdReport(IXLWorksheet sheet, int headerRow, int receiptCount)
    {
        if (receiptCount >= 30000) return true;

        for (var row = 1; row < headerRow; row++)
        {
            var lastColumn = sheet.Row(row).LastCellUsed()?.Address.ColumnNumber ?? 0;
            for (var column = 1; column <= lastColumn; column++)
            {
                var text = sheet.Cell(row, column).GetString();
                if (text.Contains("30000", StringComparison.OrdinalIgnoreCase) &&
                    (text.Contains("последн", StringComparison.OrdinalIgnoreCase) ||
                     text.Contains("слишком много", StringComparison.OrdinalIgnoreCase)))
                    return true;
            }
        }

        return false;
    }

    private static string ReadCalculationMethod(
        IXLWorksheet sheet,
        int row,
        IReadOnlyDictionary<string, int> columns)
    {
        string[] methods =
        {
            "Предоплата 100%",
            "Частичная предоплата",
            "Аванс расчет",
            "Полный расчет",
            "Частичный расчет и кредит",
            "Передача в кредит",
            "Оплата кредита",
        };

        foreach (var method in methods)
        {
            if (columns.TryGetValue(method, out var column) &&
                Math.Abs(ReadDouble(sheet.Cell(row, column))) > 0.005)
                return method;
        }

        return string.Empty;
    }

    private static string ReadOptionalText(
        IXLWorksheet sheet,
        int row,
        IReadOnlyDictionary<string, int> columns,
        string columnName)
    {
        return columns.TryGetValue(columnName, out var column)
            ? ReadText(sheet.Cell(row, column))
            : string.Empty;
    }

    private static string GetOfdReceiptKey(OfdReportRow row)
    {
        var fiscalDrive = NormalizeIdentifier(row.FiscalDriveNumber);
        if (fiscalDrive.Length > 0 && row.FiscalDocument.HasValue)
            return $"fn:{fiscalDrive}|fd:{row.FiscalDocument.Value}";

        var kkt = NormalizeIdentifier(row.KktRegistrationNumber);
        if (kkt.Length > 0 && row.FiscalDocument.HasValue && row.FiscalSign.HasValue)
            return $"kkt:{kkt}|fd:{row.FiscalDocument.Value}|fp:{row.FiscalSign.Value}";

        return $"fp:{row.FiscalSign?.ToString(CultureInfo.InvariantCulture) ?? "-"}" +
               $"|fd:{row.FiscalDocument?.ToString(CultureInfo.InvariantCulture) ?? "-"}" +
               $"|dt:{row.RegisteredAt:O}|sum:{row.Amount.ToString("0.00", CultureInfo.InvariantCulture)}";
    }

    private static string NormalizeIdentifier(string value) =>
        value.Replace(" ", string.Empty).Trim();

    private static DateTime? ReadDate(IXLCell cell)
    {
        try
        {
            if (cell.DataType == XLDataType.DateTime) return cell.GetDateTime();
        }
        catch
        {
            // Fallback ниже обрабатывает текстовое представление даты.
        }
        return ParseDate(cell.GetString());
    }

    private static string ReadText(IXLCell cell) => cell.GetString().Trim();

    private static long? ReadLong(IXLCell cell)
    {
        if (cell.DataType == XLDataType.Number)
            return Convert.ToInt64(Math.Round(cell.GetDouble()));
        return ParseInt64(cell.GetString());
    }

    private static double ReadDouble(IXLCell cell)
    {
        if (cell.DataType == XLDataType.Number) return cell.GetDouble();
        return ParseDouble(cell.GetString());
    }

    private static XElement? Child(XElement? parent, string localName) =>
        parent?.Elements().FirstOrDefault(e => e.Name.LocalName == localName);

    private static string Value(XElement? element) => element?.Value.Trim() ?? string.Empty;

    private static DateTime? ParseDate(string value)
    {
        if (string.IsNullOrWhiteSpace(value)) return null;
        if (DateTime.TryParseExact(value.Trim(), DateFormats, CultureInfo.InvariantCulture,
                DateTimeStyles.AssumeLocal, out var exact))
            return exact;
        return DateTime.TryParse(value, CultureInfo.GetCultureInfo("ru-RU"), DateTimeStyles.AssumeLocal, out var parsed)
            ? parsed
            : null;
    }

    private static double ParseDouble(string value)
    {
        var normalized = value.Trim().Replace(" ", string.Empty).Replace(',', '.');
        return double.TryParse(normalized, NumberStyles.Any, CultureInfo.InvariantCulture, out var result)
            ? result
            : 0;
    }

    private static long? ParseInt64(string value)
    {
        var normalized = value.Trim().Replace(" ", string.Empty);
        if (long.TryParse(normalized, NumberStyles.Integer, CultureInfo.InvariantCulture, out var integer))
            return integer;
        if (double.TryParse(normalized.Replace(',', '.'), NumberStyles.Any, CultureInfo.InvariantCulture, out var number))
            return Convert.ToInt64(Math.Round(number));
        return null;
    }

    private static string GetString(JsonElement root, string name)
    {
        if (!root.TryGetProperty(name, out var value)) return string.Empty;
        return value.ValueKind == JsonValueKind.String ? value.GetString() ?? string.Empty : value.ToString();
    }

    private static long? GetInt64(JsonElement root, string name)
    {
        if (!root.TryGetProperty(name, out var value)) return null;
        if (value.ValueKind == JsonValueKind.Number && value.TryGetInt64(out var integer)) return integer;
        return ParseInt64(value.ToString());
    }

    private static double GetDouble(JsonElement root, string name)
    {
        if (!root.TryGetProperty(name, out var value)) return 0;
        if (value.ValueKind == JsonValueKind.Number && value.TryGetDouble(out var number)) return number;
        return ParseDouble(value.ToString());
    }

    [GeneratedRegex(@"т\d{7}", RegexOptions.IgnoreCase | RegexOptions.CultureInvariant)]
    private static partial Regex RealizationNumberRegex();

    private sealed class AtolJsonInfo
    {
        public string DocumentType { get; set; } = string.Empty;
        public string BaseNumber { get; set; } = string.Empty;
        public string BaseDate { get; set; } = string.Empty;
        public long? FiscalDocument { get; set; }
        public long? FiscalSign { get; set; }
        public DateTime? RegisteredAt { get; set; }
        public string OfdUrl { get; set; } = string.Empty;
    }
}
