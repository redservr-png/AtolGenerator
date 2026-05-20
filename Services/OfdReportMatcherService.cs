using System.IO;
using System.Text.Json;
using ClosedXML.Excel;

namespace AtolGenerator.Services;

/// <summary>
/// Сопоставляет Excel-отчёт ОФД (Сводный отчёт по фискальным документам) с локальным
/// логом пробитий (punched_checks.jsonl) по полю ФПД и формирует Excel-таблицу:
///   ФПД | № ФД | дата | сумма | № реализации 1С | ссылка на чек.
///
/// Используется для пакетного обновления документов реализации в 1С (после
/// автоматической или ручной коррекции).
/// </summary>
public static class OfdReportMatcherService
{
    public class PunchedEntry
    {
        public string ts { get; set; } = string.Empty;
        public string operation { get; set; } = string.Empty;
        public string order_num { get; set; } = string.Empty;
        public string realization_num { get; set; } = string.Empty;
        public double amount { get; set; }
        public string uuid { get; set; } = string.Empty;
        public long? fiscal_doc { get; set; }
        public long? fiscal_sign { get; set; }
        public string receipt_dt { get; set; } = string.Empty;
        public string ofd_url { get; set; } = string.Empty;
        public string cashier { get; set; } = string.Empty;
    }

    public class MatchResult
    {
        public int TotalRows      { get; set; }
        public int MatchedRows    { get; set; }
        public int UnmatchedRows  { get; set; }
        public string OutputPath  { get; set; } = string.Empty;
    }

    /// <summary>
    /// Читает punched_checks.jsonl и возвращает map ФПД → запись.
    /// </summary>
    private static Dictionary<long, PunchedEntry> LoadPunchedMap()
    {
        var map = new Dictionary<long, PunchedEntry>();
        var path = AtolApiService.PunchedJsonPath;
        if (!File.Exists(path)) return map;

        foreach (var line in File.ReadLines(path))
        {
            if (string.IsNullOrWhiteSpace(line)) continue;
            try
            {
                var e = JsonSerializer.Deserialize<PunchedEntry>(line);
                if (e?.fiscal_sign is long fs) map[fs] = e;
            }
            catch { /* пропускаем битые строки */ }
        }
        return map;
    }

    /// <summary>
    /// Парсит отчёт ОФД и формирует выходной Excel с привязкой к реализациям.
    /// </summary>
    public static MatchResult MatchAndExport(string ofdReportPath, string outputPath)
    {
        var map = LoadPunchedMap();
        var result = new MatchResult { OutputPath = outputPath };

        using var src = new XLWorkbook(ofdReportPath);
        var srcSheet = src.Worksheets.First();
        // Шапка в отчёте Такском — на 11-й строке (индекс 11 в 1-based)
        var headerRow = 11;
        var firstDataRow = 12;
        var lastRow = srcSheet.LastRowUsed()?.RowNumber() ?? headerRow;

        // Находим индексы нужных колонок
        var colMap = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
        for (int c = 1; c <= srcSheet.LastColumnUsed()!.ColumnNumber(); c++)
        {
            var name = srcSheet.Cell(headerRow, c).GetString().Trim();
            if (!string.IsNullOrEmpty(name)) colMap[name] = c;
        }

        int colDate    = colMap.GetValueOrDefault("Дата и время",          1);
        int colDoc     = colMap.GetValueOrDefault("Документ",              2);
        int colShift   = colMap.GetValueOrDefault("№ смены",               3);
        int colShNum   = colMap.GetValueOrDefault("№ за смену",            4);
        int colOper    = colMap.GetValueOrDefault("Тип операции",          6);
        int colSum     = colMap.GetValueOrDefault("Сумма",                12);
        int colFd      = colMap.GetValueOrDefault("№ ФД",                 28);
        int colFp      = colMap.GetValueOrDefault("ФПД",                  29);
        int colCashier = colMap.GetValueOrDefault("Кассир",               27);
        int colKkt     = colMap.GetValueOrDefault("Название ККТ",         32);
        int colLink    = colMap.GetValueOrDefault("Ссылка на просмотр чека", 46);

        // Готовим выходной файл
        using var dst = new XLWorkbook();
        var ws = dst.Worksheets.Add("Сопоставление");
        // Шапка
        var hdr = new[]
        {
            "Дата чека", "Документ", "Тип операции",
            "№ ФД", "ФПД", "Сумма",
            "№ реализации 1С", "№ заказа", "Дата пробития",
            "Кассир", "ККТ", "Ссылка на чек", "Статус",
        };
        for (int i = 0; i < hdr.Length; i++)
        {
            ws.Cell(1, i + 1).Value = hdr[i];
            ws.Cell(1, i + 1).Style.Font.Bold = true;
            ws.Cell(1, i + 1).Style.Fill.BackgroundColor = XLColor.LightGray;
        }

        int outRow = 2;
        int matched = 0, unmatched = 0;
        for (int r = firstDataRow; r <= lastRow; r++)
        {
            var fpStr = srcSheet.Cell(r, colFp).GetString().Trim();
            if (string.IsNullOrEmpty(fpStr)) continue;
            if (!long.TryParse(fpStr, out var fp)) continue;

            var fdStr = srcSheet.Cell(r, colFd).GetString().Trim();
            long.TryParse(fdStr, out var fd);

            map.TryGetValue(fp, out var entry);

            ws.Cell(outRow, 1).Value  = srcSheet.Cell(r, colDate).GetString();
            ws.Cell(outRow, 2).Value  = srcSheet.Cell(r, colDoc).GetString();
            ws.Cell(outRow, 3).Value  = srcSheet.Cell(r, colOper).GetString();
            ws.Cell(outRow, 4).Value  = fd;
            ws.Cell(outRow, 5).Value  = fp;
            ws.Cell(outRow, 6).Value  = srcSheet.Cell(r, colSum).GetString();
            ws.Cell(outRow, 7).Value  = entry?.realization_num ?? "";
            ws.Cell(outRow, 8).Value  = entry?.order_num       ?? "";
            ws.Cell(outRow, 9).Value  = entry?.ts              ?? "";
            ws.Cell(outRow, 10).Value = srcSheet.Cell(r, colCashier).GetString();
            ws.Cell(outRow, 11).Value = srcSheet.Cell(r, colKkt).GetString();
            ws.Cell(outRow, 12).Value = srcSheet.Cell(r, colLink).GetString();
            ws.Cell(outRow, 13).Value = entry is null ? "Не найден" : "Сопоставлен";

            if (entry is null)
            {
                ws.Row(outRow).Style.Fill.BackgroundColor = XLColor.LightYellow;
                unmatched++;
            }
            else
            {
                matched++;
            }
            outRow++;
        }

        ws.Columns().AdjustToContents();
        dst.SaveAs(outputPath);

        result.TotalRows     = outRow - 2;
        result.MatchedRows   = matched;
        result.UnmatchedRows = unmatched;
        return result;
    }
}
