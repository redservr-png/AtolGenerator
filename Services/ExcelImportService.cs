using ClosedXML.Excel;
using AtolGenerator.Models;

namespace AtolGenerator.Services;

public class ExcelImportResult
{
    public List<OrderEntry> Orders        { get; set; } = new();
    public List<SkippedRow> SkippedRows   { get; set; } = new();
    public int TotalRows                  { get; set; }
}

public class SkippedRow
{
    public int    RowNum      { get; set; }
    public string OrderNum    { get; set; } = string.Empty;
    public string OrderDate   { get; set; } = string.Empty;
    public double Amount      { get; set; }
    public string CheckNum    { get; set; } = string.Empty;
    public string CheckDate   { get; set; } = string.Empty;
    public string Reason      { get; set; } = string.Empty;
}

public static class ExcelImportService
{
    // Столбцы Excel (1-based):
    // A=1  Реализация товаров и услуг <номер> от <дата>
    // B=2  ФИО покупателя
    // C=3  Заказ покупателя <номер> от <дата>
    // D=4  Сумма
    // E=5  Тип договора: "Основной договор" / "Агентский договор"
    // F=6  Город/подразделение
    // G=7  Номер кассы (пусто если чека не было)
    // H=8  Касса или "Пустая ссылка: Кассы ККМ"
    // I=9  Фискальный номер чека
    // J=10 Дата/время чека

    public static ExcelImportResult Import(string filePath)
    {
        var result = new ExcelImportResult();

        using var wb = new XLWorkbook(filePath);
        var ws = wb.Worksheets.First();

        foreach (var row in ws.RowsUsed())
        {
            result.TotalRows++;
            int r = row.RowNumber();

            var colA = row.Cell(1).GetString().Trim();
            var colB = row.Cell(2).GetString().Trim();
            var colC = row.Cell(3).GetString().Trim();
            var colD = row.Cell(4).GetString().Trim();
            var colE = row.Cell(5).GetString().Trim();
            var colF = row.Cell(6).GetString().Trim();
            var colG = row.Cell(7).GetString().Trim();
            var colI = row.Cell(9).GetString().Trim();
            var colJ = row.Cell(10).GetString().Trim();

            if (string.IsNullOrWhiteSpace(colA)) continue;

            // Парсим сумму
            if (!TryParseAmount(colD, out double amount)) continue;

            // Парсим дату реализации из колонки A
            var realizationDate = ExtractDate(colA);

            // Парсим заказ покупателя из колонки C
            var (orderNum, orderDate) = ParseOrderRef(colC);

            // Если данные чека заполнены — пропускаем с предупреждением
            bool hasExistingCheck = !string.IsNullOrEmpty(colG)
                                 || !string.IsNullOrEmpty(colI)
                                 || !string.IsNullOrEmpty(colJ);

            if (hasExistingCheck)
            {
                result.SkippedRows.Add(new SkippedRow
                {
                    RowNum    = r,
                    OrderNum  = orderNum,
                    OrderDate = realizationDate,
                    Amount    = amount,
                    CheckNum  = colI,
                    CheckDate = colJ,
                    Reason    = "Чек уже пробит — требует ручного исправления",
                });
                continue;
            }

            // Определяем тип услуги по договору
            bool isService = colE.Contains("Агентский", StringComparison.OrdinalIgnoreCase);

            var entry = new OrderEntry
            {
                OrderNum         = orderNum,
                OrderDate        = orderDate,
                Amount           = amount,
                CustomerName     = colB,
                CorrectionDate   = realizationDate,   // дата реализации = основание коррекции
                CorrectionNumber = ExtractDocNumber(colA),
                AgentInfo        = null,
                IsService        = isService,
            };

            result.Orders.Add(entry);
        }

        return result;
    }

    // "Реализация товаров и услуг т0000138544 от 12.03.2026 19:38:36" → "12.03.2026"
    private static string ExtractDate(string text)
    {
        var idx = text.IndexOf(" от ", StringComparison.Ordinal);
        if (idx < 0) return string.Empty;
        var rest = text[(idx + 4)..].Trim();
        // берём только дату без времени
        return rest.Length >= 10 ? rest[..10] : rest;
    }

    // "Реализация товаров и услуг т0000138544 от ..." → "т0000138544"
    private static string ExtractDocNumber(string text)
    {
        // ищем токен вида тXXXXXXX или ТXXXXXX
        var parts = text.Split(' ');
        foreach (var p in parts)
        {
            if (p.Length > 2 && (p[0] == 'т' || p[0] == 'Т') && char.IsDigit(p[1]))
                return p;
        }
        return text;
    }

    // "Заказ покупателя т0000018913 от 23.02.2026 16:50:38" → ("т0000018913", "23.02.2026 16:50:38")
    private static (string num, string date) ParseOrderRef(string text)
    {
        if (string.IsNullOrWhiteSpace(text)) return (string.Empty, string.Empty);

        var num  = ExtractDocNumber(text);
        var date = ExtractDate(text);
        return (num, date);
    }

    private static bool TryParseAmount(string s, out double amount)
    {
        s = s.Replace(',', '.').Trim();
        return double.TryParse(s,
            System.Globalization.NumberStyles.Any,
            System.Globalization.CultureInfo.InvariantCulture,
            out amount) && amount > 0;
    }
}
