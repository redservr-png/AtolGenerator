using System.Runtime.InteropServices;

namespace AtolGenerator.Services;

public class OneCConnectionSettings
{
    public string Server   { get; set; } = string.Empty;  // Srvr=server1c
    public string Database { get; set; } = string.Empty;  // Ref=ut_new
    public string User     { get; set; } = string.Empty;
    public string Password { get; set; } = string.Empty;

    public string ConnectionString =>
        $"Srvr=\"{Server}\";Ref=\"{Database}\";Usr=\"{User}\";Pwd=\"{Password}\";";
}

public class OneCRealization
{
    public string DocNumber      { get; set; } = string.Empty;  // т0000025218
    public string DocDate        { get; set; } = string.Empty;  // дата реализации
    public string OrderNumber    { get; set; } = string.Empty;  // номер заказа покупателя
    public string OrderDate      { get; set; } = string.Empty;
    public string CustomerName   { get; set; } = string.Empty;
    public double Amount         { get; set; }
    public bool   IsService      { get; set; }  // агентский договор
    public string City           { get; set; } = string.Empty;
    public bool   HasCheck       { get; set; }  // чек уже пробит
    public string CheckNumber    { get; set; } = string.Empty;
    public string CheckDate      { get; set; } = string.Empty;
}

public static class OneCService
{
    public static bool IsAvailable()
    {
        try
        {
            var t = Type.GetTypeFromProgID("V83.COMConnector");
            return t is not null;
        }
        catch { return false; }
    }

    public static string TestConnection(OneCConnectionSettings s)
    {
        try
        {
            var connector = CreateConnector();
            dynamic conn = connector.Connect(s.ConnectionString);
            var version = (string)conn.Метаданные.Версия;
            Marshal.ReleaseComObject(conn);
            Marshal.ReleaseComObject(connector);
            return $"OK: подключено (конфигурация v{version})";
        }
        catch (Exception ex)
        {
            return $"Ошибка: {ex.Message}";
        }
    }

    public static List<OneCRealization> LoadRealizations(
        OneCConnectionSettings s, DateTime from, DateTime to)
    {
        dynamic? conn = null;
        dynamic? connector = null;
        var result = new List<OneCRealization>();

        try
        {
            connector = CreateConnector();
            conn = connector.Connect(s.ConnectionString);

            var query = conn.NewObject("Запрос");
            query.Текст = BuildQuery();
            query.УстановитьПараметр("НачалоПериода", from.Date);
            query.УстановитьПараметр("КонецПериода",  to.Date.AddDays(1).AddSeconds(-1));

            var queryResult = query.Выполнить();
            var selection   = queryResult.Выбрать();

            while ((bool)selection.Следующий())
            {
                // Ссылка: "Реализация товаров и услуг т0000025218 от 12.03.2026 19:38:36"
                var ssylka     = selection.Ссылка?.ToString() ?? string.Empty;
                var docNumber  = ExtractDocNumber(ssylka);
                var docDate    = ExtractDate(ssylka);

                // Сделка: "Заказ покупателя т0000018913 от 23.02.2026 16:50:38"
                var sdelkaStr  = string.Empty;
                try { sdelkaStr = selection.Сделка?.ToString() ?? string.Empty; }
                catch { /* Сделка может быть пустой */ }
                var orderNum   = ExtractDocNumber(sdelkaStr);
                var orderDate  = ExtractDate(sdelkaStr);

                // Дата чека (может быть пустой датой 0001-01-01 в 1С = DateTime.MinValue)
                var checkDt    = (DateTime)selection.ДатаПечатиЧека;
                var hasCheck   = !string.IsNullOrEmpty(selection.НомерЧекаККМ?.ToString())
                              && checkDt > new DateTime(2000, 1, 1);

                // Договор: если содержит "агент" → IsService
                var dogovor    = selection.Договор?.ToString() ?? string.Empty;
                var isService  = dogovor.IndexOf("агент", StringComparison.OrdinalIgnoreCase) >= 0;

                var r = new OneCRealization
                {
                    DocNumber    = docNumber,
                    DocDate      = docDate,
                    OrderNumber  = orderNum,
                    OrderDate    = orderDate,
                    CustomerName = selection.Покупатель?.ToString() ?? string.Empty,
                    Amount       = (double)selection.СуммаДокумента,
                    IsService    = isService,
                    City         = selection.Подразделение?.ToString() ?? string.Empty,
                    HasCheck     = hasCheck,
                    CheckNumber  = selection.НомерЧекаККМ?.ToString() ?? string.Empty,
                    CheckDate    = hasCheck
                                    ? checkDt.ToString("dd.MM.yyyy HH:mm:ss")
                                    : string.Empty,
                };
                result.Add(r);
            }
        }
        finally
        {
            if (conn      is not null) Marshal.ReleaseComObject(conn);
            if (connector is not null) Marshal.ReleaseComObject(connector);
        }

        return result;
    }

    private static dynamic CreateConnector()
    {
        var t = Type.GetTypeFromProgID("V83.COMConnector")
             ?? throw new InvalidOperationException("V83.COMConnector не найден. Установите клиент 1С.");
        return Activator.CreateInstance(t)
            ?? throw new InvalidOperationException("Не удалось создать экземпляр V83.COMConnector");
    }

    // ── String helpers (same logic as ExcelImportService) ────────────────────

    // "Реализация товаров и услуг т0000025218 от 12.03.2026 19:38:36" → "12.03.2026 19:38:36"
    private static string ExtractDate(string text)
    {
        var idx = text.IndexOf(" от ", StringComparison.Ordinal);
        if (idx < 0) return string.Empty;
        return text[(idx + 4)..].Trim();
    }

    // "Реализация товаров и услуг т0000025218 от ..." → "т0000025218"
    private static string ExtractDocNumber(string text)
    {
        if (string.IsNullOrWhiteSpace(text)) return string.Empty;
        var parts = text.Split(' ');
        foreach (var p in parts)
        {
            if (p.Length > 2 && (p[0] == 'т' || p[0] == 'Т') && char.IsDigit(p[1]))
                return p;
        }
        return text;
    }

    // ── Запрос к УТ 10.3 ─────────────────────────────────────────────────────
    // TODO: скорректировать поля под реальную структуру базы
    private static string BuildQuery() => """
        ВЫБРАТЬ
            РеализацияТоваровУслуг.Ссылка                                   КАК Ссылка,
            РеализацияТоваровУслуг.Сделка.КонтактноеЛицоКонтрагента        КАК Покупатель,
            РеализацияТоваровУслуг.Сделка                                   КАК Сделка,
            РеализацияТоваровУслуг.СуммаДокумента                          КАК СуммаДокумента,
            РеализацияТоваровУслуг.ДоговорКонтрагента.Наименование         КАК Договор,
            РеализацияТоваровУслуг.Подразделение                           КАК Подразделение,
            РеализацияТоваровУслуг.НомерЧекаККМ                            КАК НомерЧекаККМ,
            РеализацияТоваровУслуг.ККМ                                     КАК ККМ,
            РеализацияТоваровУслуг.ЧекНомерФП                              КАК ЧекНомерФП,
            РеализацияТоваровУслуг.ДатаПечатиЧека                         КАК ДатаПечатиЧека
        ИЗ
            Документ.РеализацияТоваровУслуг КАК РеализацияТоваровУслуг
        ГДЕ
            РеализацияТоваровУслуг.ПометкаУдаления = ЛОЖЬ
            И РеализацияТоваровУслуг.Проведен = ИСТИНА
            И РеализацияТоваровУслуг.ЭтоРекламация = ЛОЖЬ
            И РеализацияТоваровУслуг.Дата МЕЖДУ &НачалоПериода И &КонецПериода
            И РеализацияТоваровУслуг.Подразделение.Наименование <> "OZON"
            И РеализацияТоваровУслуг.Подразделение.Наименование <> "Вологда ОПТ"
            И РеализацияТоваровУслуг.Подразделение.Наименование <> "Новодвинск"
            И РеализацияТоваровУслуг.Подразделение.Наименование <> "Интернет-магазин (продажи)"
            И РеализацияТоваровУслуг.Сделка.Контрагент.Наименование = "Розничный покупатель"
            И РеализацияТоваровУслуг.СуммаДокумента > 0
            И НАЧАЛОПЕРИОДА(РеализацияТоваровУслуг.Дата, ДЕНЬ) <> НАЧАЛОПЕРИОДА(РеализацияТоваровУслуг.ДатаПечатиЧека, ДЕНЬ)
            И (НЕ РеализацияТоваровУслуг.Комментарий ПОДОБНО "%Пробит%"
                    ИЛИ РеализацияТоваровУслуг.Комментарий ЕСТЬ NULL)
        УПОРЯДОЧИТЬ ПО
            РеализацияТоваровУслуг.Подразделение.Наименование,
            РеализацияТоваровУслуг.Дата
        """;
    // Параметры даты устанавливаются отдельно через query.УстановитьПараметр
}
