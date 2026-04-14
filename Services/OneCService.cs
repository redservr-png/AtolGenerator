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
    public string FiscalNumber   { get; set; } = string.Empty;  // ЧекНомерФП
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
        catch (COMException ex)
        {
            var hint = ex.HResult switch
            {
                unchecked((int)0x8002801D) =>
                    " → Запустите от Администратора: regsvr32 \"C:\\Program Files\\1cv8\\[версия]\\bin\\comcntr.dll\"",
                unchecked((int)0x80040154) =>
                    " → V83.COMConnector не зарегистрирован. Установите клиент 1С.",
                _ => string.Empty
            };
            return $"Ошибка COM (0x{ex.HResult:X8}): {ex.Message}{hint}";
        }
        catch (Exception ex)
        {
            return $"Ошибка ({ex.GetType().Name}): {ex.Message}";
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
                try
                {
                    // Ссылка и Сделка теперь строки (ПРЕДСТАВЛЕНИЕ в запросе)
                    var ssylka    = Str(selection.Ссылка);
                    var sdelka    = Str(selection.Сделка);
                    var docNumber = ExtractDocNumber(ssylka);
                    var docDate   = ExtractDate(ssylka);
                    var orderNum  = ExtractDocNumber(sdelka);
                    var orderDate = ExtractDate(sdelka);

                    // Дата чека — может быть null если чек не пробит
                    var checkDt  = ToDateTime(selection.ДатаПечатиЧека);
                    var checkNum = Str(selection.НомерЧекаККМ);
                    var hasCheck = !string.IsNullOrEmpty(checkNum)
                                && checkDt > new DateTime(2000, 1, 1);

                    // Договор: если содержит "агент" → IsService
                    var dogovor   = Str(selection.Договор);
                    var isService = dogovor.IndexOf("агент", StringComparison.OrdinalIgnoreCase) >= 0;

                    var r = new OneCRealization
                    {
                        DocNumber    = docNumber,
                        DocDate      = docDate,
                        OrderNumber  = orderNum,
                        OrderDate    = orderDate,
                        CustomerName = Str(selection.Покупатель),
                        Amount       = ToDouble(selection.СуммаДокумента),
                        IsService    = isService,
                        City         = Str(selection.Подразделение),
                        HasCheck     = hasCheck,
                        CheckNumber  = checkNum,
                        FiscalNumber = Str(selection.ЧекНомерФП),
                        CheckDate    = hasCheck
                                        ? checkDt.ToString("dd.MM.yyyy HH:mm:ss")
                                        : string.Empty,
                    };
                    result.Add(r);
                }
                catch (Exception rowEx)
                {
                    // Пропускаем строку с ошибкой, не прерывая всю выгрузку
                    System.Diagnostics.Debug.WriteLine($"[1С] пропущена строка: {rowEx.Message}");
                }
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

    // Безопасные приведения COM-значений
    private static string Str(dynamic? v)
    {
        try { return v?.ToString() ?? string.Empty; }
        catch { return string.Empty; }
    }

    private static DateTime ToDateTime(dynamic? v)
    {
        try { return v is null ? DateTime.MinValue : (DateTime)v; }
        catch { return DateTime.MinValue; }
    }

    private static double ToDouble(dynamic? v)
    {
        try { return v is null ? 0.0 : (double)v; }
        catch { return 0.0; }
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
    // ПРЕДСТАВЛЕНИЕ() преобразует ссылочные объекты в строку прямо в запросе
    private static string BuildQuery() => """
        ВЫБРАТЬ
            ПРЕДСТАВЛЕНИЕ(РеализацияТоваровУслуг.Ссылка)                           КАК Ссылка,
            ПРЕДСТАВЛЕНИЕ(РеализацияТоваровУслуг.Сделка.КонтактноеЛицоКонтрагента) КАК Покупатель,
            ПРЕДСТАВЛЕНИЕ(РеализацияТоваровУслуг.Сделка)                           КАК Сделка,
            РеализацияТоваровУслуг.СуммаДокумента                                  КАК СуммаДокумента,
            РеализацияТоваровУслуг.ДоговорКонтрагента.Наименование                 КАК Договор,
            РеализацияТоваровУслуг.Подразделение.Наименование                      КАК Подразделение,
            РеализацияТоваровУслуг.НомерЧекаККМ                                    КАК НомерЧекаККМ,
            РеализацияТоваровУслуг.ЧекНомерФП                                      КАК ЧекНомерФП,
            РеализацияТоваровУслуг.ДатаПечатиЧека                                  КАК ДатаПечатиЧека
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
