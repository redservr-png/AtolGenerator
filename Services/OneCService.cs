using System.IO;
using System.Runtime.InteropServices;
using System.Text.Json;

namespace AtolGenerator.Services;

public class OneCConnectionSettings
{
    public string Server   { get; set; } = string.Empty;  // Srvr=server1c
    public string Database { get; set; } = string.Empty;  // Ref=ut_new
    public string User     { get; set; } = string.Empty;
    public string Password { get; set; } = string.Empty;

    public string ConnectionString =>
        $"Srvr=\"{Server}\";Ref=\"{Database}\";Usr=\"{User}\";Pwd=\"{Password}\";";

    private static string SettingsPath => Path.Combine(
        AppDomain.CurrentDomain.BaseDirectory, "onec_settings.json");

    public void Save()
    {
        var json = JsonSerializer.Serialize(this, new JsonSerializerOptions { WriteIndented = true });
        File.WriteAllText(SettingsPath, json);
    }

    public static OneCConnectionSettings Load()
    {
        try
        {
            if (!File.Exists(SettingsPath)) return new();
            var json = File.ReadAllText(SettingsPath);
            return JsonSerializer.Deserialize<OneCConnectionSettings>(json) ?? new();
        }
        catch { return new(); }
    }
}

public class OneCRealizationItem
{
    public string Name     { get; set; } = string.Empty;
    public double Quantity { get; set; } = 1;
    public double Sum      { get; set; }
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

    // Путь к лог-файлу (рядом с exe)
    public static string LogPath { get; } = Path.Combine(
        AppDomain.CurrentDomain.BaseDirectory, "1c_log.txt");

    private static void Log(string message)
    {
        try
        {
            File.AppendAllText(LogPath,
                $"[{DateTime.Now:dd.MM.yyyy HH:mm:ss}] {message}{Environment.NewLine}");
        }
        catch { /* логирование не должно ронять приложение */ }
    }

    public static List<OneCRealization> LoadRealizations(
        OneCConnectionSettings s, DateTime from, DateTime to)
    {
        dynamic? conn = null;
        dynamic? connector = null;
        var result     = new List<OneCRealization>();
        var skipped    = 0;

        Log($"=== LoadRealizations start: {from:dd.MM.yyyy} – {to:dd.MM.yyyy} ===");
        Log($"Строка подключения: {s.ConnectionString}");

        try
        {
            Log("Создаём коннектор...");
            connector = CreateConnector();

            Log("Подключаемся...");
            conn = connector.Connect(s.ConnectionString);

            Log("Создаём запрос...");
            var query = conn.NewObject("Запрос");
            query.Текст = BuildQuery();
            query.УстановитьПараметр("НачалоПериода", from.Date);
            query.УстановитьПараметр("КонецПериода",  to.Date.AddDays(1).AddSeconds(-1));

            Log("Выполняем запрос...");
            var queryResult = query.Выполнить();
            var selection   = queryResult.Выбрать();
            Log("Запрос выполнен, читаем строки...");

            int row = 0;
            bool hasNext;
            while (true)
            {
                try { hasNext = (bool)selection.Следующий(); }
                catch (Exception ex)
                {
                    Log($"Ошибка при вызове Следующий() на строке {row}: {ex}");
                    throw;
                }
                if (!hasNext) break;
                row++;

                try
                {
                    // Скалярные поля — строки и даты, никаких COM-объектов
                    var docNumber = Str(selection.НомерДок);
                    var docDate   = ToDateTime(selection.Дата);
                    var orderNum  = Str(selection.НомерЗаказа);
                    var orderDate = ToDateTime(selection.ДатаЗаказа);

                    var checkDt  = ToDateTime(selection.ДатаПечатиЧека);
                    var checkNum = Str(selection.НомерЧекаККМ);
                    var hasCheck = !string.IsNullOrEmpty(checkNum)
                                && checkDt > new DateTime(2000, 1, 1);

                    var dogovor   = Str(selection.Договор);
                    var isService = dogovor.IndexOf("агент", StringComparison.OrdinalIgnoreCase) >= 0;

                    result.Add(new OneCRealization
                    {
                        DocNumber    = docNumber,
                        DocDate      = docDate > DateTime.MinValue
                                        ? docDate.ToString("dd.MM.yyyy")
                                        : string.Empty,
                        OrderNumber  = orderNum,
                        OrderDate    = orderDate > DateTime.MinValue
                                        ? orderDate.ToString("dd.MM.yyyy HH:mm:ss")
                                        : string.Empty,
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
                    });
                }
                catch (Exception rowEx)
                {
                    skipped++;
                    Log($"Строка {row} пропущена: {rowEx.GetType().Name}: {rowEx.Message}{Environment.NewLine}{rowEx.StackTrace}");
                }
            }

            Log($"Готово: загружено {result.Count}, пропущено {skipped}");
        }
        catch (Exception ex)
        {
            Log($"КРИТИЧЕСКАЯ ОШИБКА: {ex.GetType().Name}: {ex.Message}{Environment.NewLine}{ex.StackTrace}");
            throw;
        }
        finally
        {
            if (conn      is not null) Marshal.ReleaseComObject(conn);
            if (connector is not null) Marshal.ReleaseComObject(connector);
        }

        return result;
    }

    /// <summary>
    /// Загружает табличную часть (Товары или Услуги) документа реализации по номеру документа.
    /// </summary>
    public static List<OneCRealizationItem> LoadRealizationItems(
        OneCConnectionSettings s, string docNumber, bool isService)
    {
        dynamic? conn = null;
        dynamic? connector = null;
        var result = new List<OneCRealizationItem>();

        Log($"=== LoadRealizationItems: docNumber={docNumber}, isService={isService} ===");

        try
        {
            connector = CreateConnector();
            conn      = connector.Connect(s.ConnectionString);

            var tableName = isService ? "Услуги" : "Товары";
            var query     = conn.NewObject("Запрос");
            query.Текст = $"""
                ВЫБРАТЬ
                    Строки.Номенклатура.Наименование КАК Наименование,
                    Строки.Количество                КАК Количество,
                    Строки.Сумма                     КАК Сумма
                ИЗ
                    Документ.РеализацияТоваровУслуг.{tableName} КАК Строки
                ГДЕ
                    Строки.Ссылка.Номер = &НомерДок
                """;
            query.УстановитьПараметр("НомерДок", docNumber);

            var queryResult = query.Выполнить();
            var selection   = queryResult.Выбрать();

            while ((bool)selection.Следующий())
            {
                result.Add(new OneCRealizationItem
                {
                    Name     = Str(selection.Наименование),
                    Quantity = ToDouble(selection.Количество),
                    Sum      = ToDouble(selection.Сумма),
                });
            }

            Log($"LoadRealizationItems: загружено {result.Count} позиций");
        }
        catch (Exception ex)
        {
            Log($"LoadRealizationItems ERROR: {ex.GetType().Name}: {ex.Message}");
            throw;
        }
        finally
        {
            if (conn      is not null) Marshal.ReleaseComObject(conn);
            if (connector is not null) Marshal.ReleaseComObject(connector);
        }

        return result;
    }

    /// <summary>
    /// Обогащает список заказов из текста данными из 1С:
    /// Подразделение (город) → IsService (из договора) → AgentInfo (из AppConstants).
    /// Запрашивает каждый заказ из Документ.ЗаказПокупателя.
    /// </summary>
    public static void EnrichOrdersFromOneC(
        OneCConnectionSettings s, List<Models.OrderEntry> orders)
    {
        if (orders.Count == 0) return;

        dynamic? conn      = null;
        dynamic? connector = null;

        Log($"=== EnrichOrdersFromOneC: {orders.Count} заказов ===");
        try
        {
            connector = CreateConnector();
            conn      = connector.Connect(s.ConnectionString);

            foreach (var order in orders)
            {
                // Пропускаем если агент уже определён (например из текста)
                if (order.AgentInfo is not null) continue;

                try
                {
                    var query = conn.NewObject("Запрос");
                    query.Текст = """
                        ВЫБРАТЬ ПЕРВЫЕ 1
                            Заказ.Подразделение.Наименование КАК Подразделение,
                            Заказ.ДоговорКонтрагента.Наименование КАК Договор,
                            Заказ.КонтактноеЛицоКонтрагента.Наименование КАК Покупатель
                        ИЗ
                            Документ.ЗаказПокупателя КАК Заказ
                        ГДЕ
                            Заказ.Номер = &НомерЗаказа
                            И Заказ.ПометкаУдаления = ЛОЖЬ
                        """;
                    query.УстановитьПараметр("НомерЗаказа", order.OrderNum);

                    var result    = query.Выполнить();
                    var selection = result.Выбрать();
                    if (!(bool)selection.Следующий()) continue;

                    var city      = Str(selection.Подразделение);
                    var dogovor   = Str(selection.Договор);
                    var customer  = Str(selection.Покупатель);

                    var isService = dogovor.IndexOf("агент", StringComparison.OrdinalIgnoreCase) >= 0;

                    if (!string.IsNullOrEmpty(city))     order.City = city;
                    if (isService)                        order.IsService = true;
                    if (string.IsNullOrEmpty(order.CustomerName) && !string.IsNullOrEmpty(customer))
                        order.CustomerName = customer;

                    // Ищем поставщика по городу + типу услуги
                    if (order.IsService && !string.IsNullOrEmpty(order.City))
                    {
                        // Определяем тип: (сборка) / (доставка) — берём из уже распознанных данных
                        // или пробуем оба варианта по очерёдности
                        var svcTypes = new[] { "Сборка", "Доставка" };
                        foreach (var svcType in svcTypes)
                        {
                            var agent = AtolGenerator.Constants.AppConstants.ServiceProviders
                                .FirstOrDefault(p =>
                                    string.Equals(p.Service, svcType, StringComparison.OrdinalIgnoreCase) &&
                                    order.City.Contains(p.City, StringComparison.OrdinalIgnoreCase));
                            if (agent is not null)
                            {
                                order.AgentInfo = agent;
                                Log($"  {order.OrderNum}: город={order.City}, агент={agent.Name}");
                                break;
                            }
                        }
                        if (order.AgentInfo is null)
                            Log($"  {order.OrderNum}: город={order.City} — агент не найден в списке");
                    }
                }
                catch (Exception ex)
                {
                    Log($"  {order.OrderNum}: ошибка запроса — {ex.Message}");
                }
            }
        }
        catch (Exception ex)
        {
            Log($"EnrichOrdersFromOneC ERROR: {ex.Message}");
        }
        finally
        {
            if (conn      is not null) Marshal.ReleaseComObject(conn);
            if (connector is not null) Marshal.ReleaseComObject(connector);
        }

        Log($"EnrichOrdersFromOneC done");
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
    // Используем скалярные атрибуты (строки/числа/даты) вместо ПРЕДСТАВЛЕНИЕ(),
    // чтобы избежать ошибки COM NullReferenceException при выполнении запроса.
    private static string BuildQuery() => """
        ВЫБРАТЬ
            РеализацияТоваровУслуг.Номер                                            КАК НомерДок,
            РеализацияТоваровУслуг.Дата                                             КАК Дата,
            РеализацияТоваровУслуг.Сделка.Номер                                     КАК НомерЗаказа,
            РеализацияТоваровУслуг.Сделка.Дата                                      КАК ДатаЗаказа,
            РеализацияТоваровУслуг.Сделка.КонтактноеЛицоКонтрагента.Наименование   КАК Покупатель,
            РеализацияТоваровУслуг.СуммаДокумента                                   КАК СуммаДокумента,
            РеализацияТоваровУслуг.ДоговорКонтрагента.Наименование                  КАК Договор,
            РеализацияТоваровУслуг.Подразделение.Наименование                       КАК Подразделение,
            РеализацияТоваровУслуг.НомерЧекаККМ                                     КАК НомерЧекаККМ,
            РеализацияТоваровУслуг.ЧекНомерФП                                       КАК ЧекНомерФП,
            РеализацияТоваровУслуг.ДатаПечатиЧека                                   КАК ДатаПечатиЧека
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
