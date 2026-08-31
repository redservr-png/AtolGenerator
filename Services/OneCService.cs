using System.IO;
using System.Runtime.InteropServices;
using System.Text.Json;
using AtolGenerator.Constants;
using AtolGenerator.Helpers;
using AtolGenerator.Models;

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
    public bool   IsOwnService   { get; set; }
    public string City           { get; set; } = string.Empty;
    public bool   HasCheck       { get; set; }  // чек уже пробит
    public string CheckNumber    { get; set; } = string.Empty;
    public string CheckDate      { get; set; } = string.Empty;
    public string FiscalNumber   { get; set; } = string.Empty;  // ЧекНомерФП
    public List<OneCRealizationItem> Items { get; set; } = new();
    public string ServiceType { get; set; } = string.Empty;
    public ServiceProvider? AgentInfo { get; set; }
}

public sealed class OneCRealizationEnrichmentError
{
    public string DocumentNumber { get; init; } = string.Empty;
    public string Message { get; init; } = string.Empty;
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
        dynamic? query = null;
        dynamic? queryResult = null;
        dynamic? selection = null;
        var result     = new List<OneCRealization>();
        var skipped    = 0;

        Log($"=== LoadRealizations start: {from:dd.MM.yyyy} – {to:dd.MM.yyyy} ===");
        Log($"Подключение: сервер={s.Server}, база={s.Database}, пользователь={s.User}");

        try
        {
            Log("Создаём коннектор...");
            connector = CreateConnector();

            Log("Подключаемся...");
            conn = connector.Connect(s.ConnectionString);

            Log("Создаём запрос...");
            query = conn.NewObject("Запрос");
            var usePropertyQuery = CorrectionPropertyRefs.TryResolve(
                conn, out CorrectionPropertyRefs? correctionProps, out List<string> missingProps);
            if (usePropertyQuery)
            {
                Log("Свойства коррекции найдены — запрос с регистром свойств.");
                query.Текст = BuildQueryWithProperties();
                BindCorrectionPropertyParameters(query, correctionProps!);
            }
            else
            {
                Log($"Свойства не найдены ({string.Join(", ", missingProps)}) — упрощённый запрос.");
                query.Текст = BuildQueryLegacy();
            }
            query.УстановитьПараметр("НачалоПериода", from.Date);
            query.УстановитьПараметр("КонецПериода",  to.Date.AddDays(1).AddSeconds(-1));

            Log("Выполняем запрос...");
            try
            {
                queryResult = query.Выполнить();
            }
            catch (Exception ex) when (usePropertyQuery)
            {
                Log($"Запрос со свойствами упал: {FormatComError(ex)} — повтор без регистра свойств.");
                ReleaseComObject(query);
                query = conn.NewObject("Запрос");
                query.Текст = BuildQueryLegacy();
                query.УстановитьПараметр("НачалоПериода", from.Date);
                query.УстановитьПараметр("КонецПериода",  to.Date.AddDays(1).AddSeconds(-1));
                usePropertyQuery = false;
                queryResult = query.Выполнить();
            }
            selection   = queryResult.Выбрать();
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

                    var effectiveCheckDt = usePropertyQuery
                        ? ToDateTime(selection.ДатаПечатиЧекаЭффективная)
                        : ToDateTime(selection.ДатаПечатиЧека);
                    if (!IsMeaningfulDate(effectiveCheckDt))
                    {
                        skipped++;
                        continue;
                    }

                    var checkNum = Str(selection.НомерЧекаККМ);
                    if (usePropertyQuery && string.IsNullOrWhiteSpace(checkNum))
                        checkNum = Str(selection.НомерЧекаККМСвойство);
                    var fiscalNumber = Str(selection.ЧекНомерФП);
                    if (usePropertyQuery && IsEmptyFp(fiscalNumber))
                        fiscalNumber = Str(selection.ЧекНомерФПСвойство);
                    var hasCheck = effectiveCheckDt > new DateTime(2000, 1, 1)
                                   && (!string.IsNullOrWhiteSpace(checkNum) || !IsEmptyFp(fiscalNumber));

                    var dogovor   = Str(selection.Договор);
                    var isService = dogovor.IndexOf("агент", StringComparison.OrdinalIgnoreCase) >= 0;

                    result.Add(new OneCRealization
                    {
                        DocNumber    = docNumber,
                        DocDate      = IsMeaningfulDate(docDate)
                                        ? docDate.ToString("dd.MM.yyyy")
                                        : string.Empty,
                        OrderNumber  = orderNum,
                        OrderDate    = IsMeaningfulDate(orderDate)
                                        ? orderDate.ToString("dd.MM.yyyy HH:mm:ss")
                                        : string.Empty,
                        CustomerName = Str(selection.Покупатель),
                        Amount       = ToDouble(selection.СуммаДокумента),
                        IsService    = isService,
                        City         = Str(selection.Подразделение),
                        HasCheck     = hasCheck,
                        CheckNumber  = checkNum,
                        FiscalNumber = fiscalNumber,
                        CheckDate    = hasCheck
                                        ? effectiveCheckDt.ToString("dd.MM.yyyy HH:mm:ss")
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
            ReleaseComObject(selection);
            ReleaseComObject(queryResult);
            ReleaseComObject(query);
            ReleaseComObject(conn);
            ReleaseComObject(connector);
        }

        return result;
    }

    /// <summary>
    /// Загружает табличную часть (Товары или Услуги) документа реализации по номеру документа.
    /// </summary>
    public static List<OneCRealizationItem> LoadRealizationItems(
        OneCConnectionSettings s, string docNumber, DateTime docDate, bool isService,
        object? existingConnection = null)
    {
        var ownsConnection = existingConnection is null;
        dynamic? conn = existingConnection;
        dynamic? connector = null;
        dynamic? query = null;
        dynamic? queryResult = null;
        dynamic? selection = null;
        var result = new List<OneCRealizationItem>();

        Log($"=== LoadRealizationItems: docNumber={docNumber}, docDate={docDate:dd.MM.yyyy}, isService={isService} ===");

        try
        {
            if (ownsConnection)
            {
                connector = CreateConnector();
                conn = connector.Connect(s.ConnectionString);
            }

            if (conn is null)
                throw new InvalidOperationException("1С вернула пустое COM-соединение.");

            query = conn.NewObject("Запрос");
            query.Текст = """
                ВЫБРАТЬ
                    Строки.Номенклатура.Наименование КАК Наименование,
                    Строки.Количество                КАК Количество,
                    Строки.Сумма                     КАК Сумма
                ИЗ
                    Документ.РеализацияТоваровУслуг.Товары КАК Строки
                ГДЕ
                    Строки.Ссылка.Номер = &НомерДок
                    И Строки.Ссылка.Дата >= &НачалоДня
                    И Строки.Ссылка.Дата < &КонецДня

                ОБЪЕДИНИТЬ ВСЕ

                ВЫБРАТЬ
                    Строки.Номенклатура.Наименование КАК Наименование,
                    Строки.Количество                КАК Количество,
                    Строки.Сумма                     КАК Сумма
                ИЗ
                    Документ.РеализацияТоваровУслуг.Услуги КАК Строки
                ГДЕ
                    Строки.Ссылка.Номер = &НомерДок
                    И Строки.Ссылка.Дата >= &НачалоДня
                    И Строки.Ссылка.Дата < &КонецДня
                """;
            query.УстановитьПараметр("НомерДок", docNumber);
            query.УстановитьПараметр("НачалоДня", docDate.Date);
            query.УстановитьПараметр("КонецДня", docDate.Date.AddDays(1));

            queryResult = query.Выполнить();
            selection   = queryResult.Выбрать();

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
            ReleaseComObject(selection);
            ReleaseComObject(queryResult);
            ReleaseComObject(query);
            if (ownsConnection)
            {
                ReleaseComObject(conn);
                ReleaseComObject(connector);
            }
        }

        return result;
    }

    public static void EnrichRealizationForReceipt(
        OneCConnectionSettings s, OneCRealization realization)
        => EnrichRealizationForReceipt(s, realization, null);

    private static void EnrichRealizationForReceipt(
        OneCConnectionSettings s, OneCRealization realization, object? existingConnection)
    {
        if (realization.Items.Count == 0 && !string.IsNullOrWhiteSpace(realization.DocNumber))
        {
            if (!DateTime.TryParseExact(realization.DocDate, "dd.MM.yyyy",
                    System.Globalization.CultureInfo.InvariantCulture,
                    System.Globalization.DateTimeStyles.None, out var docDate))
                throw new InvalidOperationException(
                    $"{realization.DocNumber}: не определена дата реализации для загрузки табличной части.");

            realization.Items = LoadRealizationItems(
                s, realization.DocNumber, docDate, realization.IsService, existingConnection);
        }

        if (string.IsNullOrWhiteSpace(realization.ServiceType))
            realization.ServiceType = DetectServiceType(realization.Items.Select(i => i.Name));

        if (ServiceClassificationService.ApplyOwnDeliveryRule(realization))
            return;

        if (realization.IsService && realization.AgentInfo is null)
            realization.AgentInfo = ResolveServiceProvider(realization.City, realization.ServiceType);
    }

    public static List<OneCRealizationEnrichmentError> EnrichRealizationsForReceipt(
        OneCConnectionSettings settings,
        IReadOnlyCollection<OneCRealization> realizations)
    {
        var errors = new List<OneCRealizationEnrichmentError>();
        if (realizations.Count == 0) return errors;

        dynamic? connector = null;
        dynamic? connection = null;
        try
        {
            Log($"=== EnrichRealizationsForReceipt: {realizations.Count} реализаций ===");
            connector = CreateConnector();
            connection = connector.Connect(settings.ConnectionString);

            foreach (var realization in realizations)
            {
                try
                {
                    EnrichRealizationForReceipt(settings, realization, (object)connection);
                }
                catch (Exception ex)
                {
                    var message = FormatComError(ex);
                    errors.Add(new OneCRealizationEnrichmentError
                    {
                        DocumentNumber = realization.DocNumber,
                        Message = message,
                    });
                    Log($"  {realization.DocNumber}: ошибка загрузки номенклатуры — {message}");
                }
            }
        }
        finally
        {
            ReleaseComObject(connection);
            ReleaseComObject(connector);
        }

        Log($"EnrichRealizationsForReceipt done: ошибок {errors.Count}");
        return errors;
    }

    public static ServiceProvider? ResolveServiceProvider(string city, string serviceType)
    {
        if (string.IsNullOrWhiteSpace(city)) return null;

        var cityMatches = AppConstants.ServiceProviders
            .Where(p => city.Contains(p.City, StringComparison.OrdinalIgnoreCase)
                     || p.City.Contains(city, StringComparison.OrdinalIgnoreCase))
            .ToList();
        if (cityMatches.Count == 0) return null;

        if (!string.IsNullOrWhiteSpace(serviceType))
        {
            var byType = cityMatches.FirstOrDefault(p =>
                string.Equals(p.Service, serviceType, StringComparison.OrdinalIgnoreCase));
            if (byType is not null) return byType;
        }

        return cityMatches.Count == 1 ? cityMatches[0] : null;
    }

    public static string DetectServiceType(IEnumerable<string> itemNames)
    {
        foreach (var rawName in itemNames)
        {
            var name = rawName ?? string.Empty;
            if (name.Contains("достав", StringComparison.OrdinalIgnoreCase)
             || name.Contains("перевоз", StringComparison.OrdinalIgnoreCase))
                return "Доставка";
            if (name.Contains("сбор", StringComparison.OrdinalIgnoreCase)
             || name.Contains("монтаж", StringComparison.OrdinalIgnoreCase))
                return "Сборка";
        }

        return string.Empty;
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

                    if (!string.IsNullOrEmpty(city))
                        order.City = city;
                    if (string.IsNullOrEmpty(order.CustomerName) && !string.IsNullOrEmpty(customer))
                        order.CustomerName = customer;
                    // IsService не меняем — метод вызывается только для уже помеченных услуг

                    if (ServiceClassificationService.IsOwnDeliveryDepartmentName(order.City))
                    {
                        try
                        {
                            var ownServiceItems = LoadBuyerOrderItems(conn, order.OrderNum);
                            if (order.Items.Count == 0) order.Items = ownServiceItems;
                            if (ServiceClassificationService.ApplyOwnDeliveryRule(order))
                            {
                                Log($"  {order.OrderNum}: собственная доставка России, НДС 22%, без агента");
                                continue;
                            }
                        }
                        catch (Exception itemEx)
                        {
                            Log($"  {order.OrderNum}: не удалось проверить номенклатуру собственной доставки — {itemEx.Message}");
                        }
                    }

                    // Ищем поставщика по городу + типу услуги
                    if (order.IsService && !string.IsNullOrEmpty(order.City))
                    {
                        order.AgentInfo = ResolveServiceProvider(order.City, order.ServiceType);
                        if (order.AgentInfo is not null)
                            Log($"  {order.OrderNum}: город={order.City}, агент={order.AgentInfo.Name}");
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

    public class ApplyResult
    {
        public int Total       { get; set; }
        public int Updated     { get; set; }
        public int Skipped     { get; set; }
        public int Failed      { get; set; }
        public List<string> Errors         { get; set; } = new();
        public List<string> SkippedSamples { get; set; } = new();  // первые N пропусков с подробностями
        public string  CsvBackupPath       { get; set; } = string.Empty;   // путь к CSV для ручного импорта
    }

    public class PunchedRecord
    {
        public string RealizationNum { get; set; } = string.Empty;
        public long?  FiscalDoc      { get; set; }
        public long?  FiscalSign     { get; set; }
        public string ReceiptDt      { get; set; } = string.Empty;
    }

    /// <summary>
    /// Читает Excel-отчёт ОФД (Сводный отчёт по фискальным документам Такском),
    /// для каждой строки извлекает: № реализации (тег 1086, колонка «Значение
    /// дополнительного реквизита пользователя»), ФПД, № ФД, дату чека.
    /// </summary>
    public static List<PunchedRecord> ReadOfdReport(string ofdReportPath)
    {
        var records = new List<PunchedRecord>();
        using var wb = new ClosedXML.Excel.XLWorkbook(ofdReportPath);
        var ws = wb.Worksheets.First();

        const int headerRow   = 11;
        const int firstDataRow = 12;
        var lastRow = ws.LastRowUsed()?.RowNumber() ?? headerRow;

        // Находим колонки по заголовку (порядок может отличаться)
        var colMap = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
        for (int c = 1; c <= ws.LastColumnUsed()!.ColumnNumber(); c++)
        {
            var name = ws.Cell(headerRow, c).GetString().Trim();
            if (!string.IsNullOrEmpty(name)) colMap[name] = c;
        }

        int colDate    = colMap.GetValueOrDefault("Дата и время", 1);
        int colFd      = colMap.GetValueOrDefault("№ ФД",        28);
        int colFp      = colMap.GetValueOrDefault("ФПД",         29);
        int colUserVal = colMap.GetValueOrDefault(
            "Значение дополнительного реквизита пользователя", 44);
        int colUserName = colMap.GetValueOrDefault(
            "Наименование дополнительного реквизита пользователя", 43);

        for (int r = firstDataRow; r <= lastRow; r++)
        {
            var userVal = ws.Cell(r, colUserVal).GetString().Trim();
            if (string.IsNullOrEmpty(userVal)) continue;

            // Парсим ФПД и № ФД
            var fpStr = ws.Cell(r, colFp).GetString().Trim();
            var fdStr = ws.Cell(r, colFd).GetString().Trim();
            if (!long.TryParse(fpStr, out var fp) || !long.TryParse(fdStr, out var fd))
                continue;

            // Дата
            var dateCell = ws.Cell(r, colDate);
            string dateStr;
            try
            {
                if (dateCell.DataType == ClosedXML.Excel.XLDataType.DateTime)
                    dateStr = dateCell.GetDateTime().ToString("dd.MM.yyyy HH:mm:ss");
                else
                    dateStr = dateCell.GetString().Trim();
            }
            catch { dateStr = dateCell.GetString().Trim(); }

            records.Add(new PunchedRecord
            {
                RealizationNum = userVal,
                FiscalDoc      = fd,
                FiscalSign     = fp,
                ReceiptDt      = dateStr,
            });
        }

        return records;
    }

    /// <summary>
    /// Применяет данные из списка пробитых чеков к документам РеализацияТоваровУслуг в 1С.
    /// Реквизиты: ЧекНомерФП (ФПД), НомерЧекаККМ (№ ФД), ДатаПечатиЧека.
    /// skipFilled=true (по умолчанию) — пропускать документы, у которых ЧекНомерФП уже непустой.
    /// </summary>
    public static ApplyResult ApplyPunchedChecks(
        OneCConnectionSettings s, List<PunchedRecord> records, bool skipFilled = true)
    {
        var res = new ApplyResult { Total = records.Count };
        if (records.Count == 0) return res;

        dynamic? conn      = null;
        dynamic? connector = null;

        Log($"=== ApplyPunchedChecks: {records.Count} записей ===");
        dynamic? docsManager = null;
        try
        {
            connector = CreateConnector();
            conn      = connector.Connect(s.ConnectionString);
            var correctionProps = CorrectionPropertyRefs.Resolve(conn);
            // Получаем менеджер документа один раз — будем использовать его ПолучитьОбъект(Ссылка)
            docsManager = conn.Документы.РеализацияТоваровУслуг;

            foreach (var rec in records)
            {
                if (string.IsNullOrEmpty(rec.RealizationNum) ||
                    rec.FiscalDoc is null || rec.FiscalSign is null)
                {
                    res.Skipped++;
                    continue;
                }

                string lastStep = "init";
                try
                {
                    // Дата чека (граница поиска документа: реализация должна быть НЕ ПОЗЖЕ даты чека)
                    DateTime checkDate = DateTime.Now;
                    if (!string.IsNullOrEmpty(rec.ReceiptDt))
                    {
                        if (DateTime.TryParseExact(rec.ReceiptDt, "dd.MM.yyyy HH:mm:ss",
                                System.Globalization.CultureInfo.InvariantCulture,
                                System.Globalization.DateTimeStyles.None, out var parsed)
                         || DateTime.TryParse(rec.ReceiptDt, out parsed))
                        {
                            checkDate = parsed;
                        }
                    }

                    // 1. Находим документ запросом.
                    //    Номера документов в УТ 10.3 — годовая нумерация: один номер может
                    //    встречаться в разных годах. Поэтому фильтруем по дате (документ
                    //    должен быть НЕ ПОЗЖЕ даты пробитого чека) и берём самый свежий.
                    var query = conn.NewObject("Запрос");
                    query.Текст = """
                        ВЫБРАТЬ ПЕРВЫЕ 1
                            Док.Ссылка       КАК ДокСсылка,
                            Док.Дата         КАК ДатаДок,
                            Док.ЧекНомерФП   КАК ЧекНомерФП
                        ИЗ
                            Документ.РеализацияТоваровУслуг КАК Док
                        ГДЕ
                            Док.Номер = &НомерДок
                            И Док.ПометкаУдаления = ЛОЖЬ
                            И Док.Дата <= &ДатаЧека
                        УПОРЯДОЧИТЬ ПО
                            Док.Дата УБЫВ
                        """;
                    query.УстановитьПараметр("НомерДок", rec.RealizationNum);
                    query.УстановитьПараметр("ДатаЧека", checkDate);
                    var qResult = query.Выполнить();
                    var sel     = qResult.Выбрать();
                    if (!(bool)sel.Следующий())
                    {
                        res.Failed++;
                        var msg = $"{rec.RealizationNum}: документ не найден (до {checkDate:dd.MM.yyyy})";
                        res.Errors.Add(msg);
                        Log("  " + msg);
                        continue;
                    }

                    var docDate         = ToDateTime(sel.ДатаДок);

                    // Получаем «сырое» значение ЧекНомерФП с типом для диагностики
                    dynamic rawFp = sel.ЧекНомерФП;
                    string  fpTypeName = "null";
                    string  fpRaw      = string.Empty;
                    try
                    {
                        if (rawFp is not null)
                        {
                            fpTypeName = ((object)rawFp).GetType().FullName ?? "?";
                            fpRaw      = (rawFp.ToString() ?? string.Empty).Trim();
                        }
                    }
                    catch { /* игнорируем — оставим пустое */ }

                    // 2. Проверяем skipFilled — поле считается заполненным, если значение
                    //    не входит в список «пустых» представлений
                    bool isFilled = !IsEmptyFp(fpRaw);
                    if (skipFilled && isFilled)
                    {
                        res.Skipped++;
                        var detail = $"{rec.RealizationNum}: дата={docDate:dd.MM.yyyy} ЧекНомерФП[{fpTypeName}] = «{fpRaw}»";
                        Log($"  {detail} — пропуск");
                        if (res.SkippedSamples.Count < 15) res.SkippedSamples.Add(detail);
                        continue;
                    }

                    Log($"  {rec.RealizationNum}: дата={docDate:dd.MM.yyyy} текущ.ФП[{fpTypeName}]=«{fpRaw}» → пишем ФПД={rec.FiscalSign}");

                    // 3. Получаем объект через ссылку и пишем реквизиты — каждый шаг в try-catch
                    //    для точной диагностики где падает.
                    lastStep = "sel.ДокСсылка";
                    var docRef = sel.ДокСсылка;
                    if (docRef is null)
                    {
                        res.Failed++;
                        var msg = $"{rec.RealizationNum}: ссылка пустая (sel.ДокСсылка == null)";
                        res.Errors.Add(msg);
                        Log("  " + msg);
                        continue;
                    }

                    // Освобождаем курсор запроса ДО ПолучитьОбъект — на случай, если
                    // удерживаемый курсор мешает платформе захватить блокировку документа.
                    try { System.Runtime.InteropServices.Marshal.ReleaseComObject(sel); } catch { }
                    try { System.Runtime.InteropServices.Marshal.ReleaseComObject(qResult); } catch { }
                    try { System.Runtime.InteropServices.Marshal.ReleaseComObject(query); } catch { }

                    // ПолучитьОбъект — две попытки:
                    //   1) свежая ссылка через mgr.НайтиПоНомеру → .ПолучитьОбъект()
                    //   2) исходная docRef.ПолучитьОбъект()
                    dynamic? obj = null;
                    string failReasons = string.Empty;

                    try
                    {
                        lastStep = "mgr.НайтиПоНомеру → ПолучитьОбъект()";
                        var freshRef = docsManager!.НайтиПоНомеру(rec.RealizationNum, checkDate);
                        if (freshRef is not null && !(bool)freshRef.Пустая())
                            obj = freshRef.ПолучитьОбъект();
                    }
                    catch (Exception ex1) { failReasons += $"[fresh: {ex1.Message}] "; }

                    if (obj is null)
                    {
                        try
                        {
                            lastStep = "docRef.ПолучитьОбъект()";
                            obj = docRef.ПолучитьОбъект();
                        }
                        catch (Exception ex2) { failReasons += $"[ref: {ex2.Message}] "; }
                    }

                    if (obj is null)
                    {
                        res.Failed++;
                        var msg = $"{rec.RealizationNum}: оба способа ПолучитьОбъект упали. {failReasons}";
                        res.Errors.Add(msg);
                        Log("  ОШИБКА " + msg);
                        continue;
                    }


                    // Пробуем писать ЧИСЛОВЫЕ значения (а не строки) — поле ЧекНомерФП в УТ
                    // 10.3 имеет тип Число (видно по значениям типа 3155950491 в логе)
                    lastStep = "set obj.ЧекНомерФП";
                    obj.ЧекНомерФП = (double)rec.FiscalSign.Value;

                    lastStep = "set obj.НомерЧекаККМ";
                    obj.НомерЧекаККМ = (double)rec.FiscalDoc.Value;

                    lastStep = "set obj.ДатаПечатиЧека";
                    obj.ДатаПечатиЧека = checkDate;

                    lastStep = "obj.Записать()";
                    obj.Записать();

                    lastStep = "write properties";
                    string correctionComment = ReadPropertyString(conn, docRef, correctionProps.Comment);
                    var comment = $"{checkDate:dd.MM.yyyy} Пробит чек коррекции \"Приход\" ФП: {rec.FiscalSign!.Value}";
                    WriteCheckPropertiesBundle(
                        conn,
                        docRef,
                        correctionProps,
                        comment,
                        checkDate,
                        rec.FiscalDoc,
                        rec.FiscalSign,
                        ref correctionComment,
                        appendComment: true);

                    res.Updated++;
                    Log($"  {rec.RealizationNum}: дата={docDate:dd.MM.yyyy} ФПД={rec.FiscalSign} №ФД={rec.FiscalDoc} → реквизиты + свойства");

                    // Освобождаем COM-объект документа
                    try { System.Runtime.InteropServices.Marshal.ReleaseComObject(obj); } catch { }
                }
                catch (Exception ex)
                {
                    res.Failed++;
                    var msg = $"{rec.RealizationNum} [шаг: {lastStep}]: {ex.GetType().Name}: {ex.Message}";
                    res.Errors.Add(msg);
                    Log($"  ОШИБКА {msg}\n{ex.StackTrace}");
                }
            }
        }
        catch (Exception ex)
        {
            Log($"ApplyPunchedChecks ERROR: {ex.Message}");
            res.Errors.Add(ex.Message);
        }
        finally
        {
            if (conn      is not null) Marshal.ReleaseComObject(conn);
            if (connector is not null) Marshal.ReleaseComObject(connector);
        }

        Log($"=== Применено: обновлено {res.Updated}, пропущено {res.Skipped}, ошибок {res.Failed} ===");

        // CSV-резерв для ручного импорта через внешнюю 1С-обработку (на случай COM-сбоев)
        // Пишем в Windows-1251 — родная кодировка УТ 10.3 (платформа 8.2), без BOM.
        try
        {
            // Регистрируем провайдер кодовых страниц (нужно в .NET Core+)
            System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);
            var enc = System.Text.Encoding.GetEncoding(1251);

            var csvDir  = AppDomain.CurrentDomain.BaseDirectory;
            var csvName = $"atol_to_1c_{DateTime.Now:yyyyMMdd_HHmmss}.csv";
            var csvPath = System.IO.Path.Combine(csvDir, csvName);
            using var sw = new System.IO.StreamWriter(csvPath, false, enc);
            sw.WriteLine("НомерРеализации;ФПД;НомерФД;ДатаПечатиЧека");
            foreach (var rec in records)
            {
                if (string.IsNullOrEmpty(rec.RealizationNum) ||
                    rec.FiscalDoc is null || rec.FiscalSign is null) continue;
                sw.WriteLine($"{rec.RealizationNum};{rec.FiscalSign};{rec.FiscalDoc};{rec.ReceiptDt}");
            }
            res.CsvBackupPath = csvPath;
            Log($"CSV для ручного импорта (Windows-1251): {csvPath}");
        }
        catch (Exception ex) { Log($"Ошибка записи CSV: {ex.Message}"); }

        return res;
    }

    /// <summary>
    /// Записывает в 1С результат сверки XML + АТОЛ (+ ОФД):
    /// <c>update_fields</c> — реквизиты чека на документе + свойства;
    /// <c>comment_only</c> — только свойства (по одному набору на каждый чек пары).
    /// В реквизит <c>Комментарий</c> документа ничего не пишется.
    /// </summary>
    public static ApplyResult ApplyOneCExportRows(
        OneCConnectionSettings s, IReadOnlyList<OneCExportRow> rows, bool skipFilled = true)
    {
        var ready = rows.Where(x => x.IsReady).ToList();
        var res = new ApplyResult { Total = ready.Count };
        if (ready.Count == 0) return res;

        var groups = ready
            .GroupBy(x => x.RealizationNumber, StringComparer.OrdinalIgnoreCase)
            .OrderBy(g => g.Key, StringComparer.OrdinalIgnoreCase)
            .ToList();
        res.Total = groups.Count;

        dynamic? conn = null;
        dynamic? connector = null;
        dynamic? docsManager = null;

        Log($"=== ApplyOneCExportRows: {ready.Count} строк, {groups.Count} реализаций ===");
        try
        {
            connector = CreateConnector();
            conn = connector.Connect(s.ConnectionString);
            docsManager = conn.Документы.РеализацияТоваровУслуг;
            var correctionProps = CorrectionPropertyRefs.Resolve(conn);

            foreach (var group in groups)
            {
                var realizationNum = group.Key;
                var ordered = group
                    .OrderBy(x => x.WriteMode == "comment_only" && x.CheckType == "sell_refund" ? 0 : 1)
                    .ThenBy(x => x.RegisteredAt ?? DateTime.MinValue)
                    .ToList();
                var checkDate = ordered
                    .Select(x => x.RegisteredAt)
                    .Where(x => x.HasValue)
                    .Select(x => x!.Value)
                    .DefaultIfEmpty(DateTime.Now)
                    .Max();
                var updateRow = ordered.FirstOrDefault(x =>
                    string.Equals(x.WriteMode, "update_fields", StringComparison.OrdinalIgnoreCase) &&
                    x.FiscalSign.HasValue && x.FiscalDocument.HasValue);
                var commentOnly = updateRow is null;
                string lastStep = "init";

                try
                {
                    lastStep = "query";
                    var query = conn.NewObject("Запрос");
                    query.Текст = """
                        ВЫБРАТЬ ПЕРВЫЕ 1
                            Док.Ссылка       КАК ДокСсылка,
                            Док.Дата         КАК ДатаДок,
                            Док.ЧекНомерФП   КАК ЧекНомерФП
                        ИЗ
                            Документ.РеализацияТоваровУслуг КАК Док
                        ГДЕ
                            Док.Номер = &НомерДок
                            И Док.ПометкаУдаления = ЛОЖЬ
                            И Док.Дата <= &ДатаЧека
                        УПОРЯДОЧИТЬ ПО
                            Док.Дата УБЫВ
                        """;
                    query.УстановитьПараметр("НомерДок", realizationNum);
                    query.УстановитьПараметр("ДатаЧека", checkDate);
                    var qResult = query.Выполнить();
                    var sel = qResult.Выбрать();
                    if (!(bool)sel.Следующий())
                    {
                        res.Failed++;
                        var msg = $"{realizationNum}: документ не найден (до {checkDate:dd.MM.yyyy})";
                        res.Errors.Add(msg);
                        Log("  " + msg);
                        continue;
                    }

                    var docDate = ToDateTime(sel.ДатаДок);
                    dynamic rawFp = sel.ЧекНомерФП;
                    var fpRaw = string.Empty;
                    try
                    {
                        if (rawFp is not null)
                            fpRaw = (rawFp.ToString() ?? string.Empty).Trim();
                    }
                    catch { /* ignore */ }

                    var isFilled = !IsEmptyFp(fpRaw);
                    if (!commentOnly && skipFilled && isFilled)
                        Log($"  {realizationNum}: ЧекНомерФП уже «{fpRaw}» — реквизиты не трогаем, свойства проверим");

                    lastStep = "sel.ДокСсылка";
                    var docRef = sel.ДокСсылка;
                    try { Marshal.ReleaseComObject(sel); } catch { }
                    try { Marshal.ReleaseComObject(qResult); } catch { }
                    try { Marshal.ReleaseComObject(query); } catch { }

                    string correctionComment = ReadPropertyString(conn, docRef, correctionProps.Comment);
                    var changed = false;

                    if (commentOnly)
                    {
                        lastStep = "write properties (comment_only)";
                        foreach (var row in ordered)
                        {
                            if (WriteCheckPropertiesBundle(
                                    conn,
                                    docRef,
                                    correctionProps,
                                    row.Comment,
                                    row.RegisteredAt,
                                    row.FiscalDocument,
                                    row.FiscalSign,
                                    ref correctionComment,
                                    appendComment: true))
                                changed = true;
                        }

                        if (!changed)
                        {
                            res.Skipped++;
                            var detail = $"{realizationNum}: дата={docDate:dd.MM.yyyy} — свойства без изменений";
                            Log("  " + detail);
                            if (res.SkippedSamples.Count < 15) res.SkippedSamples.Add(detail);
                            continue;
                        }

                        res.Updated++;
                        Log($"  {realizationNum}: дата={docDate:dd.MM.yyyy} mode=comment_only → {ordered.Count} чек(ов) в свойства");
                        continue;
                    }

                    dynamic? obj = null;
                    var failReasons = string.Empty;
                    try
                    {
                        lastStep = "mgr.НайтиПоНомеру → ПолучитьОбъект()";
                        var freshRef = docsManager!.НайтиПоНомеру(realizationNum, checkDate);
                        if (freshRef is not null && !(bool)freshRef.Пустая())
                            obj = freshRef.ПолучитьОбъект();
                    }
                    catch (Exception ex1) { failReasons += $"[fresh: {ex1.Message}] "; }

                    if (obj is null)
                    {
                        try
                        {
                            lastStep = "docRef.ПолучитьОбъект()";
                            obj = docRef.ПолучитьОбъект();
                        }
                        catch (Exception ex2) { failReasons += $"[ref: {ex2.Message}] "; }
                    }

                    if (obj is null)
                    {
                        res.Failed++;
                        var msg = $"{realizationNum}: ПолучитьОбъект упал. {failReasons}";
                        res.Errors.Add(msg);
                        Log("  ОШИБКА " + msg);
                        continue;
                    }

                    if (updateRow is not null && !(skipFilled && isFilled))
                    {
                        lastStep = "set fields";
                        obj.ЧекНомерФП = (double)updateRow.FiscalSign!.Value;
                        obj.НомерЧекаККМ = (double)updateRow.FiscalDocument!.Value;
                        obj.ДатаПечатиЧека = updateRow.RegisteredAt ?? checkDate;
                        changed = true;
                    }

                    if (changed)
                    {
                        lastStep = "obj.Записать()";
                        obj.Записать();
                    }

                    try { Marshal.ReleaseComObject(obj); } catch { }

                    if (updateRow is not null)
                    {
                        lastStep = "write properties (update_fields)";
                        if (WriteCheckPropertiesBundle(
                                conn,
                                docRef,
                                correctionProps,
                                updateRow.Comment,
                                updateRow.RegisteredAt,
                                updateRow.FiscalDocument,
                                updateRow.FiscalSign,
                                ref correctionComment,
                                appendComment: true))
                            changed = true;
                    }

                    if (!changed)
                    {
                        res.Skipped++;
                        var detail = $"{realizationNum}: дата={docDate:dd.MM.yyyy} — без изменений";
                        Log("  " + detail);
                        if (res.SkippedSamples.Count < 15) res.SkippedSamples.Add(detail);
                        continue;
                    }

                    res.Updated++;
                    Log($"  {realizationNum}: дата={docDate:dd.MM.yyyy} mode=update_fields → реквизиты + свойства");
                }
                catch (Exception ex)
                {
                    res.Failed++;
                    var msg = $"{realizationNum} [шаг: {lastStep}]: {ex.GetType().Name}: {ex.Message}";
                    res.Errors.Add(msg);
                    Log($"  ОШИБКА {msg}\n{ex.StackTrace}");
                }
            }
        }
        catch (Exception ex)
        {
            Log($"ApplyOneCExportRows ERROR: {ex.Message}");
            res.Errors.Add(ex.Message);
        }
        finally
        {
            if (docsManager is not null) try { Marshal.ReleaseComObject(docsManager); } catch { }
            if (conn is not null) Marshal.ReleaseComObject(conn);
            if (connector is not null) Marshal.ReleaseComObject(connector);
        }

        Log($"=== ApplyOneCExportRows: обновлено {res.Updated}, пропущено {res.Skipped}, ошибок {res.Failed} ===");

        try
        {
            Directory.CreateDirectory(FileHelper.OutputDir);
            var csvPath = Path.Combine(
                FileHelper.OutputDir,
                $"atol_to_1c_{DateTime.Now:yyyyMMdd_HHmmss}.csv");
            ReportReconciliationService.ExportOneCCsv(csvPath, ready);
            res.CsvBackupPath = csvPath;
            Log($"CSV-резерв: {csvPath}");
        }
        catch (Exception ex) { Log($"Ошибка записи CSV-резерва: {ex.Message}"); }

        return res;
    }

    private static bool CommentAlreadyHasFiscalSign(string comment, long? fiscalSign)
    {
        if (!fiscalSign.HasValue || string.IsNullOrWhiteSpace(comment)) return false;
        var token = fiscalSign.Value.ToString(System.Globalization.CultureInfo.InvariantCulture);
        return comment.Contains($"ФП: {token}", StringComparison.OrdinalIgnoreCase);
    }

    public class FetchAmountResult
    {
        public int Total       { get; set; }
        public int Filled      { get; set; }
        public int FilledFp    { get; set; }   // сколько строк получило непустой ЧекНомерФП
        public int NotFound    { get; set; }
        public int Skipped     { get; set; }
        public List<string> Errors { get; set; } = new();
    }

    /// <summary>
    /// Универсальный добор суммы документа из 1С по списку записей.
    /// Для каждой записи находит документ по типу+номеру (с учётом годовой нумерации),
    /// записывает результат в OrderEntry.CorrectAmount и Amount (если Amount=0).
    ///
    /// OriginalCheckAmount НЕ перезаписываем — там может уже стоять ошибочная сумма
    /// введённая пользователем для сценариев DecreaseAmount/IncreaseAmount.
    /// </summary>
    public static FetchAmountResult FetchAmountsFromOneC(
        OneCConnectionSettings s, IEnumerable<Models.OrderEntry> orders)
    {
        var res = new FetchAmountResult();
        var list = orders.ToList();
        res.Total = list.Count;
        if (list.Count == 0) return res;

        dynamic? conn      = null;
        dynamic? connector = null;

        Log($"=== FetchAmountsFromOneC: {list.Count} записей ===");
        try
        {
            connector = CreateConnector();
            conn      = connector.Connect(s.ConnectionString);

            foreach (var o in list)
            {
                var metaName = DocumentMetaName(o.DocumentType);
                if (metaName is null || string.IsNullOrEmpty(o.OrderNum))
                {
                    res.Skipped++;
                    continue;
                }

                try
                {
                    // Номер документа в УТ 10.3 может повторяться между годами. Дата в
                    // Obsidian иногда отличается от реальной даты документа, поэтому она
                    // ограничивает поиск годом, а точную дату берём уже из найденной записи 1С.
                    DateTime? dateHint = null;
                    var lookupDate = string.IsNullOrWhiteSpace(o.SourceDocumentDate)
                        ? o.OrderDate
                        : o.SourceDocumentDate;
                    if (!string.IsNullOrEmpty(lookupDate))
                    {
                        var datePart = lookupDate.Split(' ').FirstOrDefault() ?? string.Empty;
                        if (DateTime.TryParseExact(datePart, "dd.MM.yyyy",
                                System.Globalization.CultureInfo.InvariantCulture,
                                System.Globalization.DateTimeStyles.None, out var d))
                            dateHint = d.Date;
                    }

                    // У большинства документов с фискалкой есть поле ЧекНомерФП.
                    // У Заказа покупателя его нет — для него выбираем только сумму.
                    bool hasFpField = HasFiscalNumberField(o.DocumentType);
                    var fpSelectPart = hasFpField ? ", Док.ЧекНомерФП КАК ЧекНомерФП" : "";
                    var checkSelectPart = hasFpField
                        ? ", Док.НомерЧекаККМ КАК НомерЧекаККМ, Док.ДатаПечатиЧека КАК ДатаПечатиЧека"
                        : "";
                    var detailsSelectPart = FetchDetailsSelectPart(o.DocumentType);
                    var dateFilter = dateHint.HasValue
                        ? "И Док.Дата >= &НачалоГода И Док.Дата < &КонецГода"
                        : string.Empty;

                    var query = conn.NewObject("Запрос");
                    query.Текст = $"""
                        ВЫБРАТЬ ПЕРВЫЕ 1
                            Док.СуммаДокумента  КАК Сумма,
                            Док.Дата             КАК ДатаДок,
                            Док.Комментарий      КАК Комментарий
                            {fpSelectPart}
                            {checkSelectPart}
                            {detailsSelectPart}
                        ИЗ
                            Документ.{metaName} КАК Док
                        ГДЕ
                            Док.Номер = &Номер
                            {dateFilter}
                        УПОРЯДОЧИТЬ ПО
                            Док.ПометкаУдаления,
                            Док.Дата УБЫВ
                        """;
                    query.УстановитьПараметр("Номер", o.OrderNum);
                    if (dateHint.HasValue)
                    {
                        var yearStart = new DateTime(dateHint.Value.Year, 1, 1);
                        query.УстановитьПараметр("НачалоГода", yearStart);
                        query.УстановитьПараметр("КонецГода", yearStart.AddYears(1));
                    }
                    var sel = query.Выполнить().Выбрать();

                    if (!(bool)sel.Следующий())
                    {
                        res.NotFound++;
                        Log(dateHint.HasValue
                            ? $"  {o.OrderNum} ({metaName}): не найден за {dateHint:yyyy} год"
                            : $"  {o.OrderNum} ({metaName}): не найден");
                        continue;
                    }

                    var sum = ToDouble(sel.Сумма);
                    var documentDate = ToDateTime(sel.ДатаДок);
                    o.OneCComment = Str(sel.Комментарий).Trim();
                    o.CorrectAmount = sum;
                    if (IsMeaningfulDate(documentDate))
                    {
                        o.OrderDate = documentDate.ToString("dd.MM.yyyy HH:mm:ss");
                        o.CorrectionDate = documentDate.ToString("dd.MM.yyyy");
                    }
                    // Если Amount был 0 (типично для строк из Obsidian без «суммы» в тексте) —
                    // подставляем правильную сумму как основную.
                    if (o.Amount <= 0) o.Amount = sum;

                    // Явная загрузка из 1С считается авторитетной: обновляем сохранённые
                    // реквизиты чека, если они заполнены в найденном документе.
                    if (hasFpField)
                    {
                        try
                        {
                            var fp = Str(sel.ЧекНомерФП).Trim();
                            if (!IsEmptyFp(fp))
                            {
                                o.OriginalFiscalNumber = fp;
                                res.FilledFp++;
                            }
                            o.OneCCheckNumber = Str(sel.НомерЧекаККМ).Trim();
                            var checkDate = ToDateTime(sel.ДатаПечатиЧека);
                            o.OneCCheckDate = IsMeaningfulDate(checkDate) ? checkDate : null;
                        }
                        catch { /* нет поля — пропускаем */ }
                    }

                    ApplyFetchedDetails(o, sel);

                    res.Filled++;
                    Log($"  {o.OrderNum} ({metaName}): сумма={sum}, ФП={o.OriginalFiscalNumber}");
                }
                catch (Exception ex)
                {
                    res.Errors.Add($"{o.OrderNum}: {ex.Message}");
                    Log($"  ОШИБКА {o.OrderNum}: {ex.Message}");
                }
            }
        }
        catch (Exception ex)
        {
            Log($"FetchAmountsFromOneC ERROR: {ex.Message}");
            res.Errors.Add(ex.Message);
        }
        finally
        {
            if (conn      is not null) Marshal.ReleaseComObject(conn);
            if (connector is not null) Marshal.ReleaseComObject(connector);
        }

        Log($"=== FetchAmounts: заполнено {res.Filled}, не найдено {res.NotFound}, пропущено {res.Skipped} ===");
        return res;
    }

    public static void EnrichCorrectionOrderForReceipt(
        OneCConnectionSettings settings,
        Models.OrderEntry order)
    {
        if (order.DocumentType != Models.SourceDocumentType.Realization)
            return;
        if (!DateTime.TryParseExact(order.OrderDate, "dd.MM.yyyy HH:mm:ss",
                System.Globalization.CultureInfo.InvariantCulture,
                System.Globalization.DateTimeStyles.None, out var documentDate) &&
            !DateTime.TryParseExact(order.OrderDate, "dd.MM.yyyy",
                System.Globalization.CultureInfo.InvariantCulture,
                System.Globalization.DateTimeStyles.None, out documentDate))
            return;

        var realization = new OneCRealization
        {
            DocNumber = order.OrderNum,
            DocDate = documentDate.ToString("dd.MM.yyyy"),
            Amount = order.CorrectAmount ?? order.Amount,
            IsService = order.IsService,
            IsOwnService = order.IsOwnService,
            City = order.City,
            CustomerName = order.CustomerName,
            ServiceType = order.ServiceType,
            AgentInfo = order.AgentInfo,
        };
        EnrichRealizationForReceipt(settings, realization);
        order.Items = realization.Items.Select(x => new Models.OrderItem
        {
            Name = x.Name,
            Quantity = x.Quantity,
            Sum = x.Sum,
        }).ToList();
        order.IsService = realization.IsService;
        order.IsOwnService = realization.IsOwnService;
        order.ServiceType = realization.ServiceType;
        order.AgentInfo = realization.AgentInfo;
    }

    private static List<Models.OrderItem> LoadBuyerOrderItems(dynamic connection, string orderNumber)
    {
        var query = connection.NewObject("Запрос");
        query.Текст = """
            ВЫБРАТЬ
                Строки.Номенклатура.Наименование КАК Наименование,
                Строки.Количество                КАК Количество,
                Строки.Сумма                     КАК Сумма
            ИЗ
                Документ.ЗаказПокупателя.Товары КАК Строки
            ГДЕ
                Строки.Ссылка.Номер = &НомерЗаказа
                И Строки.Ссылка.ПометкаУдаления = ЛОЖЬ
            """;
        query.УстановитьПараметр("НомерЗаказа", orderNumber);

        var result = new List<Models.OrderItem>();
        var selection = query.Выполнить().Выбрать();
        while ((bool)selection.Следующий())
        {
            result.Add(new Models.OrderItem
            {
                Name = Str(selection.Наименование),
                Quantity = ToDouble(selection.Количество),
                Sum = ToDouble(selection.Сумма),
            });
        }

        return result;
    }

    private static string FetchDetailsSelectPart(Models.SourceDocumentType type) => type switch
    {
        Models.SourceDocumentType.Realization => """
            , Док.Подразделение.Наименование КАК Подразделение
            , Док.ДоговорКонтрагента.Наименование КАК Договор
            , Док.Сделка.КонтактноеЛицоКонтрагента.Наименование КАК Покупатель
            """,
        Models.SourceDocumentType.BuyerOrder => """
            , Док.Подразделение.Наименование КАК Подразделение
            , Док.ДоговорКонтрагента.Наименование КАК Договор
            , Док.КонтактноеЛицоКонтрагента.Наименование КАК Покупатель
            """,
        _ => string.Empty,
    };

    private static void ApplyFetchedDetails(Models.OrderEntry order, dynamic selection)
    {
        if (order.DocumentType is not (Models.SourceDocumentType.Realization or Models.SourceDocumentType.BuyerOrder))
            return;

        var city = Str(selection.Подразделение);
        var contract = Str(selection.Договор);
        var customer = Str(selection.Покупатель);
        if (!string.IsNullOrWhiteSpace(city)) order.City = city;
        if (!string.IsNullOrWhiteSpace(customer)) order.CustomerName = customer;
        order.IsService = contract.Contains("агент", StringComparison.OrdinalIgnoreCase);
    }

    /// <summary>Маппинг SourceDocumentType → имя документа в метаданных 1С УТ 10.3.</summary>
    private static string? DocumentMetaName(Models.SourceDocumentType t) => t switch
    {
        Models.SourceDocumentType.Realization => "РеализацияТоваровУслуг",
        Models.SourceDocumentType.CardPayment => "ОплатаОтПокупателяПлатежнойКартой",
        Models.SourceDocumentType.CashPayment => "ПриходныйКассовыйОрдер",
        Models.SourceDocumentType.CashExpense => "РасходныйКассовыйОрдер",
        Models.SourceDocumentType.BuyerOrder  => "ЗаказПокупателя",
        // KkmCheck и FpOnly: либо нет документа в 1С с таким номером, либо неоднозначно — пропускаем
        _ => null,
    };

    /// <summary>Есть ли у документа реквизит ЧекНомерФП в УТ 10.3.</summary>
    private static bool HasFiscalNumberField(Models.SourceDocumentType t) => t switch
    {
        Models.SourceDocumentType.Realization => true,
        Models.SourceDocumentType.CardPayment => true,
        Models.SourceDocumentType.CashPayment => true,
        Models.SourceDocumentType.CashExpense => true,
        // Заказ покупателя — это не фискальный документ, ЧекНомерФП не имеет
        Models.SourceDocumentType.BuyerOrder  => false,
        _ => false,
    };

    /// <summary>
    /// Считает значение поля ЧекНомерФП «пустым» — учитывает разные представления,
    /// которые приходят через COM-мост и через групповую обработку 1С.
    /// </summary>
    private static bool IsEmptyFp(string s)
    {
        if (string.IsNullOrWhiteSpace(s)) return true;
        var v = s.Trim();
        // Различные представления нуля
        if (v == "0" || v == "0.0" || v == "0,0" || v == "0.00" || v == "0,00") return true;
        // Только нули (например, "000000000")
        if (v.All(c => c == '0')) return true;
        // Известные платформенные плейсхолдеры
        if (v == "999999999") return true;
        // 1С Null-маркеры
        if (string.Equals(v, "Неопределено", StringComparison.OrdinalIgnoreCase)) return true;
        if (string.Equals(v, "Null",         StringComparison.OrdinalIgnoreCase)) return true;
        // COM-обёртки (.NET без правильной строки)
        if (v.StartsWith("System.", StringComparison.Ordinal)) return true;
        return false;
    }

    private static dynamic CreateConnector()
    {
        var t = Type.GetTypeFromProgID("V83.COMConnector")
             ?? throw new InvalidOperationException("V83.COMConnector не найден. Установите клиент 1С.");
        return Activator.CreateInstance(t)
            ?? throw new InvalidOperationException("Не удалось создать экземпляр V83.COMConnector");
    }

    // Безопасные приведения COM-значений
    private static void ReleaseComObject(object? value)
    {
        if (value is null || !Marshal.IsComObject(value)) return;
        try { Marshal.ReleaseComObject(value); }
        catch { /* освобождение COM-объекта не должно прерывать основную операцию */ }
    }

    private static string FormatComError(Exception exception)
    {
        var error = exception;
        while (error.InnerException is not null)
            error = error.InnerException;

        return error is COMException com
            ? $"COM 0x{com.HResult:X8}: {com.Message}"
            : error.Message;
    }

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

    private static bool IsMeaningfulDate(DateTime value) =>
        value.Year is >= 2000 and <= 2100;

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

    // ── Свойства коррекции (ПВХ СвойстваОбъектов / ЧекиКоррекции) ─────────────
    private static class CorrectionPropertyNames
    {
        public const string Comment = "КомментарийКорректировки";
        public const string CheckDate = "ДатаПечатиЧека";
        public const string CheckNumber = "НомерЧекаККМ";
        public const string FiscalSign = "ЧекНомерФП";
    }

    private sealed class CorrectionPropertyRefs
    {
        public required dynamic Comment { get; init; }
        public required dynamic CheckDate { get; init; }
        public required dynamic CheckNumber { get; init; }
        public required dynamic FiscalSign { get; init; }

        public static CorrectionPropertyRefs Resolve(dynamic conn)
        {
            if (!TryResolve(conn, out CorrectionPropertyRefs? refs, out List<string> missing))
            {
                throw new InvalidOperationException(
                    "Не найдены свойства в ПВХ СвойстваОбъектов: " + string.Join(", ", missing));
            }

            return refs!;
        }

        public static bool TryResolve(
            dynamic conn, out CorrectionPropertyRefs? refs, out List<string> missing)
        {
            missing = new List<string>();
            refs = null;

            dynamic plan;
            try { plan = conn.ПланыВидовХарактеристик.СвойстваОбъектов; }
            catch (Exception ex)
            {
                Log($"ПВХ СвойстваОбъектов недоступен: {FormatComError(ex)}");
                return false;
            }

            var comment = FindPropertyRef(plan, CorrectionPropertyNames.Comment, missing);
            var checkDate = FindPropertyRef(plan, CorrectionPropertyNames.CheckDate, missing);
            var checkNumber = FindPropertyRef(plan, CorrectionPropertyNames.CheckNumber, missing);
            var fiscalSign = FindPropertyRef(plan, CorrectionPropertyNames.FiscalSign, missing);
            if (missing.Count > 0)
            {
                refs = null;
                return false;
            }

            refs = new CorrectionPropertyRefs
            {
                Comment = comment!,
                CheckDate = checkDate!,
                CheckNumber = checkNumber!,
                FiscalSign = fiscalSign!,
            };

            return true;
        }

        private static dynamic? FindPropertyRef(dynamic plan, string name, List<string> missing)
        {
            try
            {
                var reference = plan.НайтиПоНаименованию(name);
                if (IsEmptyOneCRef(reference))
                {
                    missing.Add(name);
                    return null;
                }

                return reference;
            }
            catch (Exception ex)
            {
                Log($"Свойство {name}: {FormatComError(ex)}");
                missing.Add(name);
                return null;
            }
        }
    }

    private static bool IsEmptyOneCRef(dynamic? reference)
    {
        if (reference is null) return true;
        try { return (bool)reference.Пустая(); }
        catch { return true; }
    }

    private static void BindCorrectionPropertyParameters(dynamic query, CorrectionPropertyRefs props)
    {
        query.УстановитьПараметр("СвойствоКомментарийКорректировки", props.Comment);
        query.УстановитьПараметр("СвойствоДатаПечатиЧека", props.CheckDate);
        query.УстановитьПараметр("СвойствоНомерЧекаККМ", props.CheckNumber);
        query.УстановитьПараметр("СвойствоЧекНомерФП", props.FiscalSign);
    }

    private static string ReadPropertyString(dynamic conn, dynamic docRef, dynamic propertyRef)
    {
        dynamic? query = null;
        dynamic? queryResult = null;
        dynamic? selection = null;
        try
        {
            query = conn.NewObject("Запрос");
            query.Текст = """
                ВЫБРАТЬ
                    Значения.Значение КАК Значение
                ИЗ
                    РегистрСведений.ЗначенияСвойствОбъектов КАК Значения
                ГДЕ
                    Значения.Объект = &Объект
                    И Значения.Свойство = &Свойство
                """;
            query.УстановитьПараметр("Объект", docRef);
            query.УстановитьПараметр("Свойство", propertyRef);
            queryResult = query.Выполнить();
            selection = queryResult.Выбрать();
            if (!(bool)selection.Следующий()) return string.Empty;
            return Str(selection.Значение);
        }
        finally
        {
            ReleaseComObject(selection);
            ReleaseComObject(queryResult);
            ReleaseComObject(query);
        }
    }

    private static void WritePropertyValue(dynamic conn, dynamic docRef, dynamic propertyRef, object value)
    {
        dynamic manager = conn.РегистрыСведений.ЗначенияСвойствОбъектов.СоздатьМенеджерЗаписи();
        try
        {
            manager.Объект = docRef;
            manager.Свойство = propertyRef;
            manager.Значение = value;
            manager.Записать();
        }
        finally
        {
            ReleaseComObject(manager);
        }
    }

    private static bool WriteCheckPropertiesBundle(
        dynamic conn,
        dynamic docRef,
        CorrectionPropertyRefs props,
        string? comment,
        DateTime? checkDate,
        long? fiscalDocument,
        long? fiscalSign,
        ref string correctionCommentState,
        bool appendComment)
    {
        if (!string.IsNullOrWhiteSpace(comment) &&
            CommentAlreadyHasFiscalSign(correctionCommentState, fiscalSign))
            return false;

        var changed = false;

        if (!string.IsNullOrWhiteSpace(comment))
        {
            var merged = appendComment && !string.IsNullOrWhiteSpace(correctionCommentState)
                ? correctionCommentState + "   ///   " + comment
                : comment;
            if (!string.Equals(merged, correctionCommentState, StringComparison.Ordinal))
            {
                WritePropertyValue(conn, docRef, props.Comment, merged);
                correctionCommentState = merged;
                changed = true;
            }
        }

        if (checkDate.HasValue && checkDate.Value > new DateTime(2000, 1, 1))
        {
            WritePropertyValue(conn, docRef, props.CheckDate, checkDate.Value);
            changed = true;
        }

        if (fiscalDocument.HasValue)
        {
            WritePropertyValue(conn, docRef, props.CheckNumber, (double)fiscalDocument.Value);
            changed = true;
        }

        if (fiscalSign.HasValue)
        {
            WritePropertyValue(conn, docRef, props.FiscalSign, (double)fiscalSign.Value);
            changed = true;
        }

        return changed;
    }

    // ── Запрос к УТ 10.3 ─────────────────────────────────────────────────────
    private static string BuildQueryLegacy() => """
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
            И РеализацияТоваровУслуг.ДатаПечатиЧека >= ДАТАВРЕМЯ(2000, 1, 1)
            И НАЧАЛОПЕРИОДА(РеализацияТоваровУслуг.Дата, ДЕНЬ) <> НАЧАЛОПЕРИОДА(РеализацияТоваровУслуг.ДатаПечатиЧека, ДЕНЬ)
            И (РеализацияТоваровУслуг.Комментарий ЕСТЬ NULL
                    ИЛИ РеализацияТоваровУслуг.Комментарий = ""
                    ИЛИ НЕ РеализацияТоваровУслуг.Комментарий ПОДОБНО "%Пробит%")
        УПОРЯДОЧИТЬ ПО
            РеализацияТоваровУслуг.Подразделение.Наименование,
            РеализацияТоваровУслуг.Дата
        """;

    private static string BuildQueryWithProperties() => """
        ВЫБРАТЬ
            РеализацияТоваровУслуг.Ссылка                                              КАК Ссылка,
            РеализацияТоваровУслуг.Номер                                               КАК НомерДок,
            РеализацияТоваровУслуг.Дата                                                КАК Дата,
            РеализацияТоваровУслуг.Сделка.Номер                                        КАК НомерЗаказа,
            РеализацияТоваровУслуг.Сделка.Дата                                         КАК ДатаЗаказа,
            РеализацияТоваровУслуг.Сделка.КонтактноеЛицоКонтрагента.Наименование      КАК Покупатель,
            РеализацияТоваровУслуг.СуммаДокумента                                      КАК СуммаДокумента,
            РеализацияТоваровУслуг.ДоговорКонтрагента.Наименование                     КАК Договор,
            РеализацияТоваровУслуг.Подразделение.Наименование                          КАК Подразделение,
            РеализацияТоваровУслуг.НомерЧекаККМ                                        КАК НомерЧекаККМ,
            РеализацияТоваровУслуг.ЧекНомерФП                                          КАК ЧекНомерФП,
            РеализацияТоваровУслуг.ДатаПечатиЧека                                      КАК ДатаПечатиЧека,
            ЗначКомментарийКорр.Значение                                               КАК КомментарийКорректировки,
            ЗначНомерЧека.Значение                                                     КАК НомерЧекаККМСвойство,
            ЗначЧекФП.Значение                                                         КАК ЧекНомерФПСвойство,
            ЗначДатаПечати.Значение                                                    КАК ДатаПечатиЧекаСвойство,
            ВЫБОР
                КОГДА РеализацияТоваровУслуг.ДатаПечатиЧека >= ДАТАВРЕМЯ(2000, 1, 1)
                    ТОГДА РеализацияТоваровУслуг.ДатаПечатиЧека
                КОГДА ЗначДатаПечати.Значение >= ДАТАВРЕМЯ(2000, 1, 1)
                    ТОГДА ЗначДатаПечати.Значение
                ИНАЧЕ ДАТАВРЕМЯ(1, 1, 1)
            КОНЕЦ                                                                      КАК ДатаПечатиЧекаЭффективная
        ИЗ
            Документ.РеализацияТоваровУслуг КАК РеализацияТоваровУслуг
                ЛЕВОЕ СОЕДИНЕНИЕ РегистрСведений.ЗначенияСвойствОбъектов КАК ЗначКомментарийКорр
                ПО ЗначКомментарийКорр.Объект = РеализацияТоваровУслуг.Ссылка
                    И ЗначКомментарийКорр.Свойство = &СвойствоКомментарийКорректировки
                ЛЕВОЕ СОЕДИНЕНИЕ РегистрСведений.ЗначенияСвойствОбъектов КАК ЗначДатаПечати
                ПО ЗначДатаПечати.Объект = РеализацияТоваровУслуг.Ссылка
                    И ЗначДатаПечати.Свойство = &СвойствоДатаПечатиЧека
                ЛЕВОЕ СОЕДИНЕНИЕ РегистрСведений.ЗначенияСвойствОбъектов КАК ЗначНомерЧека
                ПО ЗначНомерЧека.Объект = РеализацияТоваровУслуг.Ссылка
                    И ЗначНомерЧека.Свойство = &СвойствоНомерЧекаККМ
                ЛЕВОЕ СОЕДИНЕНИЕ РегистрСведений.ЗначенияСвойствОбъектов КАК ЗначЧекФП
                ПО ЗначЧекФП.Объект = РеализацияТоваровУслуг.Ссылка
                    И ЗначЧекФП.Свойство = &СвойствоЧекНомерФП
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
            И РеализацияТоваровУслуг.ДатаПечатиЧека >= ДАТАВРЕМЯ(2000, 1, 1)
            И НАЧАЛОПЕРИОДА(РеализацияТоваровУслуг.Дата, ДЕНЬ) <> НАЧАЛОПЕРИОДА(РеализацияТоваровУслуг.ДатаПечатиЧека, ДЕНЬ)
            И (РеализацияТоваровУслуг.Комментарий ЕСТЬ NULL
                    ИЛИ РеализацияТоваровУслуг.Комментарий = ""
                    ИЛИ НЕ РеализацияТоваровУслуг.Комментарий ПОДОБНО "%Пробит%")
            И (ЗначКомментарийКорр.Значение ЕСТЬ NULL
                    ИЛИ ВЫРАЗИТЬ(ЗначКомментарийКорр.Значение КАК СТРОКА(500)) = ""
                    ИЛИ НЕ ВЫРАЗИТЬ(ЗначКомментарийКорр.Значение КАК СТРОКА(500)) ПОДОБНО "%Пробит%")
        УПОРЯДОЧИТЬ ПО
            РеализацияТоваровУслуг.Подразделение.Наименование,
            РеализацияТоваровУслуг.Дата
        """;
}
