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

                    if (!string.IsNullOrEmpty(city))
                        order.City = city;
                    if (string.IsNullOrEmpty(order.CustomerName) && !string.IsNullOrEmpty(customer))
                        order.CustomerName = customer;
                    // IsService не меняем — метод вызывается только для уже помеченных услуг

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

    public class ApplyResult
    {
        public int Total       { get; set; }
        public int Updated     { get; set; }
        public int Skipped     { get; set; }
        public int Failed      { get; set; }
        public List<string> Errors         { get; set; } = new();
        public List<string> SkippedSamples { get; set; } = new();  // первые N пропусков с подробностями
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
                            Док.ЧекНомерФП   КАК ЧекНомерФП,
                            Док.Комментарий  КАК Комментарий
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
                    var existingComment = Str(sel.Комментарий);

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

                    // ПолучитьОбъект — пробуем 3 способа подряд:
                    //   1) docsManager.ПолучитьОбъект(docRef)
                    //   2) свежая ссылка через НайтиПоНомеру → .ПолучитьОбъект()
                    //   3) docRef.ПолучитьОбъект() напрямую
                    // Какой-то из них должен сработать в зависимости от поведения COM/УТ.
                    dynamic? obj = null;
                    string failReasons = string.Empty;

                    try
                    {
                        lastStep = "1) docsManager.ПолучитьОбъект(docRef)";
                        obj = docsManager!.ПолучитьОбъект(docRef);
                    }
                    catch (Exception ex1) { failReasons += $"[1: {ex1.Message}] "; }

                    if (obj is null)
                    {
                        try
                        {
                            lastStep = "2) mgr.НайтиПоНомеру → ПолучитьОбъект()";
                            var freshRef = docsManager!.НайтиПоНомеру(rec.RealizationNum, checkDate);
                            if (freshRef is not null && !(bool)freshRef.Пустая())
                                obj = freshRef.ПолучитьОбъект();
                        }
                        catch (Exception ex2) { failReasons += $"[2: {ex2.Message}] "; }
                    }

                    if (obj is null)
                    {
                        try
                        {
                            lastStep = "3) docRef.ПолучитьОбъект()";
                            obj = docRef.ПолучитьОбъект();
                        }
                        catch (Exception ex3) { failReasons += $"[3: {ex3.Message}] "; }
                    }

                    if (obj is null)
                    {
                        res.Failed++;
                        var msg = $"{rec.RealizationNum}: все 3 способа ПолучитьОбъект упали. {failReasons}";
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

                    // 4. Комментарий: пустой → marker; непустой → дописываем через "   \\\   ".
                    //                 Если уже содержит «Пробит чек коррекции» — не дописываем.
                    const string marker = "Пробит чек коррекции \"Приход\"";
                    if (!existingComment.Contains("Пробит чек коррекции", StringComparison.OrdinalIgnoreCase))
                    {
                        lastStep = "set obj.Комментарий";
                        obj.Комментарий = string.IsNullOrWhiteSpace(existingComment)
                            ? marker
                            : existingComment + "   \\\\\\   " + marker;
                    }

                    // 5. Запись — без ОбменДанными.Загрузка (он не во всех конфигурациях работает)
                    lastStep = "obj.Записать()";
                    obj.Записать();
                    res.Updated++;
                    Log($"  {rec.RealizationNum}: дата={docDate:dd.MM.yyyy} ФПД={rec.FiscalSign} №ФД={rec.FiscalDoc} → записано");
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
        return res;
    }

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
