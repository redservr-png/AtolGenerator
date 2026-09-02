using System.IO;
using System.Reflection;
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

public enum RealizationCheckKind
{
    /// <summary>Нет ККМ, ФП и даты — нужен одиночный чек коррекции.</summary>
    NoCheck,
    /// <summary>Чек пробит в другой день — исправительная пара.</summary>
    WrongDay,
    /// <summary>Есть номер ККМ и/или ФП, но нет даты печати — проверить в 1С.</summary>
    Incomplete,
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
    public bool   HasCheck       { get; set; }  // чек пробит не в день реализации
    public RealizationCheckKind CheckKind { get; set; } = RealizationCheckKind.NoCheck;
    public string CheckNumber    { get; set; } = string.Empty;
    public string CheckDate      { get; set; } = string.Empty;
    public string FiscalNumber   { get; set; } = string.Empty;  // ЧекНомерФП
    /// <summary>UUID документа реализации для точной загрузки табличной части.</summary>
    public string DocumentUuid   { get; set; } = string.Empty;
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
        var state = new RealizationLoadState();

        Log($"=== LoadRealizations start: {from:dd.MM.yyyy} – {to:dd.MM.yyyy} ===");
        Log($"Подключение: сервер={s.Server}, база={s.Database}, пользователь={s.User}");

        try
        {
            Log("Создаём коннектор...");
            connector = CreateConnector();

            Log("Подключаемся...");
            conn = connector.Connect(s.ConnectionString);

            var usePropertyQuery = CorrectionPropertyRefs.TryResolve(
                conn, out CorrectionPropertyRefs? correctionProps, out List<string> missingProps);
            if (usePropertyQuery)
                Log("Свойства коррекции найдены — запросы с регистром свойств.");
            else
                Log($"Свойства не найдены ({string.Join(", ", missingProps)}) — упрощённые запросы.");

            try
            {
                ExecuteRealizationQueries(conn, from, to, usePropertyQuery, correctionProps, state);
            }
            catch (Exception ex) when (usePropertyQuery)
            {
                Log($"Запрос со свойствами упал: {FormatComError(ex)} — новое подключение, запросы без регистра.");
                ReleaseComObject(conn);
                conn = connector.Connect(s.ConnectionString);
                state.ClearPartial();
                ExecuteRealizationQueries(conn, from, to, false, null, state);
            }

            Log($"Готово: 1С отдала {state.From1C}, загружено {state.Result.Count}" +
                $", отсеяно классификацией {state.SkippedClassify}" +
                $", «Пробит» в свойстве {state.SkippedProbit}" +
                $", ошибки строк {state.SkippedError}" +
                $", нет чека {state.Result.Count(x => x.CheckKind == RealizationCheckKind.NoCheck)}" +
                $", другой день {state.Result.Count(x => x.CheckKind == RealizationCheckKind.WrongDay)}" +
                $", без даты {state.Result.Count(x => x.CheckKind == RealizationCheckKind.Incomplete)}");
        }
        catch (Exception ex)
        {
            Log($"КРИТИЧЕСКАЯ ОШИБКА: {ex.GetType().Name}: {ex.Message}{Environment.NewLine}{ex.StackTrace}");
            throw;
        }
        finally
        {
            ReleaseComObject(conn);
            ReleaseComObject(connector);
        }

        return state.Result;
    }

    /// <summary>
    /// Два отдельных запроса: «чек в другой день» и «нет даты печати».
    /// В одном WHERE УТ 10.3 обнуляет выборку, если смешать пустую дату и НАЧАЛОПЕРИОДА.
    /// </summary>
    private static void ExecuteRealizationQueries(
        dynamic conn, DateTime from, DateTime to, bool withProperties,
        CorrectionPropertyRefs? correctionProps, RealizationLoadState state)
    {
        ReadRealizationQuery(
            conn, from, to, withProperties, correctionProps, state,
            RealizationDateFilter.WrongDay);
        ReadRealizationQuery(
            conn, from, to, withProperties, correctionProps, state,
            RealizationDateFilter.NoPrintDate);
    }

    private static void ReadRealizationQuery(
        dynamic conn, DateTime from, DateTime to, bool withProperties,
        CorrectionPropertyRefs? correctionProps, RealizationLoadState state,
        RealizationDateFilter dateFilter)
    {
        dynamic? query = null;
        dynamic? queryResult = null;
        dynamic? selection = null;
        var label = dateFilter == RealizationDateFilter.WrongDay
            ? "другой день"
            : "без даты печати";

        try
        {
            query = conn.NewObject("Запрос");
            query.Текст = BuildRealizationQuery(withProperties, dateFilter);
            if (withProperties)
                BindCorrectionPropertyParameters(query, correctionProps!);
            BindSharedQueryParameters(query, from, to);

            Log($"Выполняем запрос «{label}»...");
            queryResult = query.Выполнить();
            selection = queryResult.Выбрать();

            int row = 0;
            bool hasNext;
            while (true)
            {
                try { hasNext = (bool)selection.Следующий(); }
                catch (Exception ex)
                {
                    Log($"Ошибка при вызове Следующий() ({label}) на строке {row}: {ex}");
                    throw;
                }
                if (!hasNext) break;
                row++;
                state.From1C++;

                try
                {
                    AppendRealizationRow((object)conn, (object)selection, withProperties, state);
                }
                catch (Exception rowEx)
                {
                    state.SkippedError++;
                    Log($"Строка {row} ({label}) пропущена: {rowEx.GetType().Name}: {rowEx.Message}");
                }
            }

            Log($"Запрос «{label}»: 1С вернула {row} строк.");
        }
        finally
        {
            ReleaseComObject(selection);
            ReleaseComObject(queryResult);
            ReleaseComObject(query);
        }
    }

    private static void AppendRealizationRow(
        object connObj, object selectionRow, bool withProperties, RealizationLoadState state)
    {
        dynamic conn = connObj;
        dynamic selection = selectionRow;

        // COM-поля сначала в object: иначе dynamic «заражает» ClassifyCheckKind
        // и kind.Value падает (RuntimeBinderException на enum без .Value).
        string docNumber = Str((object?)selection.НомерДок);
        DateTime docDate = ToDateTime((object?)selection.Дата);
        var key = $"{docNumber}|{docDate:yyyyMMdd}";
        if (!state.Seen.Add(key))
            return;

        if (withProperties)
        {
            string propComment = Str((object?)selection.КомментарийКорректировки);
            if (propComment.Contains("Пробит", StringComparison.OrdinalIgnoreCase))
            {
                state.SkippedProbit++;
                return;
            }
        }

        string orderNum = Str((object?)selection.НомерЗаказа);
        DateTime orderDate = ToDateTime((object?)selection.ДатаЗаказа);

        DateTime effectiveCheckDt = ToDateTime((object?)selection.ДатаПечатиЧека);
        if (!IsMeaningfulDate(effectiveCheckDt) && withProperties)
            effectiveCheckDt = ToDateTime((object?)selection.ДатаПечатиЧекаСвойство);

        string checkNum = Str((object?)selection.НомерЧекаККМ);
        if (withProperties && IsEmptyFp(checkNum))
            checkNum = Str((object?)selection.НомерЧекаККМСвойство);
        if (IsEmptyFp(checkNum)) checkNum = string.Empty;

        string fiscalNumber = Str((object?)selection.ЧекНомерФП);
        if (withProperties && IsEmptyFp(fiscalNumber))
            fiscalNumber = Str((object?)selection.ЧекНомерФПСвойство);
        if (IsEmptyFp(fiscalNumber)) fiscalNumber = string.Empty;

        RealizationCheckKind? kind = ClassifyCheckKind(
            docDate, effectiveCheckDt, checkNum, fiscalNumber);
        if (kind is null)
        {
            state.SkippedClassify++;
            return;
        }

        RealizationCheckKind checkKind = kind.Value;
        string dogovor = Str((object?)selection.Договор);
        var isService = dogovor.IndexOf("агент", StringComparison.OrdinalIgnoreCase) >= 0;

        var documentUuid = string.Empty;
        try
        {
            object docRef = selection.ДокСсылка;
            documentUuid = TryReadDocumentUuid(conn, docRef) ?? string.Empty;
        }
        catch { /* поле ДокСсылка может отсутствовать в старом запросе */ }

        state.Result.Add(new OneCRealization
        {
            DocNumber    = docNumber,
            DocDate      = IsMeaningfulDate(docDate)
                            ? docDate.ToString("dd.MM.yyyy")
                            : string.Empty,
            OrderNumber  = orderNum,
            OrderDate    = IsMeaningfulDate(orderDate)
                            ? orderDate.ToString("dd.MM.yyyy HH:mm:ss")
                            : string.Empty,
            CustomerName = Str((object?)selection.Покупатель),
            Amount       = ToDouble((object?)selection.СуммаДокумента),
            IsService    = isService,
            City         = Str((object?)selection.Подразделение),
            HasCheck     = checkKind == RealizationCheckKind.WrongDay,
            CheckKind    = checkKind,
            CheckNumber  = checkNum,
            FiscalNumber = fiscalNumber,
            CheckDate    = IsMeaningfulDate(effectiveCheckDt)
                            ? effectiveCheckDt.ToString("dd.MM.yyyy HH:mm:ss")
                            : string.Empty,
            DocumentUuid = documentUuid,
        });
    }

    /// <summary>
    /// Загружает табличную часть (Товары или Услуги) документа реализации.
    /// Сначала читает объект документа (то, что видно в форме 1С), затем — запрос.
    /// </summary>
    public static List<OneCRealizationItem> LoadRealizationItems(
        OneCConnectionSettings s, string docNumber, DateTime docDate, bool isService,
        object? existingConnection = null, string? documentUuid = null)
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

            dynamic docsManager = conn.Документы.РеализацияТоваровУслуг;
            result = TryLoadRealizationItemsFromObject(
                conn, docsManager, docNumber, docDate, documentUuid);
            if (result.Count > 0)
            {
                Log($"LoadRealizationItems: из объекта документа загружено {result.Count} позиций");
                return result;
            }

            Log("LoadRealizationItems: объект документа недоступен — запрос к табличной части");
            query = conn.NewObject("Запрос");
            query.Текст = """
                ВЫБРАТЬ
                    Строки.НомерСтроки                    КАК НомерСтроки,
                    Строки.Номенклатура.Наименование     КАК Наименование,
                    Строки.Количество                    КАК Количество,
                    Строки.Цена                          КАК Цена,
                    Строки.Сумма                         КАК Сумма
                ИЗ
                    Документ.РеализацияТоваровУслуг.Товары КАК Строки
                ГДЕ
                    Строки.Ссылка.Номер = &НомерДок
                    И Строки.Ссылка.Дата >= &НачалоДня
                    И Строки.Ссылка.Дата < &КонецДня

                ОБЪЕДИНИТЬ ВСЕ

                ВЫБРАТЬ
                    Строки.НомерСтроки,
                    Строки.Номенклатура.Наименование,
                    Строки.Количество,
                    Строки.Цена,
                    Строки.Сумма
                ИЗ
                    Документ.РеализацияТоваровУслуг.Услуги КАК Строки
                ГДЕ
                    Строки.Ссылка.Номер = &НомерДок
                    И Строки.Ссылка.Дата >= &НачалоДня
                    И Строки.Ссылка.Дата < &КонецДня

                УПОРЯДОЧИТЬ ПО
                    НомерСтроки
                """;
            query.УстановитьПараметр("НомерДок", docNumber);
            query.УстановитьПараметр("НачалоДня", docDate.Date);
            query.УстановитьПараметр("КонецДня", docDate.Date.AddDays(1));

            queryResult = query.Выполнить();
            selection   = queryResult.Выбрать();

            while ((bool)selection.Следующий())
                AppendLineFromQuerySelection(result, selection);

            if (result.Count == 0)
            {
                Log("LoadRealizationItems: по дате строк нет — повтор за календарный год...");
                result = LoadRealizationItemsByNumberYear(conn, docNumber, docDate.Year);
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

    private static List<OneCRealizationItem> TryLoadRealizationItemsFromObject(
        dynamic conn, dynamic docsManager, string docNumber, DateTime docDate, string? documentUuid)
    {
        var result = new List<OneCRealizationItem>();
        object? docRef = null;

        if (!string.IsNullOrWhiteSpace(documentUuid))
        {
            try
            {
                object uid = conn.NewObject("УникальныйИдентификатор", documentUuid);
                docRef = ComInvoke((object)docsManager, ["GetRef", "ПолучитьСсылку"], uid);
            }
            catch (Exception ex)
            {
                Log($"LoadRealizationItems: UUID {documentUuid}: {ex.Message}");
            }
        }

        if (docRef is null || IsEmptyOneCRef(docRef))
        {
            if (!TryFindRealizationForWrite(
                    (object)conn, docNumber, docDate.Year,
                    out docRef, out var foundDate, out _, out var findError))
            {
                Log($"LoadRealizationItems: документ не найден — {findError}");
                return result;
            }

            if (IsMeaningfulDate(foundDate))
                docDate = foundDate;
        }

        string failReasons;
        var docObj = TryGetRealizationObject(
            (object)conn, (object)docsManager, (object)docRef!, docNumber, docDate, out failReasons);
        if (docObj is null)
        {
            Log($"LoadRealizationItems: ПолучитьОбъект — {failReasons}");
            return result;
        }

        try
        {
            AppendTabularSectionItems(result, docObj, "Товары");
            AppendTabularSectionItems(result, docObj, "Услуги");
        }
        finally
        {
            try { Marshal.ReleaseComObject(docObj); } catch { /* ignore */ }
        }

        return result;
    }

    private static void AppendTabularSectionItems(
        List<OneCRealizationItem> target, dynamic docObj, string sectionName)
    {
        dynamic rows;
        try { rows = docObj.GetType().InvokeMember(
            sectionName, BindingFlags.GetProperty, null, docObj, null)!; }
        catch { return; }

        var count = ReadTabularRowCount(rows);
        if (count <= 0) return;

        for (var i = 0; i < count; i++)
        {
            object? rowObj = null;
            try { rowObj = ComInvoke((object)rows, ["Get", "Получить"], i); }
            catch { continue; }

            if (rowObj is null) continue;

            dynamic row = rowObj;
            var name = ReadNomenclatureName(row);
            var price = ToDouble(row.Цена);
            var rawQty = ToQuantity(row.Количество);
            var sum = ToDouble(row.Сумма);
            if (string.IsNullOrWhiteSpace(name) || sum <= 0) continue;

            var qty = ResolveLineQuantity(price, rawQty, sum);
            target.Add(new OneCRealizationItem
            {
                Name     = name,
                Quantity = qty,
                Sum      = sum,
            });
            Log($"  · [{sectionName}] {name}: qty={qty}, price={price:F2}, sum={sum:F2}");
        }
    }

    private static int ReadTabularRowCount(dynamic rows)
    {
        try
        {
            var count = ComInvoke((object)rows, ["Count", "Количество"]);
            if (count is not null)
                return Convert.ToInt32(count);
        }
        catch { /* dynamic fallback below */ }

        try { return (int)rows.Количество(); }
        catch { return 0; }
    }

    private static string ReadNomenclatureName(dynamic row)
    {
        try
        {
            dynamic nom = row.Номенклатура;
            if (nom is null) return string.Empty;
            try { return Str(nom.Наименование); }
            catch { return Str(nom); }
        }
        catch { return string.Empty; }
    }

    private static void AppendLineFromQuerySelection(List<OneCRealizationItem> target, dynamic selection)
    {
        var price = ToDouble(selection.Цена);
        var rawQty = ToQuantity(selection.Количество);
        var sum = ToDouble(selection.Сумма);
        var name = Str(selection.Наименование);
        if (string.IsNullOrWhiteSpace(name) || sum <= 0) return;

        var qty = ResolveLineQuantity(price, rawQty, sum);
        target.Add(new OneCRealizationItem
        {
            Name     = name,
            Quantity = qty,
            Sum      = sum,
        });
        Log($"  · {name}: qty={qty}, price={price:F2}, sum={sum:F2}");
    }

    /// <summary>
    /// Согласует количество с ценой и суммой строки реализации.
    /// </summary>
    private static double ResolveLineQuantity(double price, double quantity, double sum)
    {
        sum = Math.Round(sum, 2);
        if (sum <= 0)
            return quantity > 0 ? quantity : 1;

        price = Math.Round(price, 2);
        if (price > 0)
        {
            var qtyFromPrice = sum / price;
            for (var decimals = 0; decimals <= 3; decimals++)
            {
                var factor = Math.Pow(10, decimals);
                var rounded = Math.Round(qtyFromPrice * factor, MidpointRounding.AwayFromZero) / factor;
                if (rounded <= 0) continue;
                if (Math.Abs(Math.Round(price * rounded, 2) - sum) <= 0.01)
                    return rounded;
            }
        }

        if (quantity > 0 && price > 0 && Math.Abs(Math.Round(price * quantity, 2) - sum) <= 0.01)
            return quantity;

        return quantity > 0 ? quantity : 1;
    }

    private static List<OneCRealizationItem> LoadRealizationItemsByNumberYear(
        dynamic conn, string docNumber, int year)
    {
        var result = new List<OneCRealizationItem>();
        var query = conn.NewObject("Запрос");
        query.Текст = """
            ВЫБРАТЬ
                Строки.НомерСтроки                    КАК НомерСтроки,
                Строки.Номенклатура.Наименование     КАК Наименование,
                Строки.Количество                    КАК Количество,
                Строки.Цена                          КАК Цена,
                Строки.Сумма                         КАК Сумма
            ИЗ
                Документ.РеализацияТоваровУслуг.Товары КАК Строки
            ГДЕ
                Строки.Ссылка.Номер = &НомерДок
                И Строки.Ссылка.Дата >= &НачалоГода
                И Строки.Ссылка.Дата < &КонецГода

            ОБЪЕДИНИТЬ ВСЕ

            ВЫБРАТЬ
                Строки.НомерСтроки,
                Строки.Номенклатура.Наименование,
                Строки.Количество,
                Строки.Цена,
                Строки.Сумма
            ИЗ
                Документ.РеализацияТоваровУслуг.Услуги КАК Строки
            ГДЕ
                Строки.Ссылка.Номер = &НомерДок
                И Строки.Ссылка.Дата >= &НачалоГода
                И Строки.Ссылка.Дата < &КонецГода

            УПОРЯДОЧИТЬ ПО
                НомерСтроки
            """;
        var yearStart = new DateTime(year, 1, 1);
        query.УстановитьПараметр("НомерДок", docNumber);
        query.УстановитьПараметр("НачалоГода", yearStart);
        query.УстановитьПараметр("КонецГода", yearStart.AddYears(1));

        var selection = query.Выполнить().Выбрать();
        while ((bool)selection.Следующий())
            AppendLineFromQuerySelection(result, selection);

        ReleaseComObject(selection);
        ReleaseComObject(query);
        return result;
    }

    /// <summary>
    /// Перечитывает табличную часть реализации из 1С в позиции заказа перед формированием чека.
    /// </summary>
    public static void RefreshRealizationLineItems(OneCConnectionSettings settings, Models.OrderEntry order)
    {
        if (order.DocumentType != Models.SourceDocumentType.Realization)
            return;

        var docNumber = !string.IsNullOrWhiteSpace(order.CorrectionNumber)
            ? order.CorrectionNumber
            : order.OrderNum;
        var dateRaw = !string.IsNullOrWhiteSpace(order.CorrectionDate)
            ? order.CorrectionDate
            : order.OrderDate;
        if (!TryParseDocumentDate(dateRaw, out var docDate))
            throw new InvalidOperationException(
                $"{docNumber}: не определена дата реализации для загрузки табличной части.");

        var realization = new OneCRealization
        {
            DocNumber = docNumber,
            DocDate = docDate.ToString("dd.MM.yyyy"),
            IsService = order.IsService,
            DocumentUuid = order.DocumentUuid,
        };
        EnrichRealizationForReceipt(settings, realization);
        order.Items = realization.Items.Select(x => new Models.OrderItem
        {
            Name = x.Name,
            Quantity = x.Quantity,
            Sum = x.Sum,
        }).ToList();

        if (order.Kind == Models.OrderKind.RefundCorrectionPair ||
            order.OriginalItems.Count > 0 ||
            !string.IsNullOrWhiteSpace(order.PlannedReverseOperation))
        {
            order.OriginalItems = order.Items.Select(x => new Models.OrderItem
            {
                Name = x.Name,
                Quantity = x.Quantity,
                Sum = x.Sum,
                VatType = x.VatType,
            }).ToList();
        }
    }

    private static bool TryParseDocumentDate(string raw, out DateTime date)
    {
        if (DateTime.TryParseExact(raw, "dd.MM.yyyy HH:mm:ss",
                System.Globalization.CultureInfo.InvariantCulture,
                System.Globalization.DateTimeStyles.None, out date))
            return true;
        if (DateTime.TryParseExact(raw, "dd.MM.yyyy",
                System.Globalization.CultureInfo.InvariantCulture,
                System.Globalization.DateTimeStyles.None, out date))
            return true;

        date = default;
        return false;
    }

    public static void EnrichRealizationForReceipt(
        OneCConnectionSettings s, OneCRealization realization)
        => EnrichRealizationForReceipt(s, realization, null);

    private static void EnrichRealizationForReceipt(
        OneCConnectionSettings s, OneCRealization realization, object? existingConnection)
    {
        if (!string.IsNullOrWhiteSpace(realization.DocNumber))
        {
            if (!DateTime.TryParseExact(realization.DocDate, "dd.MM.yyyy",
                    System.Globalization.CultureInfo.InvariantCulture,
                    System.Globalization.DateTimeStyles.None, out var docDate))
                throw new InvalidOperationException(
                    $"{realization.DocNumber}: не определена дата реализации для загрузки табличной части.");

            // Всегда перечитываем из 1С: в чек идут строки реализации, не заказа покупателя.
            realization.Items = LoadRealizationItems(
                s,
                realization.DocNumber,
                docDate,
                realization.IsService,
                existingConnection,
                string.IsNullOrWhiteSpace(realization.DocumentUuid) ? null : realization.DocumentUuid);
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
        public List<string> ProcessedNumbers { get; set; } = new(); // обновлённые и уже заполненные
        public string  CsvBackupPath       { get; set; } = string.Empty;   // путь к CSV для ручного импорта
    }

    public class PunchedRecord
    {
        public string RealizationNum { get; set; } = string.Empty;
        public DateTime? DocumentDate { get; set; }
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

                    lastStep = "query";
                    object? foundRef;
                    DateTime docDate;
                    string fpRaw;
                    string findError;
                    if (!TryFindRealizationForWrite(
                            (object)conn, rec.RealizationNum, DateTime.Now.Year,
                            out foundRef, out docDate, out fpRaw, out findError))
                    {
                        res.Failed++;
                        res.Errors.Add(findError);
                        Log("  " + findError);
                        continue;
                    }

                    var docRef = foundRef;
                    if (docRef is null)
                    {
                        res.Failed++;
                        var msg = $"{rec.RealizationNum}: ссылка пустая";
                        res.Errors.Add(msg);
                        Log("  " + msg);
                        continue;
                    }

                    var isFilled = !IsEmptyFp(fpRaw);
                    if (skipFilled && isFilled)
                    {
                        res.Skipped++;
                        var detail = $"{rec.RealizationNum}: дата={docDate:dd.MM.yyyy} ЧекНомерФП = «{fpRaw}»";
                        Log($"  {detail} — пропуск");
                        if (res.SkippedSamples.Count < 15) res.SkippedSamples.Add(detail);
                        continue;
                    }

                    Log($"  {rec.RealizationNum}: дата={docDate:dd.MM.yyyy} текущ.ФП=«{fpRaw}» → пишем ФПД={rec.FiscalSign}");

                    lastStep = "ПолучитьОбъект";
                    string failReasons;
                    var obj = TryGetRealizationObject(
                        (object)conn, (object)docsManager!, (object)docRef, rec.RealizationNum, docDate, out failReasons);

                    lastStep = "write properties";
                    string correctionComment = ReadPropertyString(conn, docRef, correctionProps.Comment);
                    var comment = $"{checkDate:dd.MM.yyyy} Пробит чек коррекции \"Приход\" ФП: {rec.FiscalSign!.Value}";
                    var propsChanged = WriteCheckPropertiesBundle(
                        conn,
                        docRef,
                        correctionProps,
                        comment,
                        checkDate,
                        rec.FiscalDoc,
                        rec.FiscalSign,
                        ref correctionComment,
                        appendComment: true);

                    var fieldsChanged = false;
                    if (obj is not null)
                    {
                        lastStep = "set obj.ЧекНомерФП";
                        obj.ЧекНомерФП = (double)rec.FiscalSign.Value;

                        lastStep = "set obj.НомерЧекаККМ";
                        obj.НомерЧекаККМ = (double)rec.FiscalDoc.Value;

                        lastStep = "set obj.ДатаПечатиЧека";
                        obj.ДатаПечатиЧека = checkDate;

                        lastStep = "obj.Записать()";
                        obj.Записать();
                        fieldsChanged = true;
                        try { Marshal.ReleaseComObject(obj); } catch { }
                    }
                    else
                    {
                        Log($"  {rec.RealizationNum}: реквизиты документа не записаны ({failReasons}) — свойства " +
                            (propsChanged ? "записаны" : "без изменений"));
                    }

                    if (!fieldsChanged && !propsChanged)
                    {
                        res.Failed++;
                        var msg = $"{rec.RealizationNum}: ПолучитьОбъект упал. {failReasons}";
                        res.Errors.Add(msg);
                        Log("  ОШИБКА " + msg);
                        continue;
                    }

                    res.Updated++;
                    Log($"  {rec.RealizationNum}: дата={docDate:dd.MM.yyyy} ФПД={rec.FiscalSign} №ФД={rec.FiscalDoc}" +
                        (fieldsChanged ? " → реквизиты + свойства" : " → только свойства"));
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
                    object? foundRef;
                    DateTime docDate;
                    string fpRaw;
                    string findError;
                    if (!TryFindRealizationForWrite(
                            (object)conn, realizationNum, DateTime.Now.Year,
                            out foundRef, out docDate, out fpRaw, out findError))
                    {
                        res.Failed++;
                        res.Errors.Add(findError);
                        Log("  " + findError);
                        continue;
                    }

                    var docRef = foundRef;
                    Log($"  {realizationNum}: документ {docDate:dd.MM.yyyy} (поиск в {DateTime.Now.Year} г.)");

                    var isFilled = !IsEmptyFp(fpRaw);
                    if (!commentOnly && skipFilled && isFilled)
                        Log($"  {realizationNum}: ЧекНомерФП уже «{fpRaw}» — реквизиты не трогаем, свойства проверим");

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
                            res.ProcessedNumbers.Add(realizationNum);
                            var detail = $"{realizationNum}: дата={docDate:dd.MM.yyyy} — свойства без изменений";
                            Log("  " + detail);
                            if (res.SkippedSamples.Count < 15) res.SkippedSamples.Add(detail);
                            continue;
                        }

                        res.Updated++;
                        res.ProcessedNumbers.Add(realizationNum);
                        Log($"  {realizationNum}: дата={docDate:dd.MM.yyyy} mode=comment_only → {ordered.Count} чек(ов) в свойства");
                        continue;
                    }

                    string failReasons;
                    dynamic? obj = TryGetRealizationObject(
                        (object)conn, (object)docsManager!, (object)docRef, realizationNum, docDate, out failReasons);

                    var fieldsChanged = false;
                    if (obj is not null)
                    {
                        if (updateRow is not null && !(skipFilled && isFilled))
                        {
                            lastStep = "set fields";
                            obj.ЧекНомерФП = (double)updateRow.FiscalSign!.Value;
                            obj.НомерЧекаККМ = (double)updateRow.FiscalDocument!.Value;
                            obj.ДатаПечатиЧека = updateRow.RegisteredAt ?? checkDate;
                            lastStep = "obj.Записать()";
                            obj.Записать();
                            fieldsChanged = true;
                            changed = true;
                        }

                        try { Marshal.ReleaseComObject(obj); } catch { }
                    }
                    else
                    {
                        Log($"  {realizationNum}: реквизиты документа недоступны ({failReasons})");
                    }

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
                        res.ProcessedNumbers.Add(realizationNum);
                        var detail = $"{realizationNum}: дата={docDate:dd.MM.yyyy} — без изменений";
                        Log("  " + detail);
                        if (res.SkippedSamples.Count < 15) res.SkippedSamples.Add(detail);
                        continue;
                    }

                    res.Updated++;
                    res.ProcessedNumbers.Add(realizationNum);
                    Log($"  {realizationNum}: дата={docDate:dd.MM.yyyy} mode=update_fields → " +
                        (fieldsChanged ? "реквизиты + свойства" : "только свойства"));
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
            Directory.CreateDirectory(FileHelper.GetProcessedXmlDirectory(DateTime.Now));
            var csvPath = Path.Combine(
                FileHelper.GetProcessedXmlDirectory(DateTime.Now),
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

        var docNumber = !string.IsNullOrWhiteSpace(order.CorrectionNumber)
            ? order.CorrectionNumber
            : order.OrderNum;
        var dateRaw = !string.IsNullOrWhiteSpace(order.CorrectionDate)
            ? order.CorrectionDate
            : order.OrderDate;
        if (!TryParseDocumentDate(dateRaw, out var documentDate))
            return;

        var realization = new OneCRealization
        {
            DocNumber = docNumber,
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
                Quantity = ToQuantity(selection.Количество),
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

    /// <summary>
    /// C# dynamic часто падает NRE на кириллических методах 1С COM.
    /// IDispatch через InvokeMember + английские алиасы (GetObject, FindByNumber).
    /// </summary>
    private static dynamic? TryGetRealizationObject(
        object connObj, object docsManagerObj, object docRef, string number, DateTime periodDate, out string errors)
    {
        dynamic conn = connObj;
        dynamic docsManager = docsManagerObj;
        var reasons = new List<string>();
        errors = string.Empty;

        dynamic? FromRef(object reference, string tag)
        {
            var obj = ComInvoke(reference, ["GetObject", "ПолучитьОбъект"]);
            if (obj is not null) return obj;
            reasons.Add($"{tag}: нет объекта");
            return null;
        }

        try
        {
            var uuid = TryReadDocumentUuid(conn, docRef);
            if (!string.IsNullOrWhiteSpace(uuid))
            {
                object? uid = null;
                try { uid = conn.NewObject("УникальныйИдентификатор", uuid); }
                catch (Exception ex) { reasons.Add($"UUID: {ex.Message}"); }

                if (uid is not null)
                {
                    var restored = ComInvoke(
                        (object)docsManager, ["GetRef", "ПолучитьСсылку"], uid);
                    if (restored is not null)
                    {
                        var obj = FromRef(restored, "uuid");
                        if (obj is not null) return obj;
                    }
                }

                try
                {
                    var xmlType = conn.XMLТипЗнч(docRef);
                    var xml = Convert.ToString(conn.XMLСтрока(docRef));
                    if (xmlType is not null && !string.IsNullOrWhiteSpace(xml))
                    {
                        object restoredXml = conn.XMLЗначение(xmlType, xml);
                        var obj = FromRef(restoredXml, "xml");
                        if (obj is not null) return obj;
                    }
                }
                catch (Exception ex) { reasons.Add($"xml: {ex.Message}"); }
            }
        }
        catch (Exception ex) { reasons.Add($"uuid-path: {ex.Message}"); }

        try
        {
            var found = ComInvoke(
                (object)docsManager,
                ["FindByNumber", "НайтиПоНомеру"],
                number, periodDate);
            if (found is not null && !IsEmptyOneCRef(found))
            {
                var obj = FromRef(found, "номер");
                if (obj is not null) return obj;
            }
            else
            {
                reasons.Add("номер: пустая ссылка");
            }
        }
        catch (Exception ex) { reasons.Add($"номер: {ex.Message}"); }

        try
        {
            var obj = FromRef(docRef, "ref");
            if (obj is not null) return obj;
        }
        catch (Exception ex) { reasons.Add($"ref: {ex.Message}"); }

        errors = string.Join("; ", reasons);
        return null;
    }

    private static string? TryReadDocumentUuid(dynamic conn, object docRef)
    {
        try
        {
            var xml = Convert.ToString(conn.XMLСтрока(docRef));
            if (!string.IsNullOrWhiteSpace(xml) && xml.Length >= 32)
                return xml.Trim();
        }
        catch { /* English alias next */ }

        try
        {
            var xml = Convert.ToString(conn.XMLString(docRef));
            if (!string.IsNullOrWhiteSpace(xml) && xml.Length >= 32)
                return xml.Trim();
        }
        catch { /* UUID method next */ }

        try
        {
            var uid = ComInvoke(docRef, ["UUID", "УникальныйИдентификатор"]);
            var text = uid?.ToString();
            if (!string.IsNullOrWhiteSpace(text))
                return text.Trim();
        }
        catch { /* ignore */ }

        return null;
    }

    private static object? ComInvoke(object target, string[] names, params object?[] args)
    {
        var flags = BindingFlags.InvokeMethod | BindingFlags.Public;
        foreach (var name in names)
        {
            try
            {
                var result = target.GetType().InvokeMember(
                    name,
                    flags,
                    binder: null,
                    target: target,
                    args: args.Length == 0 ? null : args);
                if (result is not null)
                    return result;
            }
            catch
            {
                /* следующий алиас */
            }
        }

        foreach (var name in names)
        {
            try
            {
                dynamic d = target;
                if (args.Length == 0)
                {
                    if (name is "GetObject") return d.GetObject();
                    if (name is "ПолучитьОбъект") return d.ПолучитьОбъект();
                }
            }
            catch
            {
                /* следующий алиас */
            }
        }

        return null;
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

    private static RealizationCheckKind? ClassifyCheckKind(
        DateTime docDate, DateTime checkDate, string checkNumber, string fiscalNumber)
    {
        var hasDate = IsMeaningfulDate(checkDate);
        var hasKkm = !IsEmptyFp(checkNumber);
        var hasFp = !IsEmptyFp(fiscalNumber);

        if (!hasDate && !hasKkm && !hasFp)
            return RealizationCheckKind.NoCheck;

        if (!hasDate)
            return RealizationCheckKind.Incomplete;

        if (IsMeaningfulDate(docDate) && docDate.Date == checkDate.Date)
            return null;

        return RealizationCheckKind.WrongDay;
    }

    private static double ToDouble(dynamic? v)
    {
        if (v is null) return 0.0;
        try
        {
            if (v is double d) return d;
            if (v is float f) return f;
            if (v is decimal m) return (double)m;
            if (v is int i) return i;
            if (v is long l) return l;
            var text = Str(v).Replace(" ", string.Empty).Replace(',', '.');
            if (double.TryParse(text, System.Globalization.NumberStyles.Any,
                    System.Globalization.CultureInfo.InvariantCulture, out double parsed))
                return parsed;
            return Convert.ToDouble(v);
        }
        catch { return 0.0; }
    }

    private static double ToQuantity(dynamic? v)
    {
        var qty = ToDouble(v);
        return qty > 0 ? qty : 0;
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
                var reference = plan.НайтиПоНаименованию(name, true);
                if (IsEmptyOneCRef(reference))
                    reference = plan.НайтиПоНаименованию(name);
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

    private static void BindSharedQueryParameters(dynamic query, DateTime from, DateTime to)
    {
        query.УстановитьПараметр("НачалоПериода", from.Date);
        query.УстановитьПараметр("КонецПериода", to.Date.AddDays(1).AddSeconds(-1));
    }

    /// <summary>
    /// Номер реализации в УТ 10.3 уникален в пределах года — ищем только в указанном году.
    /// </summary>
    private static bool TryFindRealizationForWrite(
        object connObj, string number, int year,
        out object? docRef, out DateTime docDate, out string fpRaw, out string error)
    {
        docRef = null;
        docDate = DateTime.MinValue;
        fpRaw = string.Empty;
        error = string.Empty;

        var yearStart = new DateTime(year, 1, 1);
        var yearEnd = yearStart.AddYears(1);

        dynamic conn = connObj;
        dynamic? query = null;
        dynamic? queryResult = null;
        dynamic? selection = null;
        try
        {
            query = conn.NewObject("Запрос");
            query.Текст = """
                ВЫБРАТЬ ПЕРВЫЕ 1
                    Док.Ссылка     КАК ДокСсылка,
                    Док.Дата       КАК ДатаДок,
                    Док.ЧекНомерФП КАК ЧекНомерФП
                ИЗ
                    Документ.РеализацияТоваровУслуг КАК Док
                ГДЕ
                    Док.Номер = &НомерДок
                    И Док.ПометкаУдаления = ЛОЖЬ
                    И Док.Дата >= &НачалоГода
                    И Док.Дата < &КонецГода
                """;
            query.УстановитьПараметр("НомерДок", number);
            query.УстановитьПараметр("НачалоГода", yearStart);
            query.УстановитьПараметр("КонецГода", yearEnd);
            queryResult = query.Выполнить();
            selection = queryResult.Выбрать();
            if (!(bool)selection.Следующий())
            {
                error = $"{number}: нет реализации в {year} году";
                return false;
            }

            object liveRef = selection.ДокСсылка;
            docDate = ToDateTime((object?)selection.ДатаДок);
            try
            {
                var rawFp = selection.ЧекНомерФП;
                if (rawFp is not null)
                    fpRaw = (rawFp.ToString() ?? string.Empty).Trim();
            }
            catch { /* ignore */ }

            if (liveRef is null)
            {
                error = $"{number}: пустая ссылка в {year} году";
                return false;
            }

            var uuid = TryReadDocumentUuid(conn, liveRef);
            if (!string.IsNullOrWhiteSpace(uuid))
            {
                try
                {
                    object uid = conn.NewObject("УникальныйИдентификатор", uuid);
                    object mgr = conn.Документы.РеализацияТоваровУслуг;
                    docRef = ComInvoke(mgr, ["GetRef", "ПолучитьСсылку"], uid) ?? liveRef;
                }
                catch
                {
                    docRef = liveRef;
                }
            }
            else
            {
                docRef = liveRef;
            }

            return true;
        }
        finally
        {
            ReleaseComObject(selection);
            ReleaseComObject(queryResult);
            ReleaseComObject(query);
        }
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
            Log($"  свойство записано ({(value is DateTime dt ? dt.ToString("dd.MM.yyyy") : value)})");
        }
        catch (Exception ex)
        {
            Log($"  ошибка записи свойства: {FormatComError(ex)}");
            throw;
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
    private enum RealizationDateFilter
    {
        WrongDay,
        NoPrintDate,
    }

    private sealed class RealizationLoadState
    {
        public List<OneCRealization> Result { get; } = new();
        public HashSet<string> Seen { get; } = new(StringComparer.OrdinalIgnoreCase);
        public int From1C;
        public int SkippedClassify;
        public int SkippedProbit;
        public int SkippedError;

        public void ClearPartial()
        {
            Result.Clear();
            Seen.Clear();
            From1C = 0;
            SkippedClassify = 0;
            SkippedProbit = 0;
            SkippedError = 0;
        }
    }

    private static string BuildRealizationQuery(bool withProperties, RealizationDateFilter dateFilter)
    {
        var propertyFields = withProperties ? """
            ,
            ЗначКомментарийКорр.Значение                                               КАК КомментарийКорректировки,
            ЗначНомерЧека.Значение                                                     КАК НомерЧекаККМСвойство,
            ЗначЧекФП.Значение                                                         КАК ЧекНомерФПСвойство,
            ЗначДатаПечати.Значение                                                    КАК ДатаПечатиЧекаСвойство
            """ : "";

        var propertyJoins = withProperties ? """
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
            """ : "";

        var dateWhere = dateFilter == RealizationDateFilter.WrongDay
            ? """
            И РеализацияТоваровУслуг.ДатаПечатиЧека >= ДАТАВРЕМЯ(2000, 1, 1)
            И НАЧАЛОПЕРИОДА(РеализацияТоваровУслуг.Дата, ДЕНЬ) <> НАЧАЛОПЕРИОДА(РеализацияТоваровУслуг.ДатаПечатиЧека, ДЕНЬ)
            """
            : """
            И РеализацияТоваровУслуг.ДатаПечатиЧека < ДАТАВРЕМЯ(2000, 1, 1)
            """;

        return $"""
            ВЫБРАТЬ
                РеализацияТоваровУслуг.Ссылка                                          КАК ДокСсылка,
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
                РеализацияТоваровУслуг.ДатаПечатиЧека                                   КАК ДатаПечатиЧека{propertyFields}
            ИЗ
                Документ.РеализацияТоваровУслуг КАК РеализацияТоваровУслуг
            {propertyJoins}
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
            {dateWhere}
                И НЕ РеализацияТоваровУслуг.Комментарий ПОДОБНО "%Пробит%"
            УПОРЯДОЧИТЬ ПО
                РеализацияТоваровУслуг.Подразделение.Наименование,
                РеализацияТоваровУслуг.Дата
            """;
    }
}
