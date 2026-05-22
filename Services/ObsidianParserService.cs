using System.IO;
using System.Text.RegularExpressions;
using AtolGenerator.Models;

namespace AtolGenerator.Services;

/// <summary>
/// Парсит файл с кейсами коррекций из Obsidian («Исправить чеки.md»).
/// Берёт только невыполненные строки (`- [ ]`), пропускает `- [x]` и заголовки.
///
/// Формат поддерживаемой строки:
///   - [ ] Реализация товаров и услуг т0000080533 от 10.09.2025 16:11:59 (Белозерск) Пробит чек, реализация удалена
///   - [ ] Оплата от покупателя платежной картой (кредит) т0000059989 от 11.09.2025 13:08:25 (Сокол) деньги не списались
///   - [ ] Приходный кассовый ордер т0000016459 от 15.09.2025 16:22:39 (Тутаев) фактического прихода не было
///   - [ ] ФП: 2899446670 (Котлас) Лишний чек
/// </summary>
public static class ObsidianParserService
{
    // ── Регулярки парсинга ──────────────────────────────────────────────────────

    // Строка с галочкой: - [ ] ... или - [x] ...
    private static readonly Regex RxTaskLine = new(
        @"^\s*-\s*\[(?<done>[ xX])\]\s*(?<body>.+)$",
        RegexOptions.Compiled);

    // Номер документа: т0000123456 (т или Т, потом 10 цифр — стандарт УТ 10.3)
    // Делаем диапазон 7-11 чтобы поддержать редкие усечённые форматы, но в большинстве случаев — 10.
    private static readonly Regex RxDocNumber = new(
        @"[тТ]\d{7,11}",
        RegexOptions.Compiled);

    // Сумма в строке: «сумма 35242,00», «сумма 1 500,00», «11030 руб»
    private static readonly Regex RxAmountInline = new(
        @"сумма\s+(?<a>\d[\d\s]*[,.]\d{2})|(?<r>\d[\d\s]*)\s*руб",
        RegexOptions.IgnoreCase | RegexOptions.Compiled);

    // Дата+время: 10.09.2025 16:11:59 или 10.09.2025 9:36:01 (часы могут быть без ведущего нуля)
    private static readonly Regex RxDateTime = new(
        @"(?<d>\d{2}\.\d{2}\.\d{4})(?:\s+(?<t>\d{1,2}:\d{2}(?::\d{2})?))?",
        RegexOptions.Compiled);

    // Город в скобках: (Белозерск), (Сокол), (Архангельск), (Микс) и т.д.
    private static readonly Regex RxCity = new(
        @"\(([^)]+)\)",
        RegexOptions.Compiled);

    // ФП-только строки: «ФП: 2899446670 (Котлас) Лишний чек»
    private static readonly Regex RxFpOnly = new(
        @"^\s*ФП:\s*(?<fp>\d+)",
        RegexOptions.Compiled | RegexOptions.IgnoreCase);

    // ── Шаблоны типов документов ────────────────────────────────────────────────

    private static readonly (Regex pattern, SourceDocumentType type)[] DocTypePatterns =
    {
        (new(@"Реализация\s+товаров\s+и\s+услуг",          RegexOptions.IgnoreCase | RegexOptions.Compiled), SourceDocumentType.Realization),
        (new(@"Оплата\s+от\s+покупателя\s+платежной\s+картой", RegexOptions.IgnoreCase | RegexOptions.Compiled), SourceDocumentType.CardPayment),
        (new(@"Приходный\s+кассовый\s+ордер",               RegexOptions.IgnoreCase | RegexOptions.Compiled), SourceDocumentType.CashPayment),
        (new(@"Расходный\s+кассовый\s+ордер",               RegexOptions.IgnoreCase | RegexOptions.Compiled), SourceDocumentType.CashExpense),
        (new(@"Заказ\s+покупателя",                         RegexOptions.IgnoreCase | RegexOptions.Compiled), SourceDocumentType.BuyerOrder),
        (new(@"Чек\s+ККМ",                                  RegexOptions.IgnoreCase | RegexOptions.Compiled), SourceDocumentType.KkmCheck),
    };

    // ── Публичное API ───────────────────────────────────────────────────────────

    /// <summary>Парсит весь файл и возвращает только активные кейсы (`- [ ]`).</summary>
    public static List<OrderEntry> ParseFile(string mdPath)
    {
        if (!File.Exists(mdPath)) return new();
        return ParseText(File.ReadAllText(mdPath));
    }

    /// <summary>Парсит произвольный текст со строками-кейсами.</summary>
    public static List<OrderEntry> ParseText(string text)
    {
        var result = new List<OrderEntry>();
        if (string.IsNullOrWhiteSpace(text)) return result;

        foreach (var rawLine in text.Split('\n'))
        {
            var line = rawLine.TrimEnd('\r').Trim();
            if (line.Length == 0) continue;

            // Заголовки «## СЕНТЯБРЬ 2025», «=====», ссылки и т.д. — пропускаем
            if (!line.StartsWith("-")) continue;

            var taskMatch = RxTaskLine.Match(line);
            if (!taskMatch.Success) continue;

            var done = taskMatch.Groups["done"].Value.Trim().ToLowerInvariant() == "x";
            if (done) continue;   // [x] — уже сделано, пропускаем

            var body = taskMatch.Groups["body"].Value.Trim();
            var entry = ParseBody(body);
            if (entry is not null) result.Add(entry);
        }

        return result;
    }

    // ── Внутренняя логика разбора одной строки ──────────────────────────────────

    private static OrderEntry? ParseBody(string body)
    {
        // 1. Особый случай: «ФП: NNNN ...»
        var fpOnly = RxFpOnly.Match(body);
        if (fpOnly.Success)
        {
            var city   = ExtractCity(body);
            var notes  = ExtractNotes(body, "", "", city);
            var fpAmt  = ExtractAmount(body);

            return new OrderEntry
            {
                Kind                 = OrderKind.SingleRefund,
                DocumentType         = SourceDocumentType.FpOnly,
                CorrectionScenario   = CorrectionScenario.FullCancel,
                OriginalFiscalNumber = fpOnly.Groups["fp"].Value,
                Notes                = notes,
                City                 = city,
                Amount               = fpAmt,
                OriginalCheckAmount  = fpAmt > 0 ? fpAmt : null,
            };
        }

        // 2. Определяем тип документа
        var docType = SourceDocumentType.Unknown;
        foreach (var (pat, type) in DocTypePatterns)
        {
            if (pat.IsMatch(body)) { docType = type; break; }
        }

        if (docType == SourceDocumentType.Unknown)
            return null;  // не смогли распознать — пропускаем

        // 3. Номер документа (тXXXXXXX)
        var numMatch = RxDocNumber.Match(body);
        var docNumber = numMatch.Success ? numMatch.Value : string.Empty;

        // 4. Дата/время
        var dtMatch = RxDateTime.Match(body);
        var date = string.Empty;
        if (dtMatch.Success)
        {
            date = dtMatch.Groups["d"].Value;
            if (dtMatch.Groups["t"].Success)
            {
                var t = dtMatch.Groups["t"].Value;
                // нормализуем часы: 9:36:01 → 09:36:01
                if (t.Length < 8 && t[1] == ':') t = "0" + t;
                if (t.Length == 5) t += ":00"; // HH:MM → HH:MM:00
                date = $"{date} {t}";
            }
        }

        // 5. Город
        var cityName = ExtractCity(body);

        // 6. Сумма (если указана в тексте: «сумма 35242,00» или «11030 руб»)
        var amt = ExtractAmount(body);

        // 7. Описание (что после города/даты/суммы)
        var notesText = ExtractNotes(body, docNumber, date.Split(' ').FirstOrDefault() ?? "", cityName);

        // 8. Заполняем entry — Kind/Scenario определит детектор позже
        return new OrderEntry
        {
            DocumentType        = docType,
            OrderNum            = docNumber,
            OrderDate           = date,
            City                = cityName,
            Notes               = notesText,
            Amount              = amt,
            OriginalCheckAmount = amt > 0 ? amt : null,
            Kind                = OrderKind.Regular,  // временно, заполнит детектор
        };
    }

    /// <summary>Пытается извлечь сумму из текста: «сумма 35 242,00», «11030 руб», «сумма 1500,00».</summary>
    private static double ExtractAmount(string body)
    {
        var m = RxAmountInline.Match(body);
        if (!m.Success) return 0;
        var raw = (m.Groups["a"].Success ? m.Groups["a"].Value : m.Groups["r"].Value)
            .Replace(" ", "").Replace("\xa0", "").Replace(",", ".");
        return double.TryParse(raw, System.Globalization.NumberStyles.Any,
            System.Globalization.CultureInfo.InvariantCulture, out var v) ? v : 0;
    }

    private static string ExtractCity(string body)
    {
        // Берём ПОСЛЕДНЕЕ совпадение в скобках — в Obsidian город обычно в конце
        // (но «(кредит)», «(сборка)» в начале не считаем городом)
        var matches = RxCity.Matches(body);
        for (int i = matches.Count - 1; i >= 0; i--)
        {
            var v = matches[i].Groups[1].Value.Trim();
            // Фильтруем не-города
            if (v.Equals("кредит",   StringComparison.OrdinalIgnoreCase)) continue;
            if (v.Equals("сборка",   StringComparison.OrdinalIgnoreCase)) continue;
            if (v.Equals("доставка", StringComparison.OrdinalIgnoreCase)) continue;
            return v;
        }
        return string.Empty;
    }

    /// <summary>Достаёт «описание» — то что осталось после удаления номера/даты/города.</summary>
    private static string ExtractNotes(string body, string docNumber, string date, string city)
    {
        var s = body;
        // Убираем «Реализация товаров и услуг» / «Оплата...» — в начале
        foreach (var (pat, _) in DocTypePatterns)
            s = pat.Replace(s, "", 1);
        if (!string.IsNullOrEmpty(docNumber)) s = s.Replace(docNumber, "");
        if (!string.IsNullOrEmpty(date))      s = s.Replace(date, "");
        if (!string.IsNullOrEmpty(city))      s = s.Replace($"({city})", "");
        // Убираем «от dd.mm.yyyy» хвосты времени
        s = Regex.Replace(s, @"от\s+\d{1,2}[:.]?\d{0,2}([:.]\d{0,2})?", "", RegexOptions.IgnoreCase);
        s = Regex.Replace(s, @"\d{1,2}:\d{2}(?::\d{2})?", "");
        s = Regex.Replace(s, @"\s{2,}", " ");
        return s.Trim(' ', ',', '.', '—', '-');
    }
}
