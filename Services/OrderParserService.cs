using System.Text.RegularExpressions;
using AtolGenerator.Constants;
using AtolGenerator.Models;

namespace AtolGenerator.Services;

public static class OrderParserService
{
    // Номер заказа
    private static readonly Regex RxNum = new(@"[тТ](\d{10})", RegexOptions.Compiled);

    // Дата: "от ДД.ММ.ГГГГ [ЧЧ:ММ:СС]" — время необязательно
    private static readonly Regex RxDate = new(
        @"от\s+(\d{2}\.\d{2}\.\d{4})(?:\s+(\d{2}:\d{2}:\d{2}))?",
        RegexOptions.Compiled);

    // Сумма
    private static readonly Regex RxAmt = new(
        @"сумм[аыу]\s*([\d\s]+(?:[,\.]\d+)?)",
        RegexOptions.Compiled | RegexOptions.IgnoreCase);

    // ФИО: "покупателя/ФИО: Фамилия Имя Отчество"
    private static readonly Regex RxFio1 = new(
        @"(?:покупател[яь]|ФИО)[:\s]+([А-ЯЁ][а-яё]+\s+[А-ЯЁ][а-яё]+(?:\s+[А-ЯЁ][а-яё]+)?)",
        RegexOptions.Compiled);

    // ФИО в скобках: "(Фамилия Имя Отчество)" — последние скобки со словами заглавными буквами
    private static readonly Regex RxFioParens = new(
        @"\(([А-ЯЁ][а-яё]+ [А-ЯЁ][а-яё]+(?: [А-ЯЁ][а-яё]+)?)\)\s*$",
        RegexOptions.Compiled);

    // ФИО в конце строки
    private static readonly Regex RxFio2 = new(
        @"([А-ЯЁ][а-яё]+\s+[А-ЯЁ][а-яё]+\s+[А-ЯЁ][а-яё]+)\s*$",
        RegexOptions.Compiled);

    // Тип услуги: "(сборка)", "(доставка)" — определяет IsService = true
    private static readonly Regex RxServiceType = new(
        @"\((сборка|доставка)\)",
        RegexOptions.Compiled | RegexOptions.IgnoreCase);

    // Город из скобок: "(Ярославль)", "(Иваново сборка)" и т.п. перед суммой
    private static readonly Regex RxCity = new(
        @"\(([А-ЯЁа-яё\s]+?)\)",
        RegexOptions.Compiled);

    private static readonly Regex RxNonDigit = new(@"[^\d.]", RegexOptions.Compiled);

    public static List<OrderEntry> Parse(string text)
    {
        var results = new List<OrderEntry>();

        foreach (var raw in text.Split(new[] { ';', '\n' }, StringSplitOptions.RemoveEmptyEntries))
        {
            var part = raw.Trim();
            if (string.IsNullOrEmpty(part)) continue;

            // Номер заказа
            var mNum = RxNum.Match(part);
            if (!mNum.Success) continue;
            var orderNum = "т" + mNum.Groups[1].Value;

            // Дата (время необязательно)
            var mDate = RxDate.Match(part);
            if (!mDate.Success) continue;
            var timeStr = mDate.Groups[2].Success ? mDate.Groups[2].Value : "00:00:00";
            var orderDate = $"{mDate.Groups[1].Value} {timeStr}";

            // Сумма
            var mAmt = RxAmt.Match(part);
            if (!mAmt.Success) continue;
            var amtRaw = mAmt.Groups[1].Value.Trim().Replace(" ", "").Replace(",", ".");
            amtRaw = RxNonDigit.Split(amtRaw)[0];
            if (!double.TryParse(amtRaw, System.Globalization.NumberStyles.Any,
                                 System.Globalization.CultureInfo.InvariantCulture, out var amount))
                continue;

            // Тип услуги: (сборка) / (доставка) → IsService = true
            var mSvcType = RxServiceType.Match(part);
            var isService   = mSvcType.Success;
            var serviceType = isService ? mSvcType.Groups[1].Value.ToLower() : string.Empty;

            // Город — берём из скобок НЕ совпадающих с типом услуги и не ФИО
            // Ищем все скобочные выражения
            var city = string.Empty;
            if (isService)
            {
                foreach (Match mc in RxCity.Matches(part))
                {
                    var val = mc.Groups[1].Value.Trim();
                    // Пропускаем если это тип услуги или ФИО (содержит 2+ слова с заглавными)
                    if (string.Equals(val, serviceType, StringComparison.OrdinalIgnoreCase)) continue;
                    if (val.Split(' ').Count(w => w.Length > 0 && char.IsUpper(w[0])) >= 2) continue;
                    city = val;
                    break;
                }
            }

            // Поставщик услуги (если известны тип + город)
            ServiceProvider? agent = null;
            if (isService && !string.IsNullOrEmpty(serviceType) && !string.IsNullOrEmpty(city))
            {
                agent = AppConstants.ServiceProviders
                    .FirstOrDefault(p =>
                        string.Equals(p.Service, serviceType, StringComparison.OrdinalIgnoreCase) &&
                        city.Contains(p.City, StringComparison.OrdinalIgnoreCase));
            }

            // ФИО покупателя — пробуем несколько паттернов
            var customerName = string.Empty;
            var mFio = RxFio1.Match(part);
            if (mFio.Success)
            {
                customerName = mFio.Groups[1].Value.Trim();
            }
            else
            {
                // "(Фамилия Имя Отчество)" в конце части
                var mFioP = RxFioParens.Match(part);
                if (mFioP.Success)
                    customerName = mFioP.Groups[1].Value.Trim();
                else
                {
                    var mFio2 = RxFio2.Match(part);
                    if (mFio2.Success)
                        customerName = mFio2.Groups[1].Value.Trim();
                }
            }

            results.Add(new OrderEntry
            {
                OrderNum     = orderNum,
                OrderDate    = orderDate,
                Amount       = amount,
                CustomerName = customerName,
                IsService    = isService,
                City         = city,
                AgentInfo    = agent,
            });
        }

        return results;
    }
}
