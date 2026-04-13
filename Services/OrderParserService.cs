using System.Text.RegularExpressions;
using AtolGenerator.Models;

namespace AtolGenerator.Services;

public static class OrderParserService
{
    private static readonly Regex RxNum  = new(@"[тТ](\d{10})", RegexOptions.Compiled);
    private static readonly Regex RxDate = new(@"от\s+(\d{2}\.\d{2}\.\d{4})\s+(\d{2}:\d{2}:\d{2})", RegexOptions.Compiled);
    private static readonly Regex RxAmt  = new(@"сумм[аыу]\s*([\d\s]+(?:[,\.]\d+)?)", RegexOptions.Compiled | RegexOptions.IgnoreCase);
    private static readonly Regex RxFio1 = new(@"(?:покупател[яь]|ФИО)[:\s]+([А-ЯЁ][а-яё]+\s+[А-ЯЁ][а-яё]+(?:\s+[А-ЯЁ][а-яё]+)?)", RegexOptions.Compiled);
    private static readonly Regex RxFio2 = new(@"([А-ЯЁ][а-яё]+\s+[А-ЯЁ][а-яё]+\s+[А-ЯЁ][а-яё]+)\s*$", RegexOptions.Compiled);
    private static readonly Regex RxNonDigit = new(@"[^\d.]", RegexOptions.Compiled);

    public static List<OrderEntry> Parse(string text)
    {
        var results = new List<OrderEntry>();

        foreach (var raw in text.Split(new[] { ';', '\n' }, StringSplitOptions.RemoveEmptyEntries))
        {
            var part = raw.Trim();
            if (string.IsNullOrEmpty(part)) continue;

            var mNum = RxNum.Match(part);
            if (!mNum.Success) continue;
            var orderNum = "т" + mNum.Groups[1].Value;

            var mDate = RxDate.Match(part);
            // Требуем наличие точного времени ЧЧ:ММ:СС — заказы без него (доставка и т.п.) пропускаются
            if (!mDate.Success) continue;
            var orderDate = $"{mDate.Groups[1].Value} {mDate.Groups[2].Value}";

            var mAmt = RxAmt.Match(part);
            if (!mAmt.Success) continue;
            var amtRaw = mAmt.Groups[1].Value.Trim()
                             .Replace(" ", "")
                             .Replace(",", ".");
            amtRaw = RxNonDigit.Split(amtRaw)[0];
            if (!double.TryParse(amtRaw, System.Globalization.NumberStyles.Any,
                                 System.Globalization.CultureInfo.InvariantCulture, out var amount))
                continue;

            var customerName = string.Empty;
            var mFio = RxFio1.Match(part);
            if (mFio.Success)
            {
                customerName = mFio.Groups[1].Value.Trim();
            }
            else
            {
                var mFio2 = RxFio2.Match(part);
                if (mFio2.Success)
                    customerName = mFio2.Groups[1].Value.Trim();
            }

            results.Add(new OrderEntry
            {
                OrderNum     = orderNum,
                OrderDate    = orderDate,
                Amount       = amount,
                CustomerName = customerName,
            });
        }

        return results;
    }
}
