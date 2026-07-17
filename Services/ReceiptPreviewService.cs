using System.Text;
using AtolGenerator.Constants;
using AtolGenerator.Models;

namespace AtolGenerator.Services;

/// <summary>
/// Генерирует текстовый «бумажный» предварительный просмотр чека (monospace).
/// </summary>
public static class ReceiptPreviewService
{
    private const int W = 44;  // ширина чека в символах

    public static string Generate(IEnumerable<CheckData> checks)
    {
        var sb = new StringBuilder();
        foreach (var c in checks)
        {
            if (sb.Length > 0) sb.AppendLine();
            sb.Append(GenerateOne(c));
        }
        return sb.ToString();
    }

    private static string GenerateOne(CheckData c)
    {
        var sb = new StringBuilder();
        var line = new string('─', W);

        // ── Шапка ─────────────────────────────────────────────────────────────
        sb.AppendLine(Center("ИП Шевелева Е.Н."));
        sb.AppendLine(Center(AppConstants.City));
        sb.AppendLine(Center($"ИНН {AppConstants.InnOrg}"));
        sb.AppendLine(Center(AppConstants.PaymentAddress));
        sb.AppendLine(line);

        // ── Тип операции ──────────────────────────────────────────────────────
        sb.AppendLine(Center(OperationName(c.OperationType)));
        sb.AppendLine(line);

        // ── Позиции ───────────────────────────────────────────────────────────
        if (c.Items.Count > 0)
        {
            foreach (var item in c.Items)
            {
                // Длинное название переносим
                var nameLine = item.Name.Length > W ? item.Name[..W] : item.Name;
                sb.AppendLine(nameLine);
                var qtyPrice = $"  {FmtNum(item.Quantity)} × {FmtMoney(item.Price)}";
                sb.AppendLine(PadRight(qtyPrice, FmtMoney(item.Sum)));
                if (!string.IsNullOrEmpty(VatLine(item.VatType, item.VatSum)))
                    sb.AppendLine($"  {VatLine(item.VatType, item.VatSum)}");
            }
            sb.AppendLine(line);
        }

        // ── Итог / оплата ─────────────────────────────────────────────────────
        sb.AppendLine(PadRight("ИТОГО:", FmtMoney(c.Amount) + " руб."));

        var payName = c.Tab == "realization" ? "Предоплата (аванс):"
                    : c.PaymentType == "cash" ? "Наличными:"
                    : "Безналичными:";
        sb.AppendLine(PadRight(payName, FmtMoney(c.Amount) + " руб."));
        sb.AppendLine(line);

        // ── НДС ───────────────────────────────────────────────────────────────
        var vatGroups = c.Items.Count > 0
            ? c.Items
                .GroupBy(item => VatRateCatalog.Normalize(item.VatType, c.CheckVatType),
                    StringComparer.OrdinalIgnoreCase)
                .Select(group => new
                {
                    Type = group.Key,
                    Sum = group.Sum(item => VatRateCatalog.Calculate(item.Sum, group.Key)),
                })
            : new[]
            {
                new
                {
                    Type = VatRateCatalog.Normalize(c.CheckVatType, "none"),
                    Sum = VatRateCatalog.Calculate(c.Amount, c.CheckVatType),
                },
            };
        foreach (var vat in vatGroups.Where(vat => vat.Type is not "none" and not "vat0"))
        {
            sb.AppendLine(PadRight(VatRateCatalog.LabelFor(vat.Type) + ":",
                FmtMoney(vat.Sum) + " руб."));
        }
        sb.AppendLine(line);

        // ── Служебные поля ────────────────────────────────────────────────────
        sb.AppendLine($"Кассир: {c.CashierShort}");
        if (!string.IsNullOrEmpty(c.AdditionalCheckProps))
            sb.AppendLine($"ФП исх. чека: {c.AdditionalCheckProps}");
        if (c.IsCorrection)
        {
            sb.AppendLine($"Осн. дата: {c.CorrectionBaseDate}");
            sb.AppendLine($"Осн. номер: {c.CorrectionBaseNumber}");
        }
        sb.Append(line);

        return sb.ToString();
    }

    // ── Helpers ──────────────────────────────────────────────────────────────

    private static string OperationName(string op) => op switch
    {
        "sell"            => "ПРИХОД",
        "sell_refund"     => "ВОЗВРАТ ПРИХОДА",
        "buy"             => "РАСХОД",
        "buy_refund"      => "ВОЗВРАТ РАСХОДА",
        "sell_correction" => "КОРРЕКЦИЯ ПРИХОДА",
        "buy_correction"  => "КОРРЕКЦИЯ РАСХОДА",
        _                 => op.ToUpperInvariant(),
    };

    private static string VatLine(string vatType, double vatSum)
    {
        if (vatType is "none" or "vat0" || vatSum == 0) return string.Empty;
        return $"{VatRateCatalog.LabelFor(vatType)}: {FmtMoney(vatSum)} руб.";
    }

    private static string Center(string text)
    {
        if (text.Length >= W) return text[..W];
        var pad = (W - text.Length) / 2;
        return text.PadLeft(text.Length + pad).PadRight(W);
    }

    private static string PadRight(string left, string right)
    {
        var gap = W - left.Length - right.Length;
        return gap > 0
            ? left + new string(' ', gap) + right
            : left + " " + right;
    }

    private static string FmtMoney(double v) =>
        v.ToString("N2", System.Globalization.CultureInfo.GetCultureInfo("ru-RU"));

    private static string FmtNum(double v) =>
        v == Math.Floor(v) ? ((int)v).ToString() : v.ToString("F3").TrimEnd('0');
}
