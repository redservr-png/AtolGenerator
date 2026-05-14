using System.Text;
using AtolGenerator.Constants;

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
                    : !c.IsService             ? "Аванс:"
                    : "Безналичными:";
        sb.AppendLine(PadRight(payName, FmtMoney(c.Amount) + " руб."));
        sb.AppendLine(line);

        // ── НДС ───────────────────────────────────────────────────────────────
        if (c.CheckVatType != "none")
        {
            var totalVat = c.Items.Count > 0
                ? c.Items.Sum(i => i.VatSum)
                : CalcVat(c.CheckVatType, c.Amount);
            if (totalVat > 0)
                sb.AppendLine(PadRight(VatTypeName(c.CheckVatType) + ":", FmtMoney(totalVat) + " руб."));
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

    private static string VatTypeName(string vat) => vat switch
    {
        "vat5"   => "НДС 5%",
        "vat105" => "НДС 5% (вкл.)",
        "vat22"  => "НДС 22%",
        "vat122" => "НДС 22% (вкл.)",
        "none"   => "Без НДС",
        _        => vat,
    };

    private static string VatLine(string vatType, double vatSum)
    {
        if (vatType == "none" || vatSum == 0) return string.Empty;
        return $"{VatTypeName(vatType)}: {FmtMoney(vatSum)} руб.";
    }

    private static double CalcVat(string vatType, double amount) => vatType switch
    {
        "vat5"   => Math.Round(amount * 5.0  / 100.0, 2),
        "vat105" => Math.Round(amount * 5.0  / 105.0, 2),
        "vat22"  => Math.Round(amount * 22.0 / 122.0, 2),
        "vat122" => Math.Round(amount * 22.0 / 122.0, 2),
        _        => 0,
    };

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
