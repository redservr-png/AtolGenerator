namespace AtolGenerator.Models;

public sealed record VatRateOption(string Code, string Label)
{
    public override string ToString() => Label;
}

public static class VatRateCatalog
{
    public static IReadOnlyList<VatRateOption> All { get; } = new[]
    {
        new VatRateOption("none", "Без НДС"),
        new VatRateOption("vat0", "НДС 0%"),
        new VatRateOption("vat10", "НДС 10%"),
        new VatRateOption("vat20", "НДС 20%"),
        new VatRateOption("vat110", "НДС 10/110"),
        new VatRateOption("vat120", "НДС 20/120"),
        new VatRateOption("vat5", "НДС 5%"),
        new VatRateOption("vat105", "НДС 5/105"),
        new VatRateOption("vat7", "НДС 7%"),
        new VatRateOption("vat107", "НДС 7/107"),
        new VatRateOption("vat22", "НДС 22%"),
        new VatRateOption("vat122", "НДС 22/122"),
    };

    private static readonly IReadOnlyDictionary<string, VatRateOption> ByCode = All
        .ToDictionary(x => x.Code, StringComparer.OrdinalIgnoreCase);

    public static bool IsKnown(string? code) =>
        !string.IsNullOrWhiteSpace(code) && ByCode.ContainsKey(code.Trim());

    public static string Normalize(string? code, string fallback = "none") =>
        IsKnown(code) ? code!.Trim().ToLowerInvariant() : fallback;

    public static string LabelFor(string? code) =>
        !string.IsNullOrWhiteSpace(code) && ByCode.TryGetValue(code.Trim(), out var option)
            ? option.Label
            : code ?? string.Empty;

    public static double Calculate(double grossAmount, string? code)
    {
        var amount = Math.Abs(grossAmount);
        var (rate, denominator) = Normalize(code) switch
        {
            "vat5" or "vat105" => (5.0, 105.0),
            "vat7" or "vat107" => (7.0, 107.0),
            "vat10" or "vat110" => (10.0, 110.0),
            "vat20" or "vat120" => (20.0, 120.0),
            "vat22" or "vat122" => (22.0, 122.0),
            _ => (0.0, 1.0),
        };

        return rate == 0 ? 0 : Math.Round(amount * rate / denominator, 2);
    }

    /// <summary>
    /// Значение поля sum в VAT-группе АТОЛ. Для none и vat0 это сумма расчёта,
    /// относящаяся к ставке, для остальных типов — сумма налога.
    /// </summary>
    public static double CalculateFiscalSum(double grossAmount, string? code)
    {
        var normalized = Normalize(code);
        return normalized is "none" or "vat0"
            ? Math.Round(Math.Abs(grossAmount), 2)
            : Calculate(grossAmount, normalized);
    }
}
