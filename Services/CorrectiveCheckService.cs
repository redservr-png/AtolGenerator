using System.IO;
using AtolGenerator.Constants;
using AtolGenerator.Helpers;
using AtolGenerator.Models;

namespace AtolGenerator.Services;

/// <summary>
/// Формирует пару исправительных чеков для реализации, по которой уже пробит чек:
///   1) sell_refund   — возврат прихода с позициями из 1С + тег 1192 (ФП исходного чека)
///   2) sell_correction — коррекция прихода (самостоятельная)
/// </summary>
public static class CorrectiveCheckService
{
    public static List<GenerationResult> Generate(
        OneCRealization real,
        List<OneCRealizationItem> items,
        string outputDir,
        CashierInfo? cashier = null)
    {
        Directory.CreateDirectory(outputDir);
        cashier ??= AppConstants.DefaultCashier;
        var results = new List<GenerationResult>();

        var ts      = DateTime.Now.ToString("yyyyMMdd_HHmmss_fff")[..17];
        var safeNum = FileHelper.SafeFilename(real.DocNumber);

        // ── 1. sell_refund: возврат прихода с позициями + ФП ──────────────────
        var refundItems = items.Select(i =>
        {
            var qty   = Math.Max(i.Quantity, 0.001);
            var price = Math.Round(i.Sum / qty, 2);
            return new CheckItem
            {
                Name          = i.Name,
                Price         = price,
                Quantity      = qty,
                Sum           = i.Sum,
                PaymentMethod = "full_payment",
                PaymentObject = real.IsService ? "service" : "commodity",
                VatType       = real.IsService ? "none" : "vat22",
                VatSum        = real.IsService ? 0 : Math.Round(i.Sum * 22.0 / 122.0, 2),
                IsService     = real.IsService,
            };
        }).ToList();

        var refundData = new CheckData
        {
            OperationType        = "sell_refund",
            IsCorrection         = false,
            Tab                  = "realization",
            Amount               = real.Amount,
            PaymentType          = "card",
            CheckVatType         = real.IsService ? "none" : "vat22",
            Items                = refundItems,
            IsService            = real.IsService,
            AdditionalCheckProps = real.FiscalNumber,
            CashierName          = cashier.FullName,
            CashierShort         = cashier.ShortName,
        };

        var refundName = $"{ts}_{safeNum}_sell_refund";
        var refundXml  = Path.Combine(outputDir, $"{refundName}.xml");
        XmlGeneratorService.GenerateFile(new[] { refundData }, refundXml);

        results.Add(new GenerationResult
        {
            OrderNum  = real.OrderNumber,
            Amount    = real.Amount,
            CheckData = refundData,
            BaseName  = refundName,
            XmlPath   = refundXml,
        });

        // ── 2. sell_correction: коррекция прихода (самостоятельная) ───────────
        // Дата основания — дата печати оригинального чека (или дата реализации)
        var corrRawDate = !string.IsNullOrEmpty(real.CheckDate)
            ? real.CheckDate.Split(' ')[0]
            : !string.IsNullOrEmpty(real.DocDate)
                ? real.DocDate
                : DateTime.Today.ToString("dd.MM.yyyy");

        var corrDateIso = TryToIso(corrRawDate);

        // Номер основания — номер реализации (НомерДок)
        var corrNumber = !string.IsNullOrEmpty(real.DocNumber) ? real.DocNumber : "б/н";

        var corrData = new CheckData
        {
            OperationType        = "sell_correction",
            IsCorrection         = true,
            Tab                  = "realization",
            Amount               = real.Amount,
            PaymentType          = "card",
            CheckVatType         = real.IsService ? "none" : "vat22",
            Items                = new List<CheckItem>(),
            IsService            = real.IsService,
            CorrectionBaseDate   = corrDateIso,
            CorrectionBaseNumber = corrNumber,
            CashierName          = cashier.FullName,
            CashierShort         = cashier.ShortName,
        };

        var corrName = $"{ts}_{safeNum}_sell_correction";
        var corrXml  = Path.Combine(outputDir, $"{corrName}.xml");
        XmlGeneratorService.GenerateFile(new[] { corrData }, corrXml);

        results.Add(new GenerationResult
        {
            OrderNum  = real.OrderNumber,
            Amount    = real.Amount,
            CheckData = corrData,
            BaseName  = corrName,
            XmlPath   = corrXml,
        });

        return results;
    }

    private static string TryToIso(string ddMmYyyy)
    {
        if (DateTime.TryParseExact(ddMmYyyy.Trim(), "dd.MM.yyyy",
            System.Globalization.CultureInfo.InvariantCulture,
            System.Globalization.DateTimeStyles.None, out var d))
            return d.ToString("yyyy-MM-dd");
        return ddMmYyyy;
    }
}
