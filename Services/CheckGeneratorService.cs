using System.IO;
using System.Text.RegularExpressions;
using AtolGenerator.Constants;
using AtolGenerator.Helpers;
using AtolGenerator.Models;

namespace AtolGenerator.Services;

public static class CheckGeneratorService
{
    private static readonly Regex RxDdMmYyyy = new(@"^\d{2}\.\d{2}\.\d{4}", RegexOptions.Compiled);

    public static List<GenerationResult> Generate(GenerationParams p)
    {
        Directory.CreateDirectory(p.OutputDir);
        var isCorrection = p.CheckType is "sell_correction" or "buy_correction";
        var results = new List<GenerationResult>();

        foreach (var order in p.Orders)
        {
            var amount      = order.Amount;
            var agentInfo   = order.AgentInfo;
            var corrDateRaw = order.CorrectionDate;
            // IsService: из заказа (импорт Excel) или из параметров генерации (ручной ввод)
            bool orderIsService = order.IsService || agentInfo is not null || p.IsService;

            // ── Build items ──
            var items = new List<CheckItem>();
            if (!isCorrection)
            {
                if (p.Tab == "payment")
                {
                    bool isService = orderIsService;
                    var name = $"Аванс от покупателя по заказу № {order.OrderNum}";
                    if (!string.IsNullOrEmpty(order.OrderDate))
                        name += $" от {order.OrderDate}";
                    items.Add(new CheckItem
                    {
                        Name          = name,
                        Price         = amount,
                        Quantity      = 1,
                        Sum           = amount,
                        // full_payment в итоге даёт только БЕЗНАЛИЧНЫМИ/НАЛИЧНЫМИ без строки АВАНС
                        PaymentMethod = "full_payment",
                        PaymentObject = "payment",
                        VatType       = isService ? "none" : "vat122",
                        VatSum        = isService ? amount : CalcVat122(amount),
                        IsService     = isService,
                    });
                }
                else
                {
                    foreach (var raw in order.Items)
                    {
                        var qty   = raw.Quantity > 0 ? raw.Quantity : 1;
                        var s     = raw.Sum;
                        var price = Math.Round(s / qty, 2);
                        bool isService = orderIsService;
                        items.Add(new CheckItem
                        {
                            Name          = raw.Name,
                            Price         = price,
                            Quantity      = qty,
                            Sum           = s,
                            PaymentMethod = "full_payment",
                            PaymentObject = "commodity",
                            VatType       = isService ? "none" : "vat22",
                            VatSum        = isService ? s : CalcVat20(s),
                            IsService     = isService,
                        });
                    }
                }
            }

            // CheckVatType для итога чека: услуга/агент → none; оплата → vat122; реализация → vat22
            var checkVatType = orderIsService        ? "none"
                             : p.Tab == "payment"   ? "vat122"
                             :                        "vat22";

            // ── Build CheckData for XML ──
            var checkData = new CheckData
            {
                OperationType = p.CheckType,
                IsCorrection  = isCorrection,
                Tab           = p.Tab,
                Amount        = amount,
                PaymentType   = p.PaymentType,
                CheckVatType  = checkVatType,
                Items         = items,
                Agent         = agentInfo,
                IsService     = orderIsService,
                CashierName   = p.Cashier.FullName,
                CashierShort  = p.Cashier.ShortName,
            };

            if (isCorrection)
            {
                checkData.CorrectionBaseDate = string.IsNullOrEmpty(corrDateRaw)
                    ? DateTime.Today.ToString("yyyy-MM-dd")
                    : RxDdMmYyyy.IsMatch(corrDateRaw)
                        ? ToIsoDate(corrDateRaw)
                        : corrDateRaw;

                var manual = corrDateRaw?.Trim() ?? string.Empty;
                var corrNum = order.CorrectionNumber?.Trim() ?? string.Empty;
                if (!string.IsNullOrEmpty(corrNum) && corrNum != "б/н")
                    checkData.CorrectionBaseNumber = corrNum;
                else if (!string.IsNullOrEmpty(order.OrderNum))
                    checkData.CorrectionBaseNumber = $"Основание_{order.OrderNum}";
                else
                    checkData.CorrectionBaseNumber = "б/н";
            }

            // ── Имена файлов (используются для DOCX и при раздельном XML) ──
            var ts       = DateTime.Now.ToString("yyyyMMdd_HHmmss_fff")[..17];
            var safeNum  = FileHelper.SafeFilename(order.OrderNum);
            var baseName = $"{ts}_{safeNum}_{p.CheckType}";

            var result = new GenerationResult
            {
                OrderNum  = order.OrderNum,
                Amount    = amount,
                CheckData = checkData,
                BaseName  = baseName,
            };

            // ── DOCX (только для коррекций, всегда отдельный файл) ──
            if (isCorrection)
            {
                var docxPath = Path.Combine(p.OutputDir, $"{baseName}_служебка.docx");

                var orderInfoStr = order.OrderNum;
                if (!string.IsNullOrEmpty(order.OrderDate))
                    orderInfoStr += $" от {order.OrderDate} (Заказ покупателя)";

                var eventDate = !string.IsNullOrEmpty(corrDateRaw)
                    ? corrDateRaw
                    : !string.IsNullOrEmpty(order.OrderDate)
                        ? order.OrderDate.Split(' ')[0]
                        : DateTime.Today.ToString("dd.MM.yyyy");

                // Текст «при ...»
                var operationDesc = p.Tab == "realization"
                    ? (orderIsService
                        ? "реализации услуги по перевозке/сборки товара покупателю"
                        : "реализации товара покупателю")
                    : p.CheckType == "buy_correction"
                        ? "возврате безналичных денежных средств покупателю"
                        : AppConstants.OperationDescriptions.GetValueOrDefault(p.CheckType, "расчёте");

                // ККТ блок 4 (не пробит чек): по городу/подразделению для реализаций,
                //   для оплат — Интернет-магазин (DefaultKkt).
                // ККТ блок 5 (пробита коррекция): всегда Интернет-магазин (DefaultKkt).
                var storeKkt  = p.Tab == "realization"
                    ? AppConstants.GetKktByCity(order.City)
                    : AppConstants.DefaultKkt;
                var onlineKkt = AppConstants.DefaultKkt;

                var memo = new MemoData
                {
                    EventDate        = eventDate,
                    TodayDate        = DateTime.Today.ToString("dd.MM.yyyy"),
                    OperationDesc    = operationDesc,
                    CustomerName     = order.CustomerName,
                    Amount           = amount,
                    OrderInfo        = orderInfoStr,
                    CorrectionDesc   = AppConstants.CorrectionDescriptions.GetValueOrDefault(p.CheckType, string.Empty),
                    FromPosition     = p.Cashier.Position,
                    FromNameGenitive = p.Cashier.NameGenitive,
                    CashierShort     = p.Cashier.ShortName,
                    // Блок 4 — касса магазина
                    KktModel         = storeKkt.Model,
                    KktSerial        = storeKkt.Serial,
                    KktReg           = storeKkt.RegNum,
                    KktFfd           = storeKkt.Ffd,
                    // Блок 5 — касса интернет-магазина
                    KktModelOnline   = onlineKkt.Model,
                    KktSerialOnline  = onlineKkt.Serial,
                    KktRegOnline     = onlineKkt.RegNum,
                    KktFfdOnline     = onlineKkt.Ffd,
                };

                DocxGeneratorService.Generate(memo, docxPath);
                result.DocxPath = docxPath;
            }

            results.Add(result);
        }

        // ── Запись XML ──
        if (p.MergeXml && results.Count > 0)
        {
            // Один файл для всех чеков
            var ts      = DateTime.Now.ToString("yyyyMMdd_HHmmss_fff")[..17];
            var xmlPath = Path.Combine(p.OutputDir, $"{ts}_{results.Count}чеков_{p.CheckType}.xml");
            XmlGeneratorService.GenerateFile(results.Select(r => r.CheckData!), xmlPath);
            foreach (var r in results) r.XmlPath = xmlPath;
        }
        else
        {
            // Отдельный XML для каждого чека
            foreach (var r in results)
            {
                var xmlPath = Path.Combine(p.OutputDir, $"{r.BaseName}.xml");
                XmlGeneratorService.GenerateFile(new[] { r.CheckData! }, xmlPath);
                r.XmlPath = xmlPath;
            }
        }

        return results;
    }

    private static double CalcVat122(double amount) => Math.Round(amount * 22.0 / 122.0, 2);
    private static double CalcVat20(double amount)  => Math.Round(amount * 0.22, 2);

    private static string ToIsoDate(string ddMmYyyy)
    {
        if (DateTime.TryParseExact(ddMmYyyy.Trim(), "dd.MM.yyyy",
            System.Globalization.CultureInfo.InvariantCulture,
            System.Globalization.DateTimeStyles.None, out var d))
            return d.ToString("yyyy-MM-dd");
        return ddMmYyyy;
    }
}

public class GenerationParams
{
    public string            Tab         { get; set; } = "payment";
    public string            CheckType   { get; set; } = "sell";
    public string            PaymentType { get; set; } = "card";
    public bool              IsService   { get; set; } = false;
    public bool              MergeXml    { get; set; } = false;
    public List<OrderEntry>  Orders      { get; set; } = new();
    public string            OutputDir   { get; set; } = FileHelper.OutputDir;
    public CashierInfo       Cashier     { get; set; } = AppConstants.DefaultCashier;
}
