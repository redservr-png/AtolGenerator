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
            ValidateCorrectionVat(order);
            ValidateRepairPairItems(order);

            // ── Исправительные кейсы из Obsidian: делегируем спецсервису ──
            if (order.IsCorrection)
            {
                var checks = CorrectionGeneratorService.BuildCheckDataList(order, p);
                var orderResults = new List<GenerationResult>();
                foreach (var cd in checks)
                {
                    if (string.IsNullOrWhiteSpace(cd.ExternalId))
                        cd.ExternalId = Guid.NewGuid().ToString("N");
                    var tsC      = DateTime.Now.ToString("yyyyMMdd_HHmmss_fff")[..17];
                    var orderNumC = order.OrderNum ?? string.Empty;
                    var safeNumC = FileHelper.SafeFilename(orderNumC);
                    var baseNmC  = $"{tsC}_{safeNumC}_{cd.OperationType}";
                    orderResults.Add(new GenerationResult
                    {
                        ObsidianCaseId = order.ObsidianCaseId,
                        ExternalId = cd.ExternalId,
                        OrderNum  = orderNumC,
                        Amount    = cd.Amount,
                        CheckData = cd,
                        BaseName  = baseNmC,
                    });
                }

                if (orderResults.Count > 0)
                {
                    var memoTs = DateTime.Now.ToString("yyyyMMdd_HHmmss_fff")[..17];
                    var memoSafeNum = FileHelper.SafeFilename(order.OrderNum ?? string.Empty);
                    var docxPath = Path.Combine(p.OutputDir, $"{memoTs}_{memoSafeNum}_служебка.docx");
                    GenerateCorrectionMemo(order, p, order.Amount, checks.Last().OperationType, docxPath);
                    foreach (var r in orderResults)
                        r.DocxPath = docxPath;
                    results.AddRange(orderResults);
                }

                continue;   // обычная логика ниже этому заказу не нужна
            }

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
                    var serviceVatType = order.IsOwnService ? "vat122" : "none";
                    var name = $"Аванс от покупателя по заказу № {order.OrderNum}";
                    if (!string.IsNullOrEmpty(order.OrderDate))
                        name += $" от {order.OrderDate}";
                    items.Add(new CheckItem
                    {
                        Name          = name,
                        Price         = amount,
                        Quantity      = 1,
                        Sum           = amount,
                        PaymentMethod = isService ? "full_prepayment" : "advance",
                        PaymentObject = "payment",
                        VatType       = isService ? serviceVatType : "vat122",
                        VatSum        = isService ? CalcServiceVat(amount, serviceVatType) : CalcVat122(amount),
                        IsService     = isService,
                    });
                }
                else
                {
                    var sourceItems = order.Items.Count > 0
                        ? order.Items
                        : new List<OrderItem>
                        {
                            new()
                            {
                                Name = $"{(orderIsService ? "Услуга" : "Товар")} по реализации {(string.IsNullOrWhiteSpace(order.CorrectionNumber) ? order.OrderNum : order.CorrectionNumber)}",
                                Quantity = 1,
                                Sum = amount,
                            }
                        };

                    foreach (var raw in sourceItems)
                    {
                        var qty   = raw.Quantity > 0 ? raw.Quantity : 1;
                        var s     = raw.Sum;
                        var price = Math.Round(s / qty, 2);
                        bool isService = orderIsService;
                        var serviceVatType = ServiceClassificationService.ResolveVatType(
                            order.IsOwnService, agentInfo, "realization");
                        items.Add(new CheckItem
                        {
                            Name          = raw.Name,
                            Price         = price,
                            Quantity      = qty,
                            Sum           = s,
                            PaymentMethod = "full_payment",
                            PaymentObject = isService ? "service" : "commodity",
                            VatType       = isService ? serviceVatType : "vat22",
                            VatSum        = isService ? CalcServiceVat(s, serviceVatType) : CalcVat22(s),
                            IsService     = isService,
                        });
                    }
                }
            }

            // Собственная услуга использует vat122 для аванса и vat22 для реализации.
            var checkVatType = orderIsService
                             ? ServiceClassificationService.ResolveVatType(
                                 order.IsOwnService, agentInfo, p.Tab)
                             : p.Tab == "payment" ? "vat122" : "vat22";

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

                // Тег 1086: номер реализации (или основание) — будет виден в отчёте ОФД
                checkData.UserAttributeName  = "Номер реализации";
                checkData.UserAttributeValue = !string.IsNullOrEmpty(corrNum) && corrNum != "б/н"
                    ? corrNum
                    : (order.OrderNum ?? string.Empty);
            }

            // ── Имена файлов (используются для DOCX и при раздельном XML) ──
            var ts       = DateTime.Now.ToString("yyyyMMdd_HHmmss_fff")[..17];
            var orderNum = order.OrderNum ?? string.Empty;
            var safeNum  = FileHelper.SafeFilename(orderNum);
            var baseName = $"{ts}_{safeNum}_{p.CheckType}";

            var result = new GenerationResult
            {
                ObsidianCaseId = order.ObsidianCaseId,
                OrderNum  = orderNum,
                Amount    = amount,
                CheckData = checkData,
                BaseName  = baseName,
            };
            if (string.IsNullOrWhiteSpace(checkData.ExternalId))
                checkData.ExternalId = Guid.NewGuid().ToString("N");
            result.ExternalId = checkData.ExternalId;

            // ── DOCX (только для коррекций, всегда отдельный файл) ──
            if (isCorrection)
            {
                var docxPath = Path.Combine(p.OutputDir, $"{baseName}_служебка.docx");

                var orderInfoStr = orderNum;
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
            var ts = DateTime.Now.ToString("yyyyMMdd_HHmmss_fff")[..17];
            var operations = results
                .Select(r => r.CheckData?.OperationType)
                .Where(operation => !string.IsNullOrWhiteSpace(operation))
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .ToList();
            var operationLabel = operations.Count == 1 ? operations[0]! : "mixed";
            var xmlPath = Path.Combine(p.OutputDir, $"{ts}_{results.Count}чеков_{operationLabel}.xml");
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

    private static void GenerateCorrectionMemo(
        OrderEntry order, GenerationParams p, double amount, string operationType, string docxPath)
    {
        var tab = order.DocumentType == SourceDocumentType.Realization ? "realization" : p.Tab;
        var orderIsService = order.IsService || order.AgentInfo is not null || p.IsService;
        var corrDateRaw = order.CorrectionDate;

        var orderInfoStr = !string.IsNullOrWhiteSpace(order.CorrectionNumber)
            ? order.CorrectionNumber
            : order.OrderNum;
        if (!string.IsNullOrEmpty(order.OrderDate))
        {
            var docLabel = order.DocumentType == SourceDocumentType.Realization
                ? "Реализация товаров и услуг"
                : "Заказ покупателя";
            orderInfoStr += $" от {order.OrderDate} ({docLabel})";
        }

        var eventDate = !string.IsNullOrEmpty(corrDateRaw)
            ? corrDateRaw
            : !string.IsNullOrEmpty(order.OrderDate)
                ? order.OrderDate.Split(' ')[0]
                : DateTime.Today.ToString("dd.MM.yyyy");

        var operationDesc = tab == "realization"
            ? (orderIsService
                ? "реализации услуги по перевозке/сборки товара покупателю"
                : "реализации товара покупателю")
            : operationType == "buy_correction"
                ? "возврате безналичных денежных средств покупателю"
                : AppConstants.OperationDescriptions.GetValueOrDefault(operationType, "расчёте");

        var storeKkt = tab == "realization"
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
            CorrectionDesc   = AppConstants.CorrectionDescriptions.GetValueOrDefault(operationType, string.Empty),
            FromPosition     = p.Cashier.Position,
            FromNameGenitive = p.Cashier.NameGenitive,
            CashierShort     = p.Cashier.ShortName,
            KktModel         = storeKkt.Model,
            KktSerial        = storeKkt.Serial,
            KktReg           = storeKkt.RegNum,
            KktFfd           = storeKkt.Ffd,
            KktModelOnline   = onlineKkt.Model,
            KktSerialOnline  = onlineKkt.Serial,
            KktRegOnline     = onlineKkt.RegNum,
            KktFfdOnline     = onlineKkt.Ffd,
        };

        DocxGeneratorService.Generate(memo, docxPath);
    }

    private static double CalcVat122(double amount) => Math.Round(amount * 22.0 / 122.0, 2);
    // НДС 22% — сумма уже включает налог в цену (22/122)
    private static double CalcVat22(double amount)  => Math.Round(amount * 22.0 / 122.0, 2);
    private static double CalcServiceVat(double amount, string vatType) => vatType switch
    {
        "vat5" => Math.Round(amount * 5.0 / 100.0, 2),
        "vat105" => Math.Round(amount * 5.0 / 105.0, 2),
        "vat22" or "vat122" => Math.Round(amount * 22.0 / 122.0, 2),
        _ => amount,
    };

    private static void ValidateCorrectionVat(OrderEntry order)
    {
        if (!order.IsCorrection ||
            order.DocumentType != SourceDocumentType.Realization ||
            !order.IsService)
            return;

        var docNum = !string.IsNullOrWhiteSpace(order.CorrectionNumber)
            ? order.CorrectionNumber
            : order.OrderNum;

        if (string.IsNullOrWhiteSpace(order.ServiceType))
            throw new InvalidOperationException(
                $"{docNum}: не определена услуга по номенклатуре (доставка/сборка), XML коррекции не сформирован.");

        if (order.AgentInfo is null && !order.IsOwnService)
            throw new InvalidOperationException(
                $"{docNum}: не найдена ставка НДС для услуги «{order.ServiceType}» и подразделения «{order.City}», XML коррекции не сформирован.");
    }

    private static void ValidateRepairPairItems(OrderEntry order)
    {
        if (!order.IsCorrection)
            return;

        var fallbackReverseScenario = order.CorrectionScenario is CorrectionScenario.FullCancel or
            CorrectionScenario.CheckLargerAmount or
            CorrectionScenario.CheckSmallerAmount or
            CorrectionScenario.WrongPaymentType or
            CorrectionScenario.WrongNomenclature or
            CorrectionScenario.WrongDate;
        var hasReverse = !string.IsNullOrWhiteSpace(order.PlannedReverseOperation) ||
                         (fallbackReverseScenario && !string.IsNullOrWhiteSpace(order.OriginalFiscalNumber));
        if (!hasReverse)
            return;

        var docNum = !string.IsNullOrWhiteSpace(order.CorrectionNumber)
            ? order.CorrectionNumber
            : order.OrderNum;

        if (string.IsNullOrWhiteSpace(order.OriginalFiscalNumber))
            throw new InvalidOperationException(
                $"{docNum}: для исправления по ФФД 1.05 не заполнен ФП исходного чека (тег 1192).");

        var originalAmount = order.OriginalCheckAmount ?? 0;
        if (originalAmount <= 0)
            throw new InvalidOperationException(
                $"{docNum}: не заполнена сумма исходного ошибочного чека.");

        if (order.DocumentType != SourceDocumentType.Realization)
            return;

        var originalItems = order.OriginalItems.Count > 0 ? order.OriginalItems : order.Items;
        var originalItemsAmount = Math.Round(originalItems.Sum(i => i.Sum), 2);

        if (originalItems.Count == 0)
            throw new InvalidOperationException(
                $"{docNum}: не заполнена табличная часть исходного ошибочного чека.");

        if (Math.Abs(originalItemsAmount - originalAmount) > 0.01)
            throw new InvalidOperationException(
                $"{docNum}: сумма позиций исходного чека ({originalItemsAmount:N2}) не равна его итогу ({originalAmount:N2}). " +
                "Исправьте табличную часть отмены перед формированием XML.");

        var hasCorrectOrdinaryReceipt = !string.IsNullOrWhiteSpace(order.PlannedCorrectOperation) &&
                                        !order.PlannedCorrectOperation.EndsWith(
                                            "_correction", StringComparison.OrdinalIgnoreCase);
        if (!hasCorrectOrdinaryReceipt)
            return;

        var correctAmount = order.CorrectAmount ?? order.Amount;
        var correctItemsAmount = Math.Round(order.Items.Sum(i => i.Sum), 2);
        if (correctAmount <= 0 || order.Items.Count == 0)
            throw new InvalidOperationException(
                $"{docNum}: не заполнена сумма или табличная часть правильного чека.");

        if (Math.Abs(correctItemsAmount - correctAmount) > 0.01)
            throw new InvalidOperationException(
                $"{docNum}: сумма позиций правильного чека ({correctItemsAmount:N2}) не равна исправленной сумме ({correctAmount:N2}). " +
                "Исправьте правильную табличную часть перед формированием XML.");
    }

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
