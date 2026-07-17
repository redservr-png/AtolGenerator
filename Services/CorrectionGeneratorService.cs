using AtolGenerator.Constants;
using AtolGenerator.Models;

namespace AtolGenerator.Services;

/// <summary>
/// Строит набор CheckData (1 или 2 чека) для одной OrderEntry-коррекции.
/// Выбор операций и параметров зависит от order.CorrectionScenario.
///
/// Возвращает CheckData в порядке: возвратная сторона → коррекционная сторона
/// (если есть). Каждый CheckData готов к сериализации XmlGeneratorService.
/// </summary>
public static class CorrectionGeneratorService
{
    public static List<CheckData> BuildCheckDataList(OrderEntry order, GenerationParams p)
    {
        if (!string.IsNullOrWhiteSpace(order.PlannedReverseOperation) ||
            !string.IsNullOrWhiteSpace(order.PlannedCorrectOperation))
            return BuildPlanned(order, p);

        return order.CorrectionScenario switch
        {
            CorrectionScenario.FullCancel         => BuildFullCancel(order, p),
            CorrectionScenario.CheckLargerAmount  => BuildRefundPlusCorrection(order, p),
            CorrectionScenario.CheckSmallerAmount => BuildRefundPlusCorrection(order, p),
            CorrectionScenario.CheckNotPunched    => BuildCheckNotPunched(order, p),
            CorrectionScenario.WrongPaymentType   => BuildWrongPaymentType(order, p),
            CorrectionScenario.WrongNomenclature  => BuildWrongNomenclature(order, p),
            CorrectionScenario.RealRefund         => throw new InvalidOperationException(
                "Реальный возврат оформляется в разделе «Возвраты по заказам», а не в исправлениях Obsidian."),
            CorrectionScenario.WrongDate          => BuildWrongDate(order, p),
            CorrectionScenario.ExpenseCorrection  => BuildExpenseCorrection(order, p),
            // Unknown — генерация пропускается, пользователь должен выбрать сценарий
            _                                     => new List<CheckData>(),
        };
    }

    private static List<CheckData> BuildPlanned(OrderEntry o, GenerationParams p)
    {
        var result = new List<CheckData>();
        if (!string.IsNullOrWhiteSpace(o.PlannedReverseOperation))
        {
            result.Add(MakeReceiptCheckData(
                o,
                p,
                o.PlannedReverseOperation,
                o.OriginalCheckAmount ?? o.Amount,
                includeOriginalFp: true,
                paymentIsCashOverride: o.OriginalPaymentWasCash,
                useOriginalItems: true));
        }

        if (!string.IsNullOrWhiteSpace(o.PlannedCorrectOperation))
        {
            var amount = o.CorrectAmount ?? o.Amount;
            if (o.PlannedCorrectOperation.EndsWith("_correction", StringComparison.OrdinalIgnoreCase))
            {
                result.Add(MakeCorrectionCheckData(
                    o,
                    p,
                    amount,
                    isExpense: o.PlannedCorrectOperation == "buy_correction",
                    paymentIsCashOverride: o.CorrectPaymentIsCash));
            }
            else
            {
                result.Add(MakeReceiptCheckData(
                    o,
                    p,
                    o.PlannedCorrectOperation,
                    amount,
                    includeOriginalFp: !string.IsNullOrWhiteSpace(o.PlannedReverseOperation),
                    paymentIsCashOverride: o.CorrectPaymentIsCash,
                    useOriginalItems: false));
            }
        }

        return result;
    }

    // ── СЦЕНАРИИ ────────────────────────────────────────────────────────────────

    /// <summary>1) Полная отмена «лишнего» чека — один обычный обратный чек с тегом 1192.</summary>
    private static List<CheckData> BuildFullCancel(OrderEntry o, GenerationParams p)
    {
        var originalOperation = ResolveOriginalOperation(o);
        var reverseOperation = ReverseOperation(originalOperation);
        var reverse = MakeReceiptCheckData(o, p, reverseOperation,
            amount: o.OriginalCheckAmount ?? o.Amount,
            includeOriginalFp: true,
            paymentIsCashOverride: o.OriginalPaymentWasCash,
            useOriginalItems: true);
        return new List<CheckData> { reverse };
    }

    /// <summary>
    /// 2) Чек большей суммой / меньшей суммой → пара чеков.
    /// Логика одинаковая для обоих: чек УЖЕ пробит с ошибочной суммой, поэтому:
    ///   обычный обратный чек + правильный обычный чек; оба с тегом 1192.
    /// </summary>
    private static List<CheckData> BuildRefundPlusCorrection(OrderEntry o, GenerationParams p)
    {
        var wrong   = o.OriginalCheckAmount ?? o.Amount;
        var correct = o.CorrectAmount       ?? o.Amount;

        var originalOperation = ResolveOriginalOperation(o);
        var reverse = MakeReceiptCheckData(o, p, ReverseOperation(originalOperation), wrong,
            includeOriginalFp: true, paymentIsCashOverride: o.OriginalPaymentWasCash,
            useOriginalItems: true);
        var corrected = MakeReceiptCheckData(o, p, originalOperation, correct,
            includeOriginalFp: true, paymentIsCashOverride: o.CorrectPaymentIsCash,
            useOriginalItems: false);
        return new List<CheckData> { reverse, corrected };
    }

    /// <summary>
    /// 3) Чек не пробит — был расчёт, но фискального чека нет.
    /// → Один sell_correction на полную сумму (refund не нужен — чека и так нет).
    /// </summary>
    private static List<CheckData> BuildCheckNotPunched(OrderEntry o, GenerationParams p)
    {
        var amount = o.CorrectAmount ?? o.Amount;
        var corr   = MakeCorrectionCheckData(o, p, amount: amount, isExpense: false);
        return new List<CheckData> { corr };
    }

    /// <summary>4) Перепутали способ оплаты → два обычных чека с тегом 1192.</summary>
    private static List<CheckData> BuildWrongPaymentType(OrderEntry o, GenerationParams p)
    {
        return BuildOfficialRepairPair(o, p,
            originalPaymentIsCash: o.OriginalPaymentWasCash,
            correctPaymentIsCash: o.CorrectPaymentIsCash);
    }

    /// <summary>5) Перепутали номенклатуру → два обычных чека с разными табличными частями.</summary>
    private static List<CheckData> BuildWrongNomenclature(OrderEntry o, GenerationParams p)
    {
        return BuildOfficialRepairPair(o, p);
    }

    /// <summary>7) Чек пробит другой датой → sell_correction с правильной base_date.</summary>
    private static List<CheckData> BuildWrongDate(OrderEntry o, GenerationParams p)
    {
        if (o.DocumentType == SourceDocumentType.Realization &&
            !string.IsNullOrWhiteSpace(o.OriginalFiscalNumber))
        {
            return BuildOfficialRepairPair(o, p);
        }

        var corr = MakeCorrectionCheckData(o, p, amount: o.Amount, isExpense: false);
        return new List<CheckData> { corr };
    }

    /// <summary>8) Чек коррекции «Расход» → buy_correction.</summary>
    private static List<CheckData> BuildExpenseCorrection(OrderEntry o, GenerationParams p)
    {
        if (!string.IsNullOrWhiteSpace(o.OriginalFiscalNumber))
            return BuildOfficialRepairPair(o, p);

        var corr = MakeCorrectionCheckData(o, p, amount: o.Amount, isExpense: true);
        return new List<CheckData> { corr };
    }

    private static List<CheckData> BuildOfficialRepairPair(
        OrderEntry o,
        GenerationParams p,
        bool? originalPaymentIsCash = null,
        bool? correctPaymentIsCash = null)
    {
        var originalOperation = ResolveOriginalOperation(o);
        var wrongAmount = o.OriginalCheckAmount ?? o.Amount;
        var correctAmount = o.CorrectAmount ?? o.Amount;
        var reverse = MakeReceiptCheckData(o, p, ReverseOperation(originalOperation), wrongAmount,
            includeOriginalFp: true,
            paymentIsCashOverride: originalPaymentIsCash ?? o.OriginalPaymentWasCash,
            useOriginalItems: true);
        var corrected = MakeReceiptCheckData(o, p, originalOperation, correctAmount,
            includeOriginalFp: true,
            paymentIsCashOverride: correctPaymentIsCash ?? o.CorrectPaymentIsCash,
            useOriginalItems: false);
        return new List<CheckData> { reverse, corrected };
    }

    // ── HELPERS ─────────────────────────────────────────────────────────────────

    /// <summary>Создаёт CheckData для возвратного чека (sell_refund).</summary>
    private static CheckData MakeRefundCheckData(
        OrderEntry o, GenerationParams p, double amount,
        bool includeOriginalFp,
        bool? paymentIsCashOverride = null)
        => MakeReceiptCheckData(o, p, "sell_refund", amount, includeOriginalFp,
            paymentIsCashOverride, useOriginalItems: includeOriginalFp);

    private static CheckData MakeReceiptCheckData(
        OrderEntry o,
        GenerationParams p,
        string operation,
        double amount,
        bool includeOriginalFp,
        bool? paymentIsCashOverride,
        bool useOriginalItems = false)
    {
        var paymentType = ResolvePaymentType(o, p, paymentIsCashOverride);
        bool isService = ResolveIsService(o);
        string tab = ResolveTab(o);
        var vatType = ResolveReceiptVatType(o, tab, isService, useOriginalItems);
        var items = BuildItems(o, amount, isService, tab, useOriginalItems, vatType);

        return new CheckData
        {
            OperationType        = operation,
            IsCorrection         = false,
            Tab                  = tab,
            Amount               = amount,
            PaymentType          = paymentType,
            CheckVatType         = vatType,
            Items                = items,
            Agent                = o.AgentInfo,
            IsService            = isService,
            CashierName          = p.Cashier.FullName,
            CashierShort         = p.Cashier.ShortName,
            AdditionalCheckProps = includeOriginalFp ? (o.OriginalFiscalNumber ?? string.Empty) : string.Empty,
            UserAttributeName    = "Номер реализации",
            UserAttributeValue   = !string.IsNullOrEmpty(o.CorrectionNumber)
                ? o.CorrectionNumber
                : (!string.IsNullOrEmpty(o.OrderNum) ? o.OrderNum : string.Empty),
        };
    }

    /// <summary>Создаёт CheckData для коррекционного чека (sell_correction или buy_correction).</summary>
    private static CheckData MakeCorrectionCheckData(
        OrderEntry o, GenerationParams p, double amount,
        bool isExpense,
        bool? paymentIsCashOverride = null)
    {
        var paymentType = ResolvePaymentType(o, p, paymentIsCashOverride);
        bool isService  = ResolveIsService(o);
        string tab      = ResolveTab(o);
        string vatType  = ResolveCorrectVatType(o, tab, isService);

        var baseNumber = !string.IsNullOrEmpty(o.CorrectionNumber)
            ? o.CorrectionNumber
            : (!string.IsNullOrEmpty(o.OrderNum) ? o.OrderNum : "б/н");
        var rawBaseDate = !string.IsNullOrEmpty(o.CorrectionDate)
            ? o.CorrectionDate
            : (!string.IsNullOrEmpty(o.OrderDate)
                ? o.OrderDate.Split(' ').FirstOrDefault() ?? string.Empty
                : string.Empty);
        var baseDate = !string.IsNullOrEmpty(rawBaseDate)
            ? ToIsoDateOrToday(rawBaseDate)
            : DateTime.Today.ToString("yyyy-MM-dd");

        return new CheckData
        {
            OperationType        = isExpense ? "buy_correction" : "sell_correction",
            IsCorrection         = true,
            Tab                  = tab,
            Amount               = amount,
            PaymentType          = paymentType,
            CheckVatType         = vatType,
            Items                = new List<CheckItem>(),   // коррекция без табличной части
            Agent                = o.AgentInfo,
            IsService            = isService,
            CorrectionBaseDate   = baseDate,
            CorrectionBaseNumber = baseNumber,
            CashierName          = p.Cashier.FullName,
            CashierShort         = p.Cashier.ShortName,
        };
    }

    // ── Common builders ────────────────────────────────────────────────────────

    private static List<CheckItem> BuildItems(
        OrderEntry o,
        double amount,
        bool isService,
        string tab,
        bool useOriginalItems,
        string fallbackVatType)
    {
        var sourceItems = useOriginalItems && o.OriginalItems.Count > 0
            ? o.OriginalItems
            : o.Items;
        if (sourceItems.Count > 0)
            return BuildSourceItems(o, sourceItems, amount, isService, tab, fallbackVatType);

        return BuildOneItem(o, amount, isService, tab, fallbackVatType);
    }

    private static List<CheckItem> BuildSourceItems(
        OrderEntry o,
        IReadOnlyList<OrderItem> sourceItems,
        double amount,
        bool isService,
        string tab,
        string fallbackVatType)
    {
        var result = new List<CheckItem>();
        var paymentMethod = tab == "payment"
            ? isService ? "full_prepayment" : "advance"
            : "full_payment";
        var paymentObject = tab == "payment"
            ? "payment"
            : isService ? "service" : "commodity";

        foreach (var raw in sourceItems.Where(i => !string.IsNullOrWhiteSpace(i.Name) && i.Sum > 0))
        {
            var qty = raw.Quantity > 0 ? raw.Quantity : 1;
            var sum = raw.Sum;
            var vatType = VatRateCatalog.Normalize(raw.VatType, fallbackVatType);
            result.Add(new CheckItem
            {
                Name          = raw.Name.Trim(),
                Price         = Math.Round(sum / qty, 2),
                Quantity      = qty,
                Sum           = sum,
                PaymentMethod = paymentMethod,
                PaymentObject = paymentObject,
                VatType       = vatType,
                VatSum        = CalcVat(sum, vatType),
                IsService     = isService,
            });
        }

        return result.Count > 0
            ? result
            : BuildOneItem(o, amount, isService, tab, fallbackVatType);
    }

    private static List<CheckItem> BuildOneItem(
        OrderEntry o,
        double amount,
        bool isService,
        string tab,
        string fallbackVatType)
    {
        // Для оплаты — «Аванс от покупателя по заказу № X», для реализации — «Услуга/Товар по заказу X»
        string name = tab == "payment"
            ? $"Аванс от покупателя по заказу № {o.OrderNum}"
            : (isService ? $"Услуга по заказу {o.OrderNum}" : $"Товар по заказу {o.OrderNum}");

        string method = tab == "payment"
            ? (isService ? "full_prepayment" : "advance")
            : "full_payment";

        string obj = tab == "payment"
            ? "payment"
            : (isService ? "service" : "commodity");

        string vatType = VatRateCatalog.Normalize(fallbackVatType, "none");
        double vatSum = VatRateCatalog.CalculateFiscalSum(amount, vatType);

        return new List<CheckItem>
        {
            new()
            {
                Name          = name,
                Price         = amount,
                Quantity      = 1,
                Sum           = amount,
                PaymentMethod = method,
                PaymentObject = obj,
                VatType       = vatType,
                VatSum        = vatSum,
                IsService     = isService,
            }
        };
    }

    private static double CalcVat(double amount, string vatType) =>
        VatRateCatalog.CalculateFiscalSum(amount, vatType);

    private static string ResolvePaymentType(OrderEntry o, GenerationParams p, bool? cashOverride)
    {
        if (cashOverride.HasValue) return cashOverride.Value ? "cash" : "card";
        // Дефолт по типу документа: ПКО/РКО = cash; CardPayment = card; иначе из p.PaymentType
        return o.DocumentType switch
        {
            SourceDocumentType.CashPayment => "cash",
            SourceDocumentType.CashExpense => "cash",
            SourceDocumentType.CardPayment => "card",
            _ => p.PaymentType,
        };
    }

    private static bool ResolveIsService(OrderEntry o)
        => o.IsService || o.AgentInfo is not null;

    private static string ResolveTab(OrderEntry o)
        => o.DocumentType == SourceDocumentType.Realization ? "realization" : "payment";

    private static string ResolveOriginalOperation(OrderEntry o)
    {
        var value = o.OriginalCheckOperation?.Trim().ToLowerInvariant() ?? string.Empty;
        if (value.Contains("возврат прихода") || value.Contains("sell_refund")) return "sell_refund";
        if (value.Contains("возврат расхода") || value.Contains("buy_refund")) return "buy_refund";
        if (value.Contains("приход") || value == "sell") return "sell";
        if (value.Contains("расход") || value == "buy") return "buy";

        return o.DocumentType == SourceDocumentType.CashExpense ? "sell_refund" : "sell";
    }

    private static string ReverseOperation(string operation) => operation switch
    {
        "sell" => "sell_refund",
        "sell_refund" => "sell",
        "buy" => "buy_refund",
        "buy_refund" => "buy",
        _ => "sell_refund",
    };

    private static string ResolveReceiptVatType(
        OrderEntry o,
        string tab,
        bool isService,
        bool useOriginalItems) => useOriginalItems
        ? ResolveOriginalVatType(o, tab, isService)
        : ResolveCorrectVatType(o, tab, isService);

    private static string ResolveOriginalVatType(OrderEntry o, string tab, bool isService) =>
        VatRateCatalog.Normalize(o.OriginalVatType, ResolveCorrectVatType(o, tab, isService));

    private static string ResolveCorrectVatType(OrderEntry o, string tab, bool isService)
    {
        if (VatRateCatalog.IsKnown(o.CorrectVatType)) return VatRateCatalog.Normalize(o.CorrectVatType);
        if (VatRateCatalog.IsKnown(o.PlannedVatType)) return VatRateCatalog.Normalize(o.PlannedVatType);
        if (isService)
            return ServiceClassificationService.ResolveVatType(o.IsOwnService, o.AgentInfo, tab);
        return tab == "payment" ? "vat122" : "vat22";
    }

    private static string ToIsoDateOrToday(string ddMmYyyy)
    {
        if (DateTime.TryParseExact(ddMmYyyy.Trim(), "dd.MM.yyyy",
                System.Globalization.CultureInfo.InvariantCulture,
                System.Globalization.DateTimeStyles.None, out var d))
            return d.ToString("yyyy-MM-dd");
        return DateTime.Today.ToString("yyyy-MM-dd");
    }
}
