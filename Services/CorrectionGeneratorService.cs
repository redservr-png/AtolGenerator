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
        // A full cancellation is always the reverse side of the original receipt.
        // Do not let a stale or manually edited plan turn it into a sell operation.
        if (order.CorrectionScenario == CorrectionScenario.FullCancel)
            return BuildFullCancel(order, p);

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
            CorrectionScenario.RealRefund         => BuildRealRefund(order, p),
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
                paymentIsCashOverride: o.OriginalPaymentWasCash));
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
                    includeOriginalFp: false,
                    paymentIsCashOverride: o.CorrectPaymentIsCash));
            }
        }

        return result;
    }

    // ── СЦЕНАРИИ ────────────────────────────────────────────────────────────────

    /// <summary>1) Полная отмена «лишнего» чека — один sell_refund на полную сумму.</summary>
    private static List<CheckData> BuildFullCancel(OrderEntry o, GenerationParams p)
    {
        var refund = MakeRefundCheckData(o, p,
            amount: o.OriginalCheckAmount ?? o.Amount,
            includeOriginalFp: true);
        return new List<CheckData> { refund };
    }

    /// <summary>
    /// 2) Чек большей суммой / меньшей суммой → пара чеков.
    /// Логика одинаковая для обоих: чек УЖЕ пробит с ошибочной суммой, поэтому:
    ///   sell_refund(ошибочная_сумма_из_чека) + sell_correction(правильная_сумма_из_1С).
    /// </summary>
    private static List<CheckData> BuildRefundPlusCorrection(OrderEntry o, GenerationParams p)
    {
        var wrong   = o.OriginalCheckAmount ?? o.Amount;
        var correct = o.CorrectAmount       ?? o.Amount;

        var refund = MakeRefundCheckData    (o, p, amount: wrong,   includeOriginalFp: true);
        var corr   = MakeCorrectionCheckData(o, p, amount: correct, isExpense: false);
        return new List<CheckData> { refund, corr };
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

    /// <summary>4) Перепутали способ оплаты → sell_refund старого + sell_correction правильного.</summary>
    private static List<CheckData> BuildWrongPaymentType(OrderEntry o, GenerationParams p)
    {
        var sum = o.Amount;
        // refund — со старым (ошибочным) типом оплаты
        var refund = MakeRefundCheckData(o, p, amount: sum, includeOriginalFp: true,
            paymentIsCashOverride: o.OriginalPaymentWasCash);
        // correction — с правильным типом
        var corr = MakeCorrectionCheckData(o, p, amount: sum, isExpense: false,
            paymentIsCashOverride: o.CorrectPaymentIsCash);
        return new List<CheckData> { refund, corr };
    }

    /// <summary>5) Перепутали номенклатуру → sell_refund старого + sell_correction правильного.</summary>
    private static List<CheckData> BuildWrongNomenclature(OrderEntry o, GenerationParams p)
    {
        var sum = o.Amount;
        var refund = MakeRefundCheckData(o, p, amount: sum, includeOriginalFp: true);
        var corr   = MakeCorrectionCheckData(o, p, amount: sum, isExpense: false);
        return new List<CheckData> { refund, corr };
    }

    /// <summary>6) Реальный возврат денег покупателю — самостоятельный sell_refund.</summary>
    private static List<CheckData> BuildRealRefund(OrderEntry o, GenerationParams p)
    {
        // Тег 1192 не нужен — это самостоятельная операция, не исправление.
        var refund = MakeRefundCheckData(o, p, amount: o.Amount, includeOriginalFp: false);
        return new List<CheckData> { refund };
    }

    /// <summary>7) Чек пробит другой датой → sell_correction с правильной base_date.</summary>
    private static List<CheckData> BuildWrongDate(OrderEntry o, GenerationParams p)
    {
        if (o.DocumentType == SourceDocumentType.Realization &&
            !string.IsNullOrWhiteSpace(o.OriginalFiscalNumber))
        {
            var amount = o.CorrectAmount ?? o.Amount;
            var refund = MakeRefundCheckData(o, p, amount: o.OriginalCheckAmount ?? amount, includeOriginalFp: true);
            var correction = MakeCorrectionCheckData(o, p, amount: amount, isExpense: false);
            return new List<CheckData> { refund, correction };
        }

        var corr = MakeCorrectionCheckData(o, p, amount: o.Amount, isExpense: false);
        return new List<CheckData> { corr };
    }

    /// <summary>8) Чек коррекции «Расход» → buy_correction.</summary>
    private static List<CheckData> BuildExpenseCorrection(OrderEntry o, GenerationParams p)
    {
        var corr = MakeCorrectionCheckData(o, p, amount: o.Amount, isExpense: true);
        return new List<CheckData> { corr };
    }

    // ── HELPERS ─────────────────────────────────────────────────────────────────

    /// <summary>Создаёт CheckData для возвратного чека (sell_refund).</summary>
    private static CheckData MakeRefundCheckData(
        OrderEntry o, GenerationParams p, double amount,
        bool includeOriginalFp,
        bool? paymentIsCashOverride = null)
        => MakeReceiptCheckData(o, p, "sell_refund", amount, includeOriginalFp, paymentIsCashOverride);

    private static CheckData MakeReceiptCheckData(
        OrderEntry o,
        GenerationParams p,
        string operation,
        double amount,
        bool includeOriginalFp,
        bool? paymentIsCashOverride)
    {
        var paymentType = ResolvePaymentType(o, p, paymentIsCashOverride);
        bool isService = ResolveIsService(o);
        string tab = ResolveTab(o);
        var operationKind = operation.EndsWith("_refund", StringComparison.OrdinalIgnoreCase)
            ? "refund"
            : "sell";
        var items = BuildItems(o, amount, isService, tab, operationKind);

        return new CheckData
        {
            OperationType        = operation,
            IsCorrection         = false,
            Tab                  = tab,
            Amount               = amount,
            PaymentType          = paymentType,
            CheckVatType         = ResolveCheckVatType(o, tab, isService),
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
            CheckVatType         = ResolveCheckVatType(o, tab, isService),
            Items                = new List<CheckItem>(),   // коррекция без табличной части
            Agent                = o.AgentInfo,
            IsService            = isService,
            CorrectionBaseDate   = baseDate,
            CorrectionBaseNumber = baseNumber,
            CashierName          = p.Cashier.FullName,
            CashierShort         = p.Cashier.ShortName,
        };
    }

    /// <summary>Создаёт CheckData для нового «правильного» sell-чека (используется в DecreaseAmount).</summary>
    private static CheckData MakeSellCheckData(OrderEntry o, GenerationParams p, double amount)
        => MakeReceiptCheckData(o, p, "sell", amount, includeOriginalFp: false, paymentIsCashOverride: null);

    // ── Common builders ────────────────────────────────────────────────────────

    private static List<CheckItem> BuildItems(OrderEntry o, double amount, bool isService, string tab, string operationKind)
    {
        if (tab == "realization" && o.Items.Count > 0)
            return BuildRealizationItems(o, isService);

        return BuildOneItem(o, amount, isService, tab, operationKind);
    }

    private static List<CheckItem> BuildRealizationItems(OrderEntry o, bool isService)
    {
        var result = new List<CheckItem>();
        var vatType = ResolveCheckVatType(o, "realization", isService);

        foreach (var raw in o.Items.Where(i => !string.IsNullOrWhiteSpace(i.Name) && i.Sum > 0))
        {
            var qty = raw.Quantity > 0 ? raw.Quantity : 1;
            var sum = raw.Sum;
            result.Add(new CheckItem
            {
                Name          = raw.Name.Trim(),
                Price         = Math.Round(sum / qty, 2),
                Quantity      = qty,
                Sum           = sum,
                PaymentMethod = "full_payment",
                PaymentObject = isService ? "service" : "commodity",
                VatType       = vatType,
                VatSum        = CalcVat(sum, vatType),
                IsService     = isService,
            });
        }

        return result.Count > 0
            ? result
            : BuildOneItem(o, o.OriginalCheckAmount ?? o.CorrectAmount ?? o.Amount, isService, "realization", "refund");
    }

    private static List<CheckItem> BuildOneItem(OrderEntry o, double amount, bool isService, string tab, string operationKind)
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

        string vatType = ResolveCheckVatType(o, tab, isService);

        double vatSum = vatType switch
        {
            "vat5"   => Math.Round(amount * 5.0  / 100.0, 2),
            "vat105" => Math.Round(amount * 5.0  / 105.0, 2),
            "vat22"  => Math.Round(amount * 22.0 / 122.0, 2),
            "vat122" => Math.Round(amount * 22.0 / 122.0, 2),
            _        => amount,
        };

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

    private static double CalcVat(double amount, string vatType) => vatType switch
    {
        "vat5"   => Math.Round(amount * 5.0  / 100.0, 2),
        "vat105" => Math.Round(amount * 5.0  / 105.0, 2),
        "vat22"  => Math.Round(amount * 22.0 / 122.0, 2),
        "vat122" => Math.Round(amount * 22.0 / 122.0, 2),
        _        => amount,
    };

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

    private static string ResolveCheckVatType(OrderEntry o, string tab, bool isService)
    {
        if (!string.IsNullOrWhiteSpace(o.PlannedVatType)) return o.PlannedVatType;
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
