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
    {
        var paymentType = ResolvePaymentType(o, p, paymentIsCashOverride);
        bool isService  = ResolveIsService(o);
        string tab      = ResolveTab(o);
        var items       = BuildOneItem(o, amount, isService, tab, operationKind: "refund");

        return new CheckData
        {
            OperationType        = "sell_refund",
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
            UserAttributeValue   = !string.IsNullOrEmpty(o.OrderNum) ? o.OrderNum : string.Empty,
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

        var baseNumber = !string.IsNullOrEmpty(o.OrderNum) ? o.OrderNum : "б/н";
        var baseDate   = !string.IsNullOrEmpty(o.OrderDate)
            ? ToIsoDateOrToday(o.OrderDate.Split(' ').FirstOrDefault() ?? string.Empty)
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
    {
        var paymentType = ResolvePaymentType(o, p, null);
        bool isService  = ResolveIsService(o);
        string tab      = ResolveTab(o);
        var items       = BuildOneItem(o, amount, isService, tab, operationKind: "sell");

        return new CheckData
        {
            OperationType        = "sell",
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
            UserAttributeName    = "Номер реализации",
            UserAttributeValue   = !string.IsNullOrEmpty(o.OrderNum) ? o.OrderNum : string.Empty,
        };
    }

    // ── Common builders ────────────────────────────────────────────────────────

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

        string vatType = isService
            ? (o.AgentInfo?.VatType ?? "none")
            : (tab == "payment" ? "vat122" : "vat22");

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
        if (isService) return o.AgentInfo?.VatType ?? "none";
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
