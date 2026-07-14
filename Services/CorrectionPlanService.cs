using System.Globalization;
using AtolGenerator.Models;

namespace AtolGenerator.Services;

public static class CorrectionPlanService
{
    public static CorrectionPlan Build(
        OrderEntry order,
        ObsidianOriginalReceipt? originalReceipt,
        DateTime today)
    {
        if (order.CorrectionScenario == CorrectionScenario.Unknown)
            return Stop(CorrectionPlanStatus.NeedsScenario,
                "Не удалось определить сценарий по пояснению. Выберите его в редакторе.");

        var correctAmount = order.CorrectAmount ?? order.Amount;
        if (correctAmount <= 0 || !TryReadDocumentDate(order.OrderDate, out var documentDate))
            return Stop(CorrectionPlanStatus.NeedsOneC,
                "Сначала загрузите правильную сумму и дату документа из 1С.");

        var isService = order.IsService || order.AgentInfo is not null;
        if (isService && (string.IsNullOrWhiteSpace(order.ServiceType) ||
                          (order.AgentInfo is null && !order.IsOwnService)))
            return Stop(CorrectionPlanStatus.NeedsServiceRule,
                "Для услуги не определены вид услуги, агент или ставка НДС.");

        var sameDay = documentDate.Date == today.Date;
        var vatType = ResolveVatType(order, isService);
        var paymentType = ResolvePaymentType(order);
        var plan = new CorrectionPlan
        {
            Status = CorrectionPlanStatus.Ready,
            IsSameDay = sameDay,
            Message = sameDay
                ? "Документ создан сегодня: правильная сторона формируется обычным чеком."
                : "Дата документа прошла: правильная сторона формируется чеком коррекции.",
        };

        if (order.CorrectionScenario == CorrectionScenario.RealRefund)
        {
            plan.Checks.Add(Check(1, "sell_refund", "Возврат покупателю", correctAmount,
                paymentType, vatType, requiresItems: true, usesFp: false));
            return plan;
        }

        if (order.CorrectionScenario == CorrectionScenario.ExpenseCorrection)
        {
            plan.Checks.Add(Check(1, sameDay ? "sell_refund" : "buy_correction",
                sameDay ? "Возврат прихода" : "Коррекция возврата", correctAmount,
                paymentType, vatType, requiresItems: sameDay, usesFp: false));
            return plan;
        }

        if (originalReceipt is not null && IsCorrectionReceipt(originalReceipt))
            return Stop(CorrectionPlanStatus.DeferredCorrectionReceipt,
                "Исходный чек является чеком коррекции. Этот случай будет обработан на третьем этапе.");

        if (order.CorrectionScenario == CorrectionScenario.CheckNotPunched && originalReceipt is null)
        {
            if (!string.IsNullOrWhiteSpace(order.OriginalFiscalNumber))
                return Stop(CorrectionPlanStatus.NeedsOriginalReceipt,
                    "В 1С указан ФП, поэтому чек существует. Загрузите отчёт ОФД для проверки даты и операции.");
            var operation = ResolveCorrectOperation(order, sameDay);
            plan.Checks.Add(Check(1, operation, CorrectTitle(operation), correctAmount,
                paymentType, vatType, requiresItems: !IsCorrection(operation), usesFp: false));
            return plan;
        }

        if (originalReceipt is null || originalReceipt.FiscalSign is null)
            return Stop(CorrectionPlanStatus.NeedsOriginalReceipt,
                "Загрузите отчёт ОФД: исходный ошибочный чек должен быть найден строго по ФП.");

        var reverseOperation = ReverseOperation(originalReceipt.Operation, originalReceipt.Document);
        if (string.IsNullOrWhiteSpace(reverseOperation))
            return Stop(CorrectionPlanStatus.NeedsOriginalReceipt,
                "В отчёте ОФД не удалось определить операцию исходного чека.");

        var wrongAmount = Math.Abs(originalReceipt.Amount);
        if (wrongAmount <= 0)
            return Stop(CorrectionPlanStatus.NeedsOriginalReceipt,
                "В исходном чеке ОФД не заполнена сумма.");

        plan.Checks.Add(Check(1, reverseOperation, ReverseTitle(reverseOperation), wrongAmount,
            ResolveOriginalPaymentType(order, paymentType), vatType,
            requiresItems: true, usesFp: true));

        if (order.CorrectionScenario == CorrectionScenario.CheckNotPunched)
            plan.Message = sameDay
                ? "В 1С и ОФД найден исходный чек: вместо одиночного чека подготовлена отмена и правильный обычный чек."
                : "В 1С и ОФД найден исходный чек: вместо одиночной коррекции подготовлен исправительный комплект.";

        if (order.CorrectionScenario != CorrectionScenario.FullCancel)
        {
            var correctOperation = ResolveCorrectOperation(order, sameDay);
            plan.Checks.Add(Check(2, correctOperation, CorrectTitle(correctOperation), correctAmount,
                ResolveCorrectPaymentType(order, paymentType), vatType,
                requiresItems: !IsCorrection(correctOperation), usesFp: IsCorrection(correctOperation)));
        }

        return plan;
    }

    public static bool IsCorrectionReceipt(ObsidianOriginalReceipt receipt)
    {
        var text = $"{receipt.Document} {receipt.Operation}";
        return text.Contains("коррек", StringComparison.OrdinalIgnoreCase);
    }

    private static CorrectionPlan Stop(CorrectionPlanStatus status, string message) => new()
    {
        Status = status,
        Message = message,
    };

    private static CorrectionPlanCheck Check(
        int sequence,
        string operation,
        string title,
        double amount,
        string paymentType,
        string vatType,
        bool requiresItems,
        bool usesFp) => new()
    {
        Sequence = sequence,
        Operation = operation,
        Title = title,
        Amount = amount,
        PaymentType = paymentType,
        VatType = vatType,
        RequiresItems = requiresItems,
        UsesOriginalFiscalSign = usesFp,
        XmlOnly = IsCorrection(operation),
    };

    private static string ResolveCorrectOperation(OrderEntry order, bool sameDay)
    {
        var isReturn = order.DocumentType == SourceDocumentType.CashExpense;
        if (sameDay) return isReturn ? "sell_refund" : "sell";
        return isReturn ? "buy_correction" : "sell_correction";
    }

    private static string ReverseOperation(string operation, string document)
    {
        var value = $"{operation} {document}".ToLowerInvariant();
        if (value.Contains("возврат прихода") || value.Contains("sell_refund")) return "sell";
        if (value.Contains("возврат расхода") || value.Contains("buy_refund")) return "buy";
        if (value.Contains("приход") || value.Contains("sell")) return "sell_refund";
        if (value.Contains("расход") || value.Contains("buy")) return "buy_refund";
        return string.Empty;
    }

    private static string ResolveVatType(OrderEntry order, bool isService)
    {
        var tab = order.DocumentType == SourceDocumentType.Realization ? "realization" : "payment";
        if (!isService) return tab == "payment" ? "vat122" : "vat22";
        return ServiceClassificationService.ResolveVatType(order.IsOwnService, order.AgentInfo, tab);
    }

    private static string ResolvePaymentType(OrderEntry order) => order.DocumentType switch
    {
        SourceDocumentType.CashPayment or SourceDocumentType.CashExpense => "cash",
        SourceDocumentType.CardPayment => "card",
        _ => "card",
    };

    private static string ResolveOriginalPaymentType(OrderEntry order, string fallback) =>
        order.OriginalPaymentWasCash.HasValue
            ? order.OriginalPaymentWasCash.Value ? "cash" : "card"
            : fallback;

    private static string ResolveCorrectPaymentType(OrderEntry order, string fallback) =>
        order.CorrectPaymentIsCash.HasValue
            ? order.CorrectPaymentIsCash.Value ? "cash" : "card"
            : fallback;

    private static bool IsCorrection(string operation) =>
        operation.EndsWith("_correction", StringComparison.OrdinalIgnoreCase);

    private static string CorrectTitle(string operation) => operation switch
    {
        "sell" => "Правильный приход",
        "sell_refund" => "Правильный возврат прихода",
        "sell_correction" => "Коррекция прихода",
        "buy_correction" => "Коррекция возврата",
        _ => "Правильный чек",
    };

    private static string ReverseTitle(string operation) => operation switch
    {
        "sell" => "Отмена возврата прихода",
        "sell_refund" => "Возврат ошибочного прихода",
        "buy" => "Отмена возврата расхода",
        "buy_refund" => "Возврат ошибочного расхода",
        _ => "Отмена ошибочного чека",
    };

    private static bool TryReadDocumentDate(string value, out DateTime date)
    {
        var formats = new[] { "dd.MM.yyyy HH:mm:ss", "dd.MM.yyyy H:mm:ss", "dd.MM.yyyy" };
        return DateTime.TryParseExact(value?.Trim(), formats, CultureInfo.InvariantCulture,
            DateTimeStyles.AllowWhiteSpaces, out date);
    }
}
