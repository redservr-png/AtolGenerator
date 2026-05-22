namespace AtolGenerator.Models;

/// <summary>
/// Сценарии исправительных чеков (основаны на типах ошибок из «Исправить чеки.md»).
/// Определяют какую пару/один чек надо пробить и с какими параметрами.
/// </summary>
public enum CorrectionScenario
{
    /// <summary>Не определён — пользователь должен выбрать вручную</summary>
    Unknown = 0,

    /// <summary>
    /// Лишний чек — отменить полностью.
    /// «Не было», «Не должно», «Лишний», «Деньги не списались», «Оплаты не было»,
    /// «Удалить», «Покупатель не платил», «помечен на удаление чек пробит».
    /// → Один sell_refund на полную сумму.
    /// </summary>
    FullCancel,

    /// <summary>
    /// Чек большей суммой — уменьшить базу на разницу.
    /// «Чек большей суммой».
    /// → sell_refund (отмена старого) + sell (правильная сумма).
    /// </summary>
    DecreaseAmount,

    /// <summary>
    /// Чек меньшей суммой / не пробит — увеличить базу.
    /// «Чек меньшей суммой», «не пробит», «не вышел», «не в день оплаты».
    /// → sell_correction на разницу/полную сумму.
    /// </summary>
    IncreaseAmount,

    /// <summary>
    /// Неправильный способ оплаты — отменить и пробить заново.
    /// «Чек пробит наличными» (а было картой), «перепутали оплаты», «нал и безнал».
    /// → sell_refund + sell_correction с правильным payments.type.
    /// </summary>
    WrongPaymentType,

    /// <summary>
    /// Неправильная номенклатура — заменить.
    /// «Перепутали номенклатуру», «ошибочно провела сборку», «другая номенклатура».
    /// → sell_refund + sell_correction с правильной номенклатурой.
    /// </summary>
    WrongNomenclature,

    /// <summary>
    /// Реальный возврат денег (фактический).
    /// «Возврат по бухгалтерии», «деньги вернулись клиенту».
    /// → Один sell_refund (самостоятельная операция, не исправление).
    /// </summary>
    RealRefund,

    /// <summary>
    /// Чек пробит другой датой — корректировка.
    /// «Чек другой датой», «дата следующая».
    /// → sell_correction с правильной base_date.
    /// </summary>
    WrongDate,

    /// <summary>
    /// Чек коррекции «Расход» (для расходных операций по агентам).
    /// «Чек коррекции Расход», РКО с ошибкой.
    /// → buy_correction.
    /// </summary>
    ExpenseCorrection,
}

/// <summary>
/// Вид заказа в основном списке: обычный или исправительный.
/// </summary>
public enum OrderKind
{
    /// <summary>Обычный заказ (Оплата/Реализация — без исправлений)</summary>
    Regular = 0,

    /// <summary>Только sell_refund — отмена ошибочного чека</summary>
    SingleRefund,

    /// <summary>Только sell_correction — доплата/корректировка</summary>
    SingleCorrection,

    /// <summary>Пара sell_refund + sell_correction (или sell_refund + sell)</summary>
    RefundCorrectionPair,
}

/// <summary>
/// Тип исходного документа из 1С (для парсинга Obsidian).
/// </summary>
public enum SourceDocumentType
{
    Unknown = 0,
    Realization,            // Реализация товаров и услуг
    CardPayment,            // Оплата от покупателя платежной картой
    CashPayment,            // Приходный кассовый ордер (ПКО)
    CashExpense,            // Расходный кассовый ордер (РКО)
    BuyerOrder,             // Заказ покупателя
    KkmCheck,               // Чек ККМ
    FpOnly,                 // Только ФП (без документа)
}

public static class CorrectionScenarioExtensions
{
    /// <summary>Удобное русское название для combobox в UI.</summary>
    public static string ToDisplayString(this CorrectionScenario s) => s switch
    {
        CorrectionScenario.Unknown            => "— Не определён —",
        CorrectionScenario.FullCancel         => "Полная отмена (лишний чек)",
        CorrectionScenario.DecreaseAmount     => "Чек большей суммой",
        CorrectionScenario.IncreaseAmount     => "Чек меньшей суммой / не пробит",
        CorrectionScenario.WrongPaymentType   => "Неправильный способ оплаты",
        CorrectionScenario.WrongNomenclature  => "Неправильная номенклатура",
        CorrectionScenario.RealRefund         => "Реальный возврат (по бухгалтерии)",
        CorrectionScenario.WrongDate          => "Чек пробит другой датой",
        CorrectionScenario.ExpenseCorrection  => "Коррекция расхода",
        _ => s.ToString(),
    };

    /// <summary>Определяет какой Kind заказа подходит для сценария.</summary>
    public static OrderKind ToOrderKind(this CorrectionScenario s) => s switch
    {
        CorrectionScenario.FullCancel        => OrderKind.SingleRefund,
        CorrectionScenario.RealRefund        => OrderKind.SingleRefund,
        CorrectionScenario.IncreaseAmount    => OrderKind.SingleCorrection,
        CorrectionScenario.WrongDate         => OrderKind.SingleCorrection,
        CorrectionScenario.ExpenseCorrection => OrderKind.SingleCorrection,
        CorrectionScenario.DecreaseAmount    => OrderKind.RefundCorrectionPair,
        CorrectionScenario.WrongPaymentType  => OrderKind.RefundCorrectionPair,
        CorrectionScenario.WrongNomenclature => OrderKind.RefundCorrectionPair,
        _ => OrderKind.Regular,
    };

    /// <summary>Требует ли сценарий sell_correction (= нельзя через API АТОЛ).</summary>
    public static bool RequiresXmlOnly(this CorrectionScenario s) =>
        s.ToOrderKind() is OrderKind.SingleCorrection or OrderKind.RefundCorrectionPair;
}
