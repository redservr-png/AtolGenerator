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
    /// → Один обычный обратный чек с ФП исходного чека в теге 1192.
    /// </summary>
    FullCancel,

    /// <summary>
    /// Чек пробит большей суммой — нужно уменьшить базу.
    /// «Чек большей суммой».
    /// → Обычная отмена ошибочной суммы + правильный обычный чек; тег 1192 в обоих.
    /// </summary>
    CheckLargerAmount,

    /// <summary>
    /// Чек пробит меньшей суммой — нужно увеличить базу.
    /// «Чек меньшей суммой».
    /// Чек был пробит, поэтому нужно сначала отменить его, затем пробить правильный.
    /// → Два обычных чека с ФП исходного чека в теге 1192.
    /// </summary>
    CheckSmallerAmount,

    /// <summary>
    /// Чек вообще не был пробит — есть документ в 1С, но фискального чека нет.
    /// «Чек не пробит», «не вышел», «не в день оплаты», «не пробит чек в день оплаты».
    /// → Только sell_correction на полную сумму (refund не нужен — чека и так нет).
    /// </summary>
    CheckNotPunched,

    /// <summary>
    /// Неправильный способ оплаты — отменить и пробить заново.
    /// «Чек пробит наличными» (а было картой), «перепутали оплаты», «нал и безнал».
    /// → Обычная отмена + правильный обычный чек с верным payments.type; тег 1192 в обоих.
    /// </summary>
    WrongPaymentType,

    /// <summary>
    /// Неправильная номенклатура — заменить.
    /// «Перепутали номенклатуру», «ошибочно провела сборку», «другая номенклатура».
    /// → Обычная отмена + правильный обычный чек с верной номенклатурой; тег 1192 в обоих.
    /// </summary>
    WrongNomenclature,

    /// <summary>
    /// Реальный возврат денег (фактический).
    /// «Возврат по бухгалтерии», «деньги вернулись клиенту».
    /// → Один sell_refund (самостоятельная операция, не исправление).
    /// </summary>
    // Устаревшее значение для чтения сохранённых данных. Реальные возвраты
    // оформляются в отдельном сценарии «Возвраты по заказам».
    RealRefund,

    /// <summary>
    /// Чек пробит другой датой — корректировка.
    /// «Чек другой датой», «дата следующая».
    /// → Отмена исходного и правильный обычный чек с тегом 1192 в обоих.
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

    /// <summary>Пара обычных чеков: отмена исходного + правильный чек</summary>
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
        CorrectionScenario.CheckLargerAmount  => "Чек большей суммой",
        CorrectionScenario.CheckSmallerAmount => "Чек меньшей суммой",
        CorrectionScenario.CheckNotPunched    => "Чек не пробит",
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
        CorrectionScenario.FullCancel         => OrderKind.SingleRefund,
        CorrectionScenario.RealRefund         => OrderKind.SingleRefund,
        CorrectionScenario.CheckNotPunched    => OrderKind.SingleCorrection,
        CorrectionScenario.WrongDate          => OrderKind.SingleCorrection,
        CorrectionScenario.ExpenseCorrection  => OrderKind.SingleCorrection,
        CorrectionScenario.CheckLargerAmount  => OrderKind.RefundCorrectionPair,
        CorrectionScenario.CheckSmallerAmount => OrderKind.RefundCorrectionPair,
        CorrectionScenario.WrongPaymentType   => OrderKind.RefundCorrectionPair,
        CorrectionScenario.WrongNomenclature  => OrderKind.RefundCorrectionPair,
        _ => OrderKind.Regular,
    };

    /// <summary>Требует ли сценарий отдельного XML-процесса без отправки через API.</summary>
    public static bool RequiresXmlOnly(this CorrectionScenario s) =>
        s is not CorrectionScenario.Unknown and not CorrectionScenario.RealRefund;
}
