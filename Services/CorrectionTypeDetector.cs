using AtolGenerator.Models;

namespace AtolGenerator.Services;

/// <summary>
/// Определяет сценарий коррекции по описанию ошибки из Obsidian.
/// Логика основана на «Исправить чеки.md»: разбор ключевых слов.
/// </summary>
public static class CorrectionTypeDetector
{
    /// <summary>
    /// Анализирует Notes + DocumentType и выставляет CorrectionScenario + Kind.
    /// Также пытается определить OriginalPaymentWasCash/CorrectPaymentIsCash для WrongPaymentType.
    /// </summary>
    public static void Detect(OrderEntry e)
    {
        var n = e.Notes?.ToLowerInvariant() ?? string.Empty;

        // ── 1. Чек большей / меньшей суммой (приоритетно — может пересекаться с другими) ──
        if (Contains(n, "больш", "сум") || n.Contains("пробит больше"))
        {
            e.CorrectionScenario = CorrectionScenario.CheckLargerAmount;
        }
        else if (Contains(n, "меньш", "сум"))
        {
            e.CorrectionScenario = CorrectionScenario.CheckSmallerAmount;
        }
        // ── 2. «Не пробит» / «Не вышел» — нужно доплатить базу (без refund — чека и так нет) ──
        else if (n.Contains("не пробит") || n.Contains("не вышел")
              || n.Contains("чек не пробит") || n.Contains("чека не было")
              || n.Contains("чека нет")
              || n.Contains("не день в день") || n.Contains("следующий день")
              || n.Contains("на следующий"))
        {
            e.CorrectionScenario = CorrectionScenario.CheckNotPunched;
        }
        // ── 3. Способ оплаты перепутали ───────────────────────────────────────────
        else if (n.Contains("пробит наличными") || n.Contains("пробит нал")
              || n.Contains("приняли наличными") || n.Contains("оплатили наличкой")
              || n.Contains("оплатили наличными"))
        {
            e.CorrectionScenario        = CorrectionScenario.WrongPaymentType;
            e.OriginalPaymentWasCash    = true;   // пробит наличными
            e.CorrectPaymentIsCash      = false;  // должно быть картой
        }
        else if (n.Contains("перепутал") && (n.Contains("опл") || n.Contains("нал") || n.Contains("безнал")))
        {
            e.CorrectionScenario = CorrectionScenario.WrongPaymentType;
        }
        // ── 4. Перепутали номенклатуру ────────────────────────────────────────────
        else if (n.Contains("перепутали номенклатуру") || n.Contains("другая номенклатура")
              || n.Contains("другую номенклатуру") || n.Contains("ошибочно провела сборку")
              || n.Contains("по основному договору") || n.Contains("по услуге пробили"))
        {
            e.CorrectionScenario = CorrectionScenario.WrongNomenclature;
        }
        // ── 5. Другая дата ────────────────────────────────────────────────────────
        else if (n.Contains("другой датой") || n.Contains("чек другой дат")
              || n.Contains("не день в день") || n.Contains("дата другая"))
        {
            e.CorrectionScenario = CorrectionScenario.WrongDate;
        }
        // ── 6. Возврат по бухгалтерии — реальный возврат ──────────────────────────
        else if (n.Contains("возврат по бух") || n.Contains("деньги вернулись")
              || n.Contains("возврат денег"))
        {
            e.CorrectionScenario = CorrectionScenario.RealRefund;
        }
        // ── 7. Расход (для агентских РКО) ─────────────────────────────────────────
        else if (e.DocumentType == SourceDocumentType.CashExpense
              || n.Contains("чек коррекции расход") || n.Contains("без агента"))
        {
            e.CorrectionScenario = CorrectionScenario.ExpenseCorrection;
        }
        // ── 8. Все остальные «лишний/не было/удалить/не должно» — полная отмена ───
        else if (n.Contains("не было") || n.Contains("не должно")
              || n.Contains("лишний")    || n.Contains("деньги не списались")
              || n.Contains("оплаты не было") || n.Contains("реализации не было")
              || n.Contains("оплата на удалении") || n.Contains("документ удалён")
              || n.Contains("документ удален") || n.Contains("док удсален")
              || n.Contains("док удалён") || n.Contains("док удален")
              || n.Contains("реализация удалена") || n.Contains("реализация отменена")
              || n.Contains("помечена на удаление") || n.Contains("помечен на удаление")
              || n.Contains("помечена удаление") || n.Contains("чек удалить")
              || n.Contains("не платил") || n.Contains("чек аннулировался")
              || n.Contains("чек отменен") || n.Contains("чек отменён")
              || n.Contains("вернулся на склад") || n.Contains("товар вернулся")
              || n.Contains("сбой в атол") || n.Contains("сбой в аотле")
              || n.Contains("дублировал") || n.Contains("2 раза"))
        {
            e.CorrectionScenario = CorrectionScenario.FullCancel;
        }
        // ── 9. Не смогли распознать — пользователь решит ──────────────────────────
        else
        {
            e.CorrectionScenario = CorrectionScenario.Unknown;
        }

        // Выставляем Kind по выбранному сценарию.
        // Для Unknown — НЕ ставим дефолт; помечаем как SingleRefund чтобы строка
        // выделилась цветной полосой, но в UI рядом будет предупреждение
        // «выберите сценарий», и при генерации XML такие строки пропускаются.
        e.Kind = e.CorrectionScenario.ToOrderKind();
        if (e.CorrectionScenario == CorrectionScenario.Unknown)
            e.Kind = OrderKind.SingleRefund;
    }

    /// <summary>Применяет автодетект ко всем записям списка.</summary>
    public static void DetectAll(IEnumerable<OrderEntry> entries)
    {
        foreach (var e in entries) Detect(e);
    }

    private static bool Contains(string text, params string[] all)
        => all.All(text.Contains);
}
