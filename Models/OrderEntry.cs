using System.Text.Json.Serialization;

namespace AtolGenerator.Models;

public class OrderEntry
{
    public string          ObsidianCaseId    { get; set; } = string.Empty;
    public string          OrderNum         { get; set; } = string.Empty;
    public string          OrderDate        { get; set; } = string.Empty;  // "DD.MM.YYYY HH:MM:SS"
    [JsonIgnore]
    public string          SourceDocumentDate { get; set; } = string.Empty;
    public double          Amount           { get; set; }
    public string          CustomerName     { get; set; } = string.Empty;
    /// <summary>Позиции правильного чека по актуальным данным 1С.</summary>
    public List<OrderItem> Items            { get; set; } = new();

    /// <summary>
    /// Позиции исходного ошибочного чека. Для исправлений по ФФД 1.05 они
    /// используются в обычном обратном чеке и могут отличаться от Items.
    /// </summary>
    public List<OrderItem> OriginalItems    { get; set; } = new();
    public ServiceProvider? AgentInfo       { get; set; }
    public string          CorrectionDate   { get; set; } = string.Empty;  // DD.MM.YYYY
    public string          CorrectionNumber { get; set; } = string.Empty;
    public bool            IsService        { get; set; }
    public bool            IsOwnService     { get; set; }  // собственная услуга организации, агент не требуется
    public string          ServiceType      { get; set; } = string.Empty;  // "доставка" / "сборка" из текста заказа
    public string          City             { get; set; } = string.Empty;  // подразделение из 1С

    // ── Поля для исправительных чеков (загружаются из Obsidian-кейсов) ──────────
    /// <summary>Тип строки: обычный заказ или исправительный (один из вариантов).</summary>
    public OrderKind            Kind                 { get; set; } = OrderKind.Regular;

    /// <summary>Тип исходного документа из 1С (для парсинга Obsidian).</summary>
    public SourceDocumentType   DocumentType         { get; set; } = SourceDocumentType.Unknown;

    /// <summary>Сценарий коррекции (определяется автодетектом, можно править вручную).</summary>
    public CorrectionScenario   CorrectionScenario   { get; set; } = CorrectionScenario.Unknown;

    /// <summary>ФП исходного ошибочного чека для сопоставления, плана и служебной записки.</summary>
    public string               OriginalFiscalNumber { get; set; } = string.Empty;

    /// <summary>Сумма в исходном ошибочном чеке (отличается от Amount при «чек большей суммой»).</summary>
    public double?              OriginalCheckAmount  { get; set; }

    /// <summary>Текущая сумма в правильно скорректированном документе 1С.</summary>
    public double?              CorrectAmount        { get; set; }

    /// <summary>Описание ошибки из Obsidian (исходный текст после номера+даты).</summary>
    public string               Notes                { get; set; } = string.Empty;

    /// <summary>Был ли исходный чек пробит наличными (true) или картой (false).</summary>
    public bool?                OriginalPaymentWasCash { get; set; }

    /// <summary>Должен быть правильный тип оплаты — наличные (true) или карта (false).</summary>
    public bool?                CorrectPaymentIsCash   { get; set; }

    /// <summary>Операция отмены исходного чека, рассчитанная планом исправления.</summary>
    public string               PlannedReverseOperation { get; set; } = string.Empty;

    /// <summary>Операция правильного чека, рассчитанная по дате и типу документа.</summary>
    public string               PlannedCorrectOperation { get; set; } = string.Empty;

    /// <summary>Ставка правильного чека из правила 1С/агента.</summary>
    public string               PlannedVatType { get; set; } = string.Empty;

    /// <summary>Дата исходного ошибочного чека из отчёта ОФД.</summary>
    public DateTime?            OriginalCheckDate { get; set; }

    /// <summary>Операция исходного ошибочного чека из отчёта ОФД.</summary>
    public string               OriginalCheckOperation { get; set; } = string.Empty;

    // Реквизиты документа и исходного чека, полученные непосредственно из 1С.
    public string               OneCComment { get; set; } = string.Empty;
    public string               OneCCheckNumber { get; set; } = string.Empty;
    public DateTime?            OneCCheckDate { get; set; }

    /// <summary>Удобное свойство для UI: true = строка является исправительной.</summary>
    public bool IsCorrection => Kind != OrderKind.Regular;

    public string AgentVatDisplay => IsOwnService
        ? "Без агента / vat22"
        : AgentInfo is null
            ? "Товарная реализация / vat22"
            : $"{AgentInfo.Name} / {AgentInfo.VatType}";
}
