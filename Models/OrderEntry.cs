namespace AtolGenerator.Models;

public class OrderEntry
{
    public string          OrderNum         { get; set; } = string.Empty;
    public string          OrderDate        { get; set; } = string.Empty;  // "DD.MM.YYYY HH:MM:SS"
    public double          Amount           { get; set; }
    public string          CustomerName     { get; set; } = string.Empty;
    public List<OrderItem> Items            { get; set; } = new();
    public ServiceProvider? AgentInfo       { get; set; }
    public string          CorrectionDate   { get; set; } = string.Empty;  // DD.MM.YYYY
    public string          CorrectionNumber { get; set; } = string.Empty;
    public bool            IsService        { get; set; }  // true = Агентский договор → без НДС
    public string          ServiceType      { get; set; } = string.Empty;  // "доставка" / "сборка" из текста заказа
    public string          City             { get; set; } = string.Empty;  // подразделение из 1С

    // ── Поля для исправительных чеков (загружаются из Obsidian-кейсов) ──────────
    /// <summary>Тип строки: обычный заказ или исправительный (один из вариантов).</summary>
    public OrderKind            Kind                 { get; set; } = OrderKind.Regular;

    /// <summary>Тип исходного документа из 1С (для парсинга Obsidian).</summary>
    public SourceDocumentType   DocumentType         { get; set; } = SourceDocumentType.Unknown;

    /// <summary>Сценарий коррекции (определяется автодетектом, можно править вручную).</summary>
    public CorrectionScenario   CorrectionScenario   { get; set; } = CorrectionScenario.Unknown;

    /// <summary>ФП (фискальный признак) исходного ошибочного чека — для тега 1192.</summary>
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

    /// <summary>Удобное свойство для UI: true = строка является исправительной.</summary>
    public bool IsCorrection => Kind != OrderKind.Regular;
}
