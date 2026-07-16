namespace AtolGenerator.Models;

public sealed class ObsidianCaseRecord
{
    public string CaseId { get; set; } = string.Empty;
    public bool IsCompleted { get; set; }
    public int LineNumber { get; set; }
    public string Period { get; set; } = string.Empty;
    public string RawBody { get; set; } = string.Empty;
    public CorrectionScenario? ScenarioOverride { get; set; }
    public OrderEntry PrimaryDocument { get; set; } = new();
    public List<OrderEntry> RelatedDocuments { get; set; } = new();
}

public sealed class ObsidianExpectedCheck
{
    public string ExternalId { get; set; } = string.Empty;
    public string Operation { get; set; } = string.Empty;
    public double Amount { get; set; }
    public DateTime GeneratedAt { get; set; }
    public long? FiscalSign { get; set; }
    public long? FiscalDocument { get; set; }
}

public sealed class ObsidianOriginalReceipt
{
    public string Source { get; set; } = string.Empty;
    public DateTime? RegisteredAt { get; set; }
    public string Document { get; set; } = string.Empty;
    public string Operation { get; set; } = string.Empty;
    public double Amount { get; set; }
    public long? FiscalSign { get; set; }
    public long? FiscalDocument { get; set; }
    public string ReceiptUrl { get; set; } = string.Empty;
}

public enum OriginalReceiptLookupState
{
    NotChecked = 0,
    Found,
    NotFound,
    MissingFiscalSign,
    Ambiguous,
}

public enum CorrectionPlanStatus
{
    NeedsOneC = 0,
    NeedsOriginalReceipt,
    NeedsScenario,
    NeedsServiceRule,
    DeferredCorrectionReceipt,
    Ready,
}

public sealed class CorrectionPlanCheck
{
    public int Sequence { get; set; }
    public string Operation { get; set; } = string.Empty;
    public string Title { get; set; } = string.Empty;
    public double Amount { get; set; }
    public string PaymentType { get; set; } = string.Empty;
    public string VatType { get; set; } = string.Empty;
    public bool RequiresItems { get; set; }
    public bool UsesOriginalFiscalSign { get; set; }
    public bool XmlOnly { get; set; }
}

public sealed class CorrectionPlan
{
    public CorrectionPlanStatus Status { get; set; }
    public string Message { get; set; } = string.Empty;
    public bool IsSameDay { get; set; }
    public List<CorrectionPlanCheck> Checks { get; set; } = new();
    public bool IsReady => Status == CorrectionPlanStatus.Ready && Checks.Count > 0;
}

public sealed class ObsidianCaseState
{
    public string CaseId { get; set; } = string.Empty;
    public bool SentToWork { get; set; }
    public bool ServiceNoteVerified { get; set; }
    public bool CheckConfirmed { get; set; }
    public bool OneCRecorded { get; set; }
    public double? CorrectAmount { get; set; }
    public string OriginalFiscalNumber { get; set; } = string.Empty;
    public string ServiceNotePath { get; set; } = string.Empty;
    public string LastMessage { get; set; } = string.Empty;
    public DateTime? UpdatedAt { get; set; }
    public List<ObsidianExpectedCheck> ExpectedChecks { get; set; } = new();
    public OrderEntry? OneCSnapshot { get; set; }
    public ObsidianOriginalReceipt? OriginalReceipt { get; set; }
    public OriginalReceiptLookupState OriginalReceiptLookupState { get; set; }
    public CorrectionScenario? ScenarioOverride { get; set; }
}
