namespace AtolGenerator.Models;

public sealed class LocalPunchReportRow
{
    public DateTime? RegisteredAt { get; init; }
    public string Operation { get; init; } = string.Empty;
    public string OrderNumber { get; init; } = string.Empty;
    public string RealizationNumber { get; init; } = string.Empty;
    public double Amount { get; init; }
    public long? FiscalDocument { get; init; }
    public long? FiscalSign { get; init; }
    public string Uuid { get; init; } = string.Empty;
    public string Cashier { get; init; } = string.Empty;
    public string OfdUrl { get; init; } = string.Empty;
}

public sealed class AtolJournalReportRow
{
    public DateTime? RegisteredAt { get; init; }
    public DateTime? ReceivedAt { get; init; }
    public string CheckType { get; init; } = string.Empty;
    public string Operation { get; init; } = string.Empty;
    public string Status { get; init; } = string.Empty;
    public string Source { get; init; } = string.Empty;
    public double Amount { get; init; }
    public long? FiscalDocument { get; init; }
    public long? FiscalSign { get; init; }
    public string FiscalDriveNumber { get; init; } = string.Empty;
    public string Uuid { get; init; } = string.Empty;
    public string ExternalId { get; init; } = string.Empty;
    public string BaseNumber { get; init; } = string.Empty;
    public string BaseDate { get; init; } = string.Empty;
    public string OfdUrl { get; init; } = string.Empty;
    public string IncomingJson { get; init; } = string.Empty;
    public string ResultJson { get; init; } = string.Empty;
}

public sealed class OfdReportRow
{
    public DateTime? RegisteredAt { get; init; }
    public string Document { get; init; } = string.Empty;
    public string Operation { get; init; } = string.Empty;
    public string CalculationMethod { get; init; } = string.Empty;
    public double Amount { get; init; }
    public long? FiscalDocument { get; init; }
    public long? FiscalSign { get; init; }
    public string FiscalDriveNumber { get; init; } = string.Empty;
    public string KktRegistrationNumber { get; init; } = string.Empty;
    public string TradingPoint { get; init; } = string.Empty;
    public string KktName { get; init; } = string.Empty;
    public string ReceiptUrl { get; init; } = string.Empty;
    public string SourceFile { get; init; } = string.Empty;
}

public sealed class OfdReportReadResult
{
    public List<OfdReportRow> Rows { get; init; } = new();
    public bool IsTruncated { get; init; }
}

public sealed class OfdArchiveResult
{
    public List<OfdReportRow> Rows { get; init; } = new();
    public List<string> ImportedFiles { get; init; } = new();
    public List<string> TruncatedFiles { get; init; } = new();
    public List<string> FailedFiles { get; init; } = new();
}

public sealed class XmlReportCheck
{
    public int Index { get; init; }
    public DateTime? GeneratedAt { get; init; }
    public string ExternalId { get; init; } = string.Empty;
    public string Operation { get; init; } = string.Empty;
    public string RealizationNumber { get; init; } = string.Empty;
    public string BaseDate { get; init; } = string.Empty;
    public double Amount { get; init; }
    public string OriginalFiscalSign { get; init; } = string.Empty;
}

public sealed class OneCExportRow
{
    public string RealizationNumber { get; init; } = string.Empty;
    public string CheckType { get; init; } = string.Empty;
    public string WriteMode { get; init; } = string.Empty;
    public string ExternalId { get; init; } = string.Empty;
    public long? FiscalSign { get; init; }
    public long? FiscalDocument { get; init; }
    public DateTime? RegisteredAt { get; init; }
    public string Comment { get; init; } = string.Empty;
    public string OfdStatus { get; init; } = string.Empty;
    public string Status { get; init; } = string.Empty;
    public bool IsReady { get; init; }
}
