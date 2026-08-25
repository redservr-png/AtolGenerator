using AtolGenerator.Models;

namespace AtolGenerator.Services;

public sealed class CorrectionPunchPlan
{
    public IReadOnlyList<GenerationResult> ApiReceipts { get; init; } = Array.Empty<GenerationResult>();
    public IReadOnlyList<string> CorrectionXmlPaths { get; init; } = Array.Empty<string>();
    public IReadOnlyList<string> ReceiptXmlPaths { get; init; } = Array.Empty<string>();

    public bool HasApiReceipts => ApiReceipts.Count > 0;
    public bool HasCorrectionXml => CorrectionXmlPaths.Count > 0;
}

public static class CorrectionPunchPlanner
{
    public static bool IsRepairBatch(IReadOnlyList<GenerationResult> results) =>
        results.Any(result =>
            !string.IsNullOrWhiteSpace(result.ObsidianCaseId) ||
            result.CheckData is { IsCorrection: true });

    public static CorrectionPunchPlan FromResults(IReadOnlyList<GenerationResult> results)
    {
        var apiReceipts = results
            .Where(result => result.CheckData is { IsCorrection: false })
            .ToList();

        var correctionXml = DistinctPaths(results
            .Where(result => result.CheckData is { IsCorrection: true })
            .Select(result => result.XmlPath));

        var receiptXml = DistinctPaths(apiReceipts.Select(result => result.XmlPath))
            .Where(path => !correctionXml.Contains(path, StringComparer.OrdinalIgnoreCase))
            .ToList();

        return new CorrectionPunchPlan
        {
            ApiReceipts = apiReceipts,
            CorrectionXmlPaths = correctionXml,
            ReceiptXmlPaths = receiptXml,
        };
    }

    private static List<string> DistinctPaths(IEnumerable<string?> paths) =>
        paths
            .Where(path => !string.IsNullOrWhiteSpace(path))
            .Select(path => path!)
            .Distinct(StringComparer.OrdinalIgnoreCase)
            .ToList();
}
