using AtolGenerator.Models;

namespace AtolGenerator.Services;

public static class ObsidianExpectedCheckService
{
    public static int ReplaceFromGeneration(
        ObsidianCaseState state,
        IEnumerable<GenerationResult> results)
    {
        var checks = results
            .Where(x => !string.IsNullOrWhiteSpace(x.ExternalId))
            .GroupBy(x => x.ExternalId, StringComparer.OrdinalIgnoreCase)
            .Select(group => group.Last())
            .Select(result => new ObsidianExpectedCheck
            {
                ExternalId = result.ExternalId,
                Operation = result.CheckData?.OperationType ?? string.Empty,
                Amount = result.Amount,
                GeneratedAt = DateTime.Now,
            })
            .ToList();

        state.ExpectedChecks = checks;
        state.CheckConfirmed = false;
        return checks.Count;
    }

    public static int ReplaceFromXml(
        ObsidianCaseRecord record,
        ObsidianCaseState state,
        IReadOnlyList<XmlReportCheck> xmlChecks)
    {
        var documentNumbers = EnumerateDocuments(record)
            .SelectMany(document => new[] { document.OrderNum, document.CorrectionNumber })
            .Where(value => !string.IsNullOrWhiteSpace(value))
            .Select(NormalizeDocumentNumber)
            .ToHashSet(StringComparer.OrdinalIgnoreCase);

        if (documentNumbers.Count == 0) return 0;

        var matches = xmlChecks
            .Where(check => !string.IsNullOrWhiteSpace(check.ExternalId))
            .Where(check => documentNumbers.Contains(NormalizeDocumentNumber(check.RealizationNumber)))
            .GroupBy(check => check.ExternalId, StringComparer.OrdinalIgnoreCase)
            .Select(group => group.Last())
            .Select(check => new ObsidianExpectedCheck
            {
                ExternalId = check.ExternalId,
                Operation = check.Operation,
                Amount = check.Amount,
                GeneratedAt = check.GeneratedAt ?? DateTime.Now,
            })
            .ToList();

        if (matches.Count == 0) return 0;

        state.ExpectedChecks = matches;
        state.CheckConfirmed = false;
        state.LastMessage = $"Связь восстановлена из XML: {matches.Count} чек(ов)";
        state.UpdatedAt = DateTime.Now;
        return matches.Count;
    }

    private static IEnumerable<OrderEntry> EnumerateDocuments(ObsidianCaseRecord record)
    {
        yield return record.PrimaryDocument;
        foreach (var related in record.RelatedDocuments)
            yield return related;
    }

    private static string NormalizeDocumentNumber(string value) =>
        ReportImportService.NormalizeRealizationNumber(value).Trim();
}
