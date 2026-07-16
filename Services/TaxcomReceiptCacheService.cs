using System.IO;
using System.Text.Encodings.Web;
using System.Text.Json;
using AtolGenerator.Helpers;
using AtolGenerator.Models;

namespace AtolGenerator.Services;

public static class TaxcomReceiptCacheService
{
    private static readonly JsonSerializerOptions JsonOptions = new()
    {
        WriteIndented = true,
        Encoder = JavaScriptEncoder.UnsafeRelaxedJsonEscaping,
    };

    public static List<OfdReportRow> Load()
    {
        try
        {
            if (!File.Exists(FileHelper.TaxcomReceiptCachePath)) return new List<OfdReportRow>();
            var json = File.ReadAllText(FileHelper.TaxcomReceiptCachePath);
            return JsonSerializer.Deserialize<List<OfdReportRow>>(json) ?? new List<OfdReportRow>();
        }
        catch
        {
            return new List<OfdReportRow>();
        }
    }

    public static void Upsert(IEnumerable<OfdReportRow> rows)
    {
        var merged = ReportImportService.MergeOfdRows(rows.Concat(Load()));
        Directory.CreateDirectory(FileHelper.TaxcomReportDir);
        var json = JsonSerializer.Serialize(merged, JsonOptions);
        File.WriteAllText(FileHelper.TaxcomReceiptCachePath, json);
    }
}
