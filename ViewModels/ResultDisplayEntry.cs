using AtolGenerator.Models;

namespace AtolGenerator.ViewModels;

/// <summary>
/// Элемент для отображения в списке результатов.
/// Может быть одиночным чеком (Single) или парой исправительных (Refund + Correction).
/// </summary>
public class ResultDisplayEntry
{
    // ── Одиночный чек ────────────────────────────────────────────────────────
    public GenerationResult? Single { get; init; }

    // ── Исправительная пара ───────────────────────────────────────────────────
    public GenerationResult? Refund     { get; init; }   // sell_refund
    public GenerationResult? Correction { get; init; }   // sell_correction
    public string            PairLabel  { get; init; } = string.Empty; // номер реализации

    // ── Тип ──────────────────────────────────────────────────────────────────
    public bool IsPair => Refund is not null;

    // ── Удобные свойства для привязки ─────────────────────────────────────────
    public string DisplayOrderNum =>
        IsPair ? (Refund?.OrderNum ?? string.Empty)
               : (Single?.OrderNum ?? string.Empty);

    public string RefundXmlFilename     => Refund?.XmlFilename     ?? string.Empty;
    public string CorrectionXmlFilename => Correction?.XmlFilename ?? string.Empty;
}
