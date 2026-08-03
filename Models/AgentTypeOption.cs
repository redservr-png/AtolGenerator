namespace AtolGenerator.Models;

public sealed record AgentTypeOption(string Code, string Label);

public static class AgentTypeCatalog
{
    public const string DefaultCode = "another";

    public static IReadOnlyList<AgentTypeOption> All { get; } = new[]
    {
        new AgentTypeOption("bank_paying_agent", "Банковский платёжный агент"),
        new AgentTypeOption("bank_paying_subagent", "Банковский платёжный субагент"),
        new AgentTypeOption("paying_agent", "Платёжный агент"),
        new AgentTypeOption("paying_subagent", "Платёжный субагент"),
        new AgentTypeOption("attorney", "Поверенный"),
        new AgentTypeOption("commission_agent", "Комиссионер"),
        new AgentTypeOption("another", "Другой тип агента"),
    };

    public static bool IsKnown(string? code) =>
        All.Any(option => string.Equals(option.Code, code?.Trim(), StringComparison.OrdinalIgnoreCase));

    public static string Normalize(string? code) =>
        All.FirstOrDefault(option => string.Equals(
            option.Code, code?.Trim(), StringComparison.OrdinalIgnoreCase))?.Code ?? DefaultCode;

    public static string LabelFor(string? code) =>
        All.FirstOrDefault(option => string.Equals(
            option.Code, code?.Trim(), StringComparison.OrdinalIgnoreCase))?.Label ?? "Другой тип агента";
}
