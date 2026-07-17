using AtolGenerator.Models;

namespace AtolGenerator.Services;

public static class ServiceClassificationService
{
    public const string OwnDeliveryDepartment = "Интернет-магазин (Продажи)";
    public const string OwnDeliveryItem = "Услуга по доставке (Россия)";

    public static bool IsOwnDeliveryDepartmentName(string department) =>
        string.Equals(Normalize(department), OwnDeliveryDepartment, StringComparison.OrdinalIgnoreCase);

    public static bool IsOwnDelivery(string department, IEnumerable<string> itemNames) =>
        IsOwnDeliveryDepartmentName(department) &&
        itemNames.Any(name => string.Equals(
            Normalize(name), OwnDeliveryItem, StringComparison.OrdinalIgnoreCase));

    public static bool ApplyOwnDeliveryRule(OrderEntry order)
    {
        if (!IsOwnDelivery(order.City, order.Items.Select(item => item.Name))) return false;

        order.IsService = true;
        order.IsOwnService = true;
        order.ServiceType = "Доставка";
        order.AgentInfo = null;
        return true;
    }

    public static bool ApplyOwnDeliveryRule(OneCRealization realization)
    {
        if (!IsOwnDelivery(realization.City, realization.Items.Select(item => item.Name))) return false;

        realization.IsService = true;
        realization.IsOwnService = true;
        realization.ServiceType = "Доставка";
        realization.AgentInfo = null;
        return true;
    }

    public static string ResolveVatType(
        bool isOwnService,
        ServiceProvider? agent,
        string tab)
    {
        if (isOwnService) return tab == "payment" ? "vat122" : "vat22";

        var agentVat = VatRateCatalog.Normalize(agent?.VatType, "none");
        return tab == "payment" ? ToCalculatedRate(agentVat) : agentVat;
    }

    public static string ToCalculatedRate(string vatType) => VatRateCatalog.Normalize(vatType) switch
    {
        "vat5" => "vat105",
        "vat7" => "vat107",
        "vat10" => "vat110",
        "vat20" => "vat120",
        "vat22" => "vat122",
        var value => value,
    };

    private static string Normalize(string? value) => string.Join(
        " ",
        (value ?? string.Empty).Split((char[]?)null, StringSplitOptions.RemoveEmptyEntries));
}
