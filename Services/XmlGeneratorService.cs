using System.Text;
using System.Xml;
using System.Xml.Linq;
using AtolGenerator.Constants;
using AtolGenerator.Models;

namespace AtolGenerator.Services;

public static class XmlGeneratorService
{
    public static void GenerateFile(IEnumerable<CheckData> checks, string filepath)
    {
        var main = new XElement("main", checks.Select(BuildCheckXml));
        var settings = new XmlWriterSettings
        {
            Indent             = true,
            IndentChars        = "    ",
            Encoding           = new UTF8Encoding(false),
            OmitXmlDeclaration = false,
        };
        using var writer = XmlWriter.Create(filepath, settings);
        main.Save(writer);
    }

    private static XElement BuildCheckXml(CheckData c)
    {
        if (string.IsNullOrWhiteSpace(c.ExternalId))
            c.ExternalId = Guid.NewGuid().ToString("N");

        var chk = new XElement("check",
            new XElement("timestamp",   DateTime.Now.ToString("dd.MM.yyyy HH:mm:ss")),
            new XElement("external_id", c.ExternalId),
            new XElement("is_bso",      "false")
        );

        if (c.IsCorrection)
            chk.Add(BuildCorrection(c));
        else
            chk.Add(BuildReceipt(c));

        return chk;
    }

    private static XElement BuildReceipt(CheckData c)
    {
        var receipt = new XElement("receipt",
            new XElement("operation", c.OperationType),
            new XElement("client",
                new XElement("phone", AppConstants.PhoneBuyer)),
            new XElement("company",
                new XElement("email",           AppConstants.EmailOrg),
                new XElement("sno",             AppConstants.Sno),
                new XElement("inn",             AppConstants.InnOrg),
                new XElement("payment_address", AppConstants.PaymentAddress))
        );

        // Agent info
        if (c.Agent is not null)
        {
            receipt.Add(new XElement("agent_info",
                new XElement("type", "commission_agent")));
            receipt.Add(new XElement("supplier_info",
                new XElement("phones",
                    new XElement("phone", c.Agent.Phone)),
                new XElement("name", c.Agent.Name),
                new XElement("inn",  c.Agent.Inn)));
        }

        // Items
        var itemsEl = new XElement("items");
        foreach (var item in c.Items)
        {
            var it = new XElement("item",
                new XElement("name",           item.Name),
                new XElement("price",          Fmt(item.Price)),
                new XElement("quantity",       item.Quantity.ToString(System.Globalization.CultureInfo.InvariantCulture)),
                new XElement("sum",            Fmt(item.Sum)),
                new XElement("payment_method", item.PaymentMethod),
                new XElement("payment_object", item.PaymentObject)
            );
            if (c.Agent is not null && item.IsService)
                it.Add(new XElement("agent_info", new XElement("type", "commission_agent")));
            it.Add(new XElement("vat",
                new XElement("type", item.VatType),
                new XElement("sum",  Fmt(item.VatSum))));
            itemsEl.Add(it);
        }
        receipt.Add(itemsEl);

        // Payment: realization → 2 (аванс); cash → 0; card → 1 (безналичные)
        var payType = c.Tab == "realization" ? "2"
                    : c.PaymentType == "cash" ? "0" : "1";
        receipt.Add(new XElement("payments",
            new XElement("payment",
                new XElement("type", payType),
                new XElement("sum",  Fmt(c.Amount)))));

        // VAT totals are grouped by item rate so mixed-rate receipts remain valid.
        var vatGroups = c.Items
            .GroupBy(item => VatRateCatalog.Normalize(item.VatType, c.CheckVatType),
                StringComparer.OrdinalIgnoreCase)
            .Select(group => new
            {
                Type = group.Key,
                Sum = group.Sum(item => item.VatSum),
            })
            .ToList();
        if (vatGroups.Count == 0)
        {
            var fallbackType = VatRateCatalog.Normalize(c.CheckVatType, "none");
            vatGroups.Add(new
            {
                Type = fallbackType,
                Sum = VatRateCatalog.CalculateFiscalSum(c.Amount, fallbackType),
            });
        }

        receipt.Add(new XElement("vats", vatGroups.Select(group =>
            new XElement("vat",
                new XElement("type", group.Type),
                new XElement("sum", Fmt(group.Sum))))));

        receipt.Add(new XElement("total", Fmt(c.Amount)));

        // Tag 1192 must precede cashier according to the ATOL XML schema.
        if (!string.IsNullOrWhiteSpace(c.AdditionalCheckProps))
            receipt.Add(new XElement("additional_check_props", c.AdditionalCheckProps.Trim()));

        receipt.Add(new XElement("cashier", c.CashierName));

        // Доп. реквизит пользователя (теги 1084/1085/1086) — например, № реализации
        if (!string.IsNullOrWhiteSpace(c.UserAttributeValue))
        {
            receipt.Add(new XElement("additional_user_props",
                new XElement("name",  string.IsNullOrWhiteSpace(c.UserAttributeName) ? "Номер реализации" : c.UserAttributeName),
                new XElement("value", c.UserAttributeValue)));
        }

        return receipt;
    }

    private static XElement BuildCorrection(CheckData c)
    {
        // Тип оплаты: реализация → 2 (аванс), оплата → 0/1
        var payCode = c.Tab == "realization" ? "2"
                    : c.PaymentType == "cash" ? "0" : "1";

        var defaultVatType = c.Tab == "realization" ? "vat22" : "vat122";
        var vatType = VatRateCatalog.Normalize(c.CheckVatType, defaultVatType);
        var vatSum = VatRateCatalog.CalculateFiscalSum(c.Amount, vatType);

        // В ФФД 1.05 XSD АТОЛ не допускает additional_check_props и
        // additional_user_props внутри <correction>. Номер реализации уже
        // передаётся в base_number (тег 1179), ФП хранится вне XML.
        var correction = new XElement("correction",
            new XElement("operation", c.OperationType),
            new XElement("company",
                new XElement("sno",             AppConstants.Sno),
                new XElement("inn",             AppConstants.InnOrg),
                new XElement("payment_address", AppConstants.PaymentAddress)),
            new XElement("correction_info",
                new XElement("type",        "self"),
                new XElement("base_date",   c.CorrectionBaseDate),
                new XElement("base_number", c.CorrectionBaseNumber)),
            new XElement("payments",
                new XElement("payment",
                    new XElement("type", payCode),
                    new XElement("sum",  Fmt(c.Amount)))),
            new XElement("vats",
                new XElement("vat",
                    new XElement("type", vatType),
                    new XElement("sum",  Fmt(vatSum))))
        );

        correction.Add(new XElement("cashier", c.CashierName));

        return correction;
    }

    private static string Fmt(double v) => v.ToString("F2", System.Globalization.CultureInfo.InvariantCulture);
}

// ── Internal DTO ─────────────────────────────────────────────────────────────

public class CheckData
{
    public string          ExternalId           { get; set; } = string.Empty;
    public string          OperationType        { get; set; } = "sell";
    public bool            IsCorrection         { get; set; }
    public string          Tab                  { get; set; } = "payment";
    public double          Amount               { get; set; }
    public string          PaymentType          { get; set; } = "card";
    public string          CheckVatType         { get; set; } = "vat122";
    public List<CheckItem> Items                { get; set; } = new();
    public ServiceProvider? Agent               { get; set; }
    public bool            IsService            { get; set; }
    public string          CorrectionBaseDate   { get; set; } = string.Empty;
    public string          CorrectionBaseNumber { get; set; } = "б/н";
    public string          AdditionalCheckProps { get; set; } = string.Empty;  // Тег 1192: ФП исходного чека.
    public string          UserAttributeName    { get; set; } = string.Empty;  // тег 1085: наименование доп.реквизита
    public string          UserAttributeValue   { get; set; } = string.Empty;  // тег 1086: значение (например, № реализации)
    public string          CashierName          { get; set; } = AppConstants.CashierName;
    public string          CashierShort         { get; set; } = AppConstants.CashierShort;
}

public class CheckItem
{
    public string Name          { get; set; } = string.Empty;
    public double Price         { get; set; }
    public double Quantity      { get; set; } = 1;
    public double Sum           { get; set; }
    public string PaymentMethod { get; set; } = "full_prepayment";
    public string PaymentObject { get; set; } = "payment";
    public string VatType       { get; set; } = "vat122";
    public double VatSum        { get; set; }
    public bool   IsService     { get; set; }
}
