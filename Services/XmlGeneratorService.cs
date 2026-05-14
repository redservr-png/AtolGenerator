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
        var chk = new XElement("check",
            new XElement("timestamp",   DateTime.Now.ToString("dd.MM.yyyy HH:mm:ss")),
            new XElement("external_id", Guid.NewGuid().ToString("N")),
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

        // VATs
        var totalVat = c.Items.Sum(i => i.VatSum);
        receipt.Add(new XElement("vats",
            new XElement("vat",
                new XElement("type", c.CheckVatType),
                new XElement("sum",  Fmt(totalVat)))));

        receipt.Add(new XElement("total",   Fmt(c.Amount)));
        receipt.Add(new XElement("cashier", c.CashierName));

        // Доп. реквизит чека (тег 1192) — ФП исходного чека для исправительного
        if (!string.IsNullOrWhiteSpace(c.AdditionalCheckProps))
            receipt.Add(new XElement("additional_check_props", c.AdditionalCheckProps));

        return receipt;
    }

    private static XElement BuildCorrection(CheckData c)
    {
        // Тип оплаты: реализация → 2 (аванс), оплата → 0/1
        var payCode = c.Tab == "realization" ? "2"
                    : c.PaymentType == "cash" ? "0" : "1";

        // НДС: агент/услуга — берём из CheckVatType (задаётся при генерации); оплата → vat122; реализация → vat22
        string vatType; double vatSum;
        if (c.Agent is not null || c.IsService)
        {
            vatType = c.CheckVatType;  // "vat5" для Страхова, "none" для остальных
            vatSum  = vatType == "vat5"
                ? Math.Round(c.Amount * 5.0 / 100.0, 2)
                : c.Amount;
        }
        else if (c.Tab == "realization")
        {
            vatType = "vat22";
            vatSum  = Math.Round(c.Amount * 22.0 / 122.0, 2);  // 22/122 включено в цену
        }
        else
        {
            vatType = "vat122";
            vatSum  = Math.Round(c.Amount * 22.0 / 122.0, 2);  // 22/122 включено в цену
        }

        return new XElement("correction",
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
                    new XElement("sum",  Fmt(vatSum)))),
            new XElement("cashier", c.CashierName)
        );
    }

    private static string Fmt(double v) => v.ToString("F2", System.Globalization.CultureInfo.InvariantCulture);
}

// ── Internal DTO ─────────────────────────────────────────────────────────────

public class CheckData
{
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
    public string          AdditionalCheckProps { get; set; } = string.Empty;  // тег 1192: ФП исходного чека
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
