namespace AtolGenerator.Models;

public class OrderEntry
{
    public string          OrderNum         { get; set; } = string.Empty;
    public string          OrderDate        { get; set; } = string.Empty;  // "DD.MM.YYYY HH:MM:SS"
    public double          Amount           { get; set; }
    public string          CustomerName     { get; set; } = string.Empty;
    public List<OrderItem> Items            { get; set; } = new();
    public ServiceProvider? AgentInfo       { get; set; }
    public string          CorrectionDate   { get; set; } = string.Empty;  // DD.MM.YYYY
    public string          CorrectionNumber { get; set; } = string.Empty;
    public bool            IsService        { get; set; }  // true = Агентский договор → без НДС
    public string          City             { get; set; } = string.Empty;  // подразделение из 1С
}
