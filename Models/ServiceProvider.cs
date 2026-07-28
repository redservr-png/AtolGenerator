namespace AtolGenerator.Models;

public class ServiceProvider
{
    public ServiceProvider() { }

    public ServiceProvider(
        string service, string city, string name, string inn, string phone,
        string vatType = "none", string agentType = AgentTypeCatalog.DefaultCode)
    {
        Service = service;
        City = city;
        Name = name;
        Inn = inn;
        Phone = phone;
        VatType = vatType;
        AgentType = agentType;
    }

    public string Service { get; set; } = string.Empty;
    public string City    { get; set; } = string.Empty;
    public string Name    { get; set; } = string.Empty;
    public string Inn     { get; set; } = string.Empty;
    public string Phone   { get; set; } = string.Empty;
    public string VatType { get; set; } = "none";
    public string AgentType { get; set; } = AgentTypeCatalog.DefaultCode;
}
