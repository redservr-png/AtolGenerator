namespace AtolGenerator.Models;

public class ServiceProvider
{
    public ServiceProvider() { }

    public ServiceProvider(
        string service, string city, string name, string inn, string phone,
        string vatType = "none")
    {
        Service = service;
        City = city;
        Name = name;
        Inn = inn;
        Phone = phone;
        VatType = vatType;
    }

    public string Service { get; set; } = string.Empty;
    public string City    { get; set; } = string.Empty;
    public string Name    { get; set; } = string.Empty;
    public string Inn     { get; set; } = string.Empty;
    public string Phone   { get; set; } = string.Empty;
    public string VatType { get; set; } = "none";
}
