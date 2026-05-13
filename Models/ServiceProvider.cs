namespace AtolGenerator.Models;

public record ServiceProvider(
    string Service,
    string City,
    string Name,
    string Inn,
    string Phone,
    string VatType = "none");   // "vat5" для агентов на УСН 5%, "none" для остальных
