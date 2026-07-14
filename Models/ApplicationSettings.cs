using AtolGenerator.Constants;

namespace AtolGenerator.Models;

public class ApplicationSettings
{
    public string ThemeKey                { get; set; } = "light";
    public string SelectedCashierShortName { get; set; } = string.Empty;
    public string ObsidianFilePath         { get; set; } = string.Empty;
    public string LastAtolReportPath       { get; set; } = string.Empty;
    public string LastOfdReportPath        { get; set; } = string.Empty;
    public bool AutoValidateServiceNotes   { get; set; }
    public List<CashierInfo> Cashiers      { get; set; } = new();
    public List<ServiceProvider> Agents    { get; set; } = new();
}
