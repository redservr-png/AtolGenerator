using System.Windows;
using AtolGenerator.Services;

namespace AtolGenerator;

public partial class App : Application
{
    protected override void OnStartup(StartupEventArgs e)
    {
        ThemeService.ApplyTheme(ApplicationSettingsStore.Current.ThemeKey);
        base.OnStartup(e);
    }
}
