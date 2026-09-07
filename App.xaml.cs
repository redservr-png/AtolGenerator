using System.IO;
using System.Windows;
using System.Windows.Threading;
using AtolGenerator.Services;

namespace AtolGenerator;

public partial class App : Application
{
    protected override void OnStartup(StartupEventArgs e)
    {
        DispatcherUnhandledException += OnDispatcherUnhandledException;
        AppDomain.CurrentDomain.UnhandledException += OnDomainUnhandledException;
        ThemeService.ApplyTheme(ApplicationSettingsStore.Current.ThemeKey);
        base.OnStartup(e);
    }

    private static void OnDispatcherUnhandledException(object sender, DispatcherUnhandledExceptionEventArgs e)
    {
        LogCrash(e.Exception);
        e.Handled = true;
        MessageBox.Show(
            e.Exception.Message,
            "Ошибка",
            MessageBoxButton.OK,
            MessageBoxImage.Error);
    }

    private static void OnDomainUnhandledException(object sender, UnhandledExceptionEventArgs e)
    {
        if (e.ExceptionObject is Exception ex)
            LogCrash(ex);
    }

    private static void LogCrash(Exception ex)
    {
        try
        {
            var path = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "crash.log");
            File.AppendAllText(path,
                $"[{DateTime.Now:dd.MM.yyyy HH:mm:ss}] {ex}{Environment.NewLine}");
        }
        catch
        {
            /* логирование не должно ронять приложение */
        }
    }
}
