using System.IO;

namespace AtolGenerator.Helpers;

public static class WebView2ErrorHelper
{
    public static string GetStartupMessage(Exception exception)
    {
        if (!HasNativeLoader())
        {
            return "Не найден WebView2Loader.dll. Распакуйте ZIP-архив программы целиком " +
                   "в отдельную папку и запускайте EXE рядом с WebView2Loader.dll. " +
                   "Не запускайте EXE непосредственно из архива и не переносите его отдельно.\n\n" +
                   exception.Message;
        }

        return "Не удалось открыть встроенный браузер. Установите или восстановите " +
               "Microsoft Edge WebView2 Runtime и повторите попытку.\n\n" + exception.Message;
    }

    private static bool HasNativeLoader()
    {
        var baseDirectory = AppContext.BaseDirectory;
        var architecture = Environment.Is64BitProcess ? "win-x64" : "win-x86";
        return File.Exists(Path.Combine(baseDirectory, "WebView2Loader.dll")) ||
               File.Exists(Path.Combine(baseDirectory, "runtimes", architecture, "native", "WebView2Loader.dll"));
    }
}
