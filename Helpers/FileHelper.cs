using System.Diagnostics;
using System.IO;
using System.Text.RegularExpressions;

namespace AtolGenerator.Helpers;

public static class FileHelper
{
    private static readonly Regex SafePattern = new(@"[^\w\-]", RegexOptions.Compiled);

    public static string SafeFilename(string s) => SafePattern.Replace(s, "_");

    public static void OpenFolder(string path)
    {
        Directory.CreateDirectory(path);
        Process.Start(new ProcessStartInfo
        {
            FileName        = path,
            UseShellExecute = true
        });
    }

    public static string OutputDir =>
        Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "atol_output");
}
