using System.Diagnostics;
using System.Net;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text.Json;

namespace AtolGenerator.Services;

public sealed record ApplicationUpdateResult(
    bool CheckSucceeded,
    bool IsUpdateAvailable,
    string CurrentVersion,
    string LatestVersion,
    string ReleaseUrl,
    string DownloadUrl,
    string ErrorMessage);

public static class ApplicationUpdateService
{
    public const string RepositoryUrl = "https://github.com/redservr-png/AtolGenerator";
    public const string ReleasesUrl = RepositoryUrl + "/releases";

    private const string LatestReleaseApiUrl =
        "https://api.github.com/repos/redservr-png/AtolGenerator/releases/latest";

    private static readonly HttpClient HttpClient = new()
    {
        Timeout = TimeSpan.FromSeconds(10),
    };

    public static string CurrentVersionText { get; } = ReadCurrentVersionText();
    public static ApplicationUpdateResult? LastResult { get; private set; }

    public static async Task<ApplicationUpdateResult> CheckForUpdateAsync(
        CancellationToken cancellationToken = default)
    {
        try
        {
            var json = await DownloadLatestReleaseJsonAsync(cancellationToken);
            using var document = JsonDocument.Parse(json);
            var root = document.RootElement;

            var tagName = root.GetProperty("tag_name").GetString()?.Trim() ?? string.Empty;
            var releaseUrl = root.GetProperty("html_url").GetString()?.Trim() ?? ReleasesUrl;
            if (!TryParseVersion(tagName, out var latestVersion) ||
                !TryParseVersion(CurrentVersionText, out var currentVersion))
            {
                return Complete(Failed("GitHub вернул версию в неизвестном формате."));
            }

            var downloadUrl = FindDownloadUrl(root);
            return Complete(new ApplicationUpdateResult(
                true,
                latestVersion > currentVersion,
                CurrentVersionText,
                FormatVersion(latestVersion),
                releaseUrl,
                downloadUrl,
                string.Empty));
        }
        catch (OperationCanceledException) when (!cancellationToken.IsCancellationRequested)
        {
            return Complete(Failed("Сервер обновлений не ответил вовремя."));
        }
        catch (Exception ex)
        {
            return Complete(Failed($"Не удалось проверить обновления: {ex.Message}"));
        }
    }

    public static void OpenUrl(string url)
    {
        if (string.IsNullOrWhiteSpace(url)) return;
        Process.Start(new ProcessStartInfo(url) { UseShellExecute = true });
    }

    private static ApplicationUpdateResult Failed(string message) => new(
        false,
        false,
        CurrentVersionText,
        string.Empty,
        ReleasesUrl,
        string.Empty,
        message);

    private static ApplicationUpdateResult Complete(ApplicationUpdateResult result)
    {
        LastResult = result;
        return result;
    }

    private static async Task<string> DownloadLatestReleaseJsonAsync(CancellationToken cancellationToken)
    {
        try
        {
            using var request = new HttpRequestMessage(HttpMethod.Get, LatestReleaseApiUrl)
            {
                Version = HttpVersion.Version11,
                VersionPolicy = HttpVersionPolicy.RequestVersionExact,
            };
            request.Headers.Accept.Add(new MediaTypeWithQualityHeaderValue("application/vnd.github+json"));
            request.Headers.UserAgent.ParseAdd($"AtolGenerator/{CurrentVersionText}");
            request.Headers.Add("X-GitHub-Api-Version", "2026-03-10");

            using var response = await HttpClient.SendAsync(request, cancellationToken);
            response.EnsureSuccessStatusCode();
            return await response.Content.ReadAsStringAsync(cancellationToken);
        }
        catch (Exception ex) when (
            !cancellationToken.IsCancellationRequested &&
            ex is HttpRequestException or OperationCanceledException)
        {
            return await Task.Run(DownloadLatestReleaseWithWinHttp, cancellationToken);
        }
    }

    private static string DownloadLatestReleaseWithWinHttp()
    {
        Exception? lastError = null;
        for (var attempt = 1; attempt <= 3; attempt++)
        {
            try
            {
                return DownloadLatestReleaseWithWinHttpOnce();
            }
            catch (Exception ex)
            {
                lastError = ex;
                if (attempt < 3) Thread.Sleep(400 * attempt);
            }
        }

        throw new HttpRequestException("Не удалось подключиться к GitHub после трёх попыток.", lastError);
    }

    private static string DownloadLatestReleaseWithWinHttpOnce()
    {
        var requestType = Type.GetTypeFromProgID("WinHttp.WinHttpRequest.5.1")
                          ?? throw new InvalidOperationException("Компонент Windows HTTP недоступен.");
        dynamic? request = Activator.CreateInstance(requestType);
        if (request is null)
            throw new InvalidOperationException("Не удалось создать запрос Windows HTTP.");

        try
        {
            request.SetTimeouts(5000, 5000, 8000, 8000);
            request.Open("GET", LatestReleaseApiUrl, false);
            request.SetRequestHeader("Accept", "application/vnd.github+json");
            request.SetRequestHeader("User-Agent", $"AtolGenerator/{CurrentVersionText}");
            request.SetRequestHeader("X-GitHub-Api-Version", "2026-03-10");
            request.Send();

            var status = (int)request.Status;
            if (status is < 200 or >= 300)
                throw new HttpRequestException($"GitHub вернул HTTP {status}.");
            return (string)request.ResponseText;
        }
        finally
        {
            Marshal.FinalReleaseComObject(request);
        }
    }

    private static string FindDownloadUrl(JsonElement release)
    {
        if (!release.TryGetProperty("assets", out var assets) ||
            assets.ValueKind != JsonValueKind.Array)
            return string.Empty;

        var architecture = Environment.Is64BitProcess ? "x64" : "x86";
        foreach (var asset in assets.EnumerateArray())
        {
            var name = asset.GetProperty("name").GetString() ?? string.Empty;
            if (!name.Contains(architecture, StringComparison.OrdinalIgnoreCase)) continue;
            return asset.GetProperty("browser_download_url").GetString()?.Trim() ?? string.Empty;
        }

        return string.Empty;
    }

    private static string ReadCurrentVersionText()
    {
        var assembly = typeof(ApplicationUpdateService).Assembly;
        var informational = assembly
            .GetCustomAttribute<AssemblyInformationalVersionAttribute>()?
            .InformationalVersion;
        var clean = informational?.Split('+', 2)[0].Trim();
        if (!string.IsNullOrWhiteSpace(clean)) return clean;

        return FormatVersion(NormalizeVersion(assembly.GetName().Version ?? new Version(1, 0)));
    }

    private static bool TryParseVersion(string value, out Version version)
    {
        var clean = value.Trim().TrimStart('v', 'V');
        var suffixIndex = clean.IndexOfAny(['-', '+']);
        if (suffixIndex >= 0) clean = clean[..suffixIndex];

        if (Version.TryParse(clean, out var parsed))
        {
            version = NormalizeVersion(parsed);
            return true;
        }

        version = new Version(0, 0, 0, 0);
        return false;
    }

    private static Version NormalizeVersion(Version version) => new(
        version.Major,
        version.Minor,
        Math.Max(version.Build, 0),
        Math.Max(version.Revision, 0));

    private static string FormatVersion(Version version)
    {
        if (version.Revision > 0) return version.ToString(4);
        if (version.Build > 0) return version.ToString(3);
        return version.ToString(2);
    }
}
