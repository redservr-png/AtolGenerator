using System.Net.Http;
using System.Net.Http.Json;
using System.Text;
using System.Text.Json;
using System.Text.Json.Serialization;

namespace AtolGenerator.Services;

public class AtolCredentials
{
    public string Login    { get; set; } = string.Empty;
    public string Password { get; set; } = string.Empty;
    public string GroupCode { get; set; } = string.Empty;
}

public class AtolTokenResponse
{
    [JsonPropertyName("error")]     public AtolError? Error     { get; set; }
    [JsonPropertyName("token")]     public string?    Token     { get; set; }
    [JsonPropertyName("timestamp")] public string?    Timestamp { get; set; }
}

public class AtolError
{
    [JsonPropertyName("code")]     public int    Code    { get; set; }
    [JsonPropertyName("text")]     public string Text    { get; set; } = string.Empty;
    [JsonPropertyName("type")]     public string Type    { get; set; } = string.Empty;
}

public static class AtolApiService
{
    private static readonly HttpClient Http = new()
    {
        BaseAddress = new Uri("https://online.atol.ru/possystem/v4/"),
        Timeout     = TimeSpan.FromSeconds(30),
    };

    private static readonly JsonSerializerOptions JsonOpts = new()
    {
        PropertyNameCaseInsensitive = true,
    };

    // ── Получить токен ────────────────────────────────────────────────────────
    public static async Task<(string? token, string error)> GetTokenAsync(AtolCredentials c)
    {
        try
        {
            var body    = JsonSerializer.Serialize(new { login = c.Login, pass = c.Password });
            var content = new StringContent(body, Encoding.UTF8, "application/json");
            var resp    = await Http.PostAsync("getToken", content);
            var json    = await resp.Content.ReadAsStringAsync();

            var result  = JsonSerializer.Deserialize<AtolTokenResponse>(json, JsonOpts);
            if (result?.Error is { Code: not 0 } err)
                return (null, $"Ошибка АТОЛ {err.Code}: {err.Text}");
            if (string.IsNullOrEmpty(result?.Token))
                return (null, $"Токен не получен. Ответ: {json}");

            return (result.Token, string.Empty);
        }
        catch (Exception ex)
        {
            return (null, $"Ошибка запроса: {ex.Message}");
        }
    }

    // ── Найти group_code через известные эндпоинты ────────────────────────────
    // АТОЛ не публикует эндпоинт списка групп, но мы можем перебрать варианты.
    public static async Task<string> DiscoverGroupCodeAsync(string token)
    {
        var sb = new StringBuilder();

        // Пробуем /cashRegisters — некоторые версии API возвращают список касс
        var candidates = new[]
        {
            $"cashRegisters?token={token}",
            $"groups?token={token}",
            $"report?token={token}",
            $"?token={token}",
        };

        foreach (var path in candidates)
        {
            try
            {
                var resp = await Http.GetAsync(path);
                var body = await resp.Content.ReadAsStringAsync();
                sb.AppendLine($"GET /{path.Split('?')[0]}  →  HTTP {(int)resp.StatusCode}");
                sb.AppendLine(body.Length > 500 ? body[..500] + "…" : body);
                sb.AppendLine();
            }
            catch (Exception ex)
            {
                sb.AppendLine($"GET /{path.Split('?')[0]}  →  Ошибка: {ex.Message}");
            }
        }

        return sb.ToString();
    }
}
