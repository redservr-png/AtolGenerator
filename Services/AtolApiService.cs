using System.IO;
using System.Net.Http;
using System.Text;
using System.Text.Json;
using System.Text.Json.Serialization;
using AtolGenerator.Constants;
using AtolGenerator.Models;

namespace AtolGenerator.Services;

// ── Настройки подключения ─────────────────────────────────────────────────────

public class AtolCredentials
{
    public string Login     { get; set; } = string.Empty;
    public string Password  { get; set; } = string.Empty;
    public string GroupCode { get; set; } = string.Empty;

    private static string SettingsPath => Path.Combine(
        AppDomain.CurrentDomain.BaseDirectory, "atol_settings.json");

    public void Save()
    {
        var json = JsonSerializer.Serialize(this, new JsonSerializerOptions { WriteIndented = true });
        File.WriteAllText(SettingsPath, json);
    }

    public static AtolCredentials Load()
    {
        try
        {
            if (!File.Exists(SettingsPath)) return new();
            var json = File.ReadAllText(SettingsPath);
            return JsonSerializer.Deserialize<AtolCredentials>(json) ?? new();
        }
        catch { return new(); }
    }
}

// ── DTO ───────────────────────────────────────────────────────────────────────

public class AtolTokenResponse
{
    [JsonPropertyName("error")]     public AtolError? Error     { get; set; }
    [JsonPropertyName("token")]     public string?    Token     { get; set; }
    [JsonPropertyName("timestamp")] public string?    Timestamp { get; set; }
}

public class AtolReceiptResponse
{
    [JsonPropertyName("uuid")]      public string?    Uuid      { get; set; }
    [JsonPropertyName("error")]     public AtolError? Error     { get; set; }
    [JsonPropertyName("status")]    public string?    Status    { get; set; }
    [JsonPropertyName("timestamp")] public string?    Timestamp { get; set; }
}

public class AtolError
{
    [JsonPropertyName("code")] public int    Code { get; set; }
    [JsonPropertyName("text")] public string Text { get; set; } = string.Empty;
}

public class AtolPunchResult
{
    public bool   Success { get; set; }
    public string Uuid    { get; set; } = string.Empty;
    public string Error   { get; set; } = string.Empty;
    public string Status  { get; set; } = string.Empty;
}

// ── Сервис ────────────────────────────────────────────────────────────────────

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

    // ── Токен (кэшируем до истечения ~24ч) ───────────────────────────────────
    private static string? _cachedToken;
    private static DateTime _tokenExpiry = DateTime.MinValue;

    public static async Task<(string? token, string error)> GetTokenAsync(AtolCredentials c)
    {
        if (_cachedToken is not null && DateTime.Now < _tokenExpiry)
            return (_cachedToken, string.Empty);

        try
        {
            var body    = JsonSerializer.Serialize(new { login = c.Login, pass = c.Password });
            var content = new StringContent(body, Encoding.UTF8, "application/json");
            var resp    = await Http.PostAsync("getToken", content);
            var json    = await resp.Content.ReadAsStringAsync();
            var result  = JsonSerializer.Deserialize<AtolTokenResponse>(json, JsonOpts);

            if (result?.Error is { Code: not 0 } err)
                return (null, $"Ошибка {err.Code}: {err.Text}");

            if (string.IsNullOrEmpty(result?.Token))
                return (null, $"Токен не получен. Ответ: {json}");

            _cachedToken = result.Token;
            _tokenExpiry = DateTime.Now.AddHours(23);
            return (_cachedToken, string.Empty);
        }
        catch (Exception ex) { return (null, ex.Message); }
    }

    public static void InvalidateToken() => _cachedToken = null;

    // ── Лог-файл ─────────────────────────────────────────────────────────────
    public static string LogPath { get; } = Path.Combine(
        AppDomain.CurrentDomain.BaseDirectory, "atol_log.txt");

    private static void Log(string msg)
    {
        try { File.AppendAllText(LogPath, $"[{DateTime.Now:dd.MM.yyyy HH:mm:ss}] {msg}{Environment.NewLine}"); }
        catch { }
    }

    // ── Пробить чек коррекции (sell-correction) — из 1С-таблицы ─────────────
    public static async Task<AtolPunchResult> PunchCorrectionAsync(
        AtolCredentials creds, OneCRealization r)
    {
        // 1. Получаем токен
        var (token, tokenErr) = await GetTokenAsync(creds);
        if (token is null)
            return new AtolPunchResult { Error = $"Нет токена: {tokenErr}" };

        // 2. Считаем НДС
        string vatType;
        double vatSum;
        if (r.IsService)
        {
            vatType = "none";
            vatSum  = r.Amount;
        }
        else
        {
            vatType = "vat22";
            vatSum  = Math.Round(r.Amount * 22.0 / 100.0, 2);
        }

        // 3. Формируем тело запроса
        var requestBody = new
        {
            timestamp   = DateTime.Now.ToString("dd.MM.yyyy HH:mm:ss"),
            external_id = Guid.NewGuid().ToString("N"),
            correction  = new
            {
                company = new
                {
                    email           = AppConstants.EmailOrg,
                    sno             = AppConstants.Sno,
                    inn             = AppConstants.InnOrg,
                    payment_address = AppConstants.PaymentAddress,
                },
                correction_info = new
                {
                    type        = "self",
                    base_date   = r.DocDate,
                    base_number = r.DocNumber,
                },
                payments = new[] { new { type = 2, sum = r.Amount } },
                vats     = new[] { new { type = vatType, sum = vatSum } },
                cashier  = AppConstants.CashierName,
            }
        };

        // 4. Отправляем
        try
        {
            var json    = JsonSerializer.Serialize(requestBody);
            var content = new StringContent(json, Encoding.UTF8, "application/json");
            var resp    = await Http.PostAsync(
                $"{creds.GroupCode}/sell-correction?token={token}", content);
            var body    = await resp.Content.ReadAsStringAsync();
            var result  = JsonSerializer.Deserialize<AtolReceiptResponse>(body, JsonOpts);

            Log($"PunchCorrection [{r.DocNumber}]: HTTP {(int)resp.StatusCode} → {(body.Length > 400 ? body[..400] + "…" : body)}");

            if (result?.Error is { Code: not 0 } err)
            {
                // Токен мог устареть — сбрасываем
                if (err.Code == 3 || err.Code == 4) InvalidateToken();
                return new AtolPunchResult { Error = $"АТОЛ {err.Code}: {err.Text}" };
            }

            if (string.IsNullOrEmpty(result?.Uuid))
                return new AtolPunchResult { Error = $"UUID не получен. Ответ: {body}" };

            // 5. Ждём финального статуса (до 10 секунд)
            return await PollStatusAsync(creds.GroupCode, "sell_correction", result.Uuid, token,
                $"sell_correction/{r.DocNumber}");
        }
        catch (Exception ex)
        {
            Log($"PunchCorrection [{r.DocNumber}]: Exception: {ex.Message}");
            return new AtolPunchResult { Error = ex.Message };
        }
    }

    // ── Пробить произвольный чек из списка заказов ────────────────────────────
    // checkType: sell | sell_correction | buy_correction | sell_refund | buy_refund
    // paymentType: cash | card | advance
    public static async Task<AtolPunchResult> PunchOrderAsync(
        AtolCredentials creds, OrderEntry order, string checkType, string paymentType)
    {
        var (token, tokenErr) = await GetTokenAsync(creds);
        if (token is null)
            return new AtolPunchResult { Error = $"Нет токена: {tokenErr}" };

        // НДС
        string vatType;
        double vatSum;
        if (order.IsService || checkType is "buy_correction" or "buy_refund")
        {
            vatType = "none";
            vatSum  = order.Amount;
        }
        else
        {
            vatType = "vat22";
            vatSum  = Math.Round(order.Amount * 22.0 / 100.0, 2);
        }

        // Тип оплаты: 0=наличные, 1=безналичные, 2=аванс
        int payType = paymentType switch
        {
            "cash"    => 0,
            "advance" => 2,
            _         => 1,   // card по умолчанию
        };

        string   endpoint;
        object   requestBody;

        if (checkType is "sell_correction" or "buy_correction")
        {
            endpoint = checkType;   // sell_correction / buy_correction (подчёркивание, как sell_refund)

            var baseDate   = !string.IsNullOrEmpty(order.CorrectionDate)
                                 ? order.CorrectionDate
                                 : order.OrderDate.Split(' ')[0];   // берём только дату
            var baseNumber = !string.IsNullOrEmpty(order.CorrectionNumber)
                                 ? order.CorrectionNumber
                                 : order.OrderNum;

            requestBody = new
            {
                timestamp   = DateTime.Now.ToString("dd.MM.yyyy HH:mm:ss"),
                external_id = Guid.NewGuid().ToString("N"),
                correction  = new
                {
                    company = new
                    {
                        email           = AppConstants.EmailOrg,
                        sno             = AppConstants.Sno,
                        inn             = AppConstants.InnOrg,
                        payment_address = AppConstants.PaymentAddress,
                    },
                    correction_info = new
                    {
                        type        = "self",
                        base_date   = baseDate,
                        base_number = baseNumber,
                    },
                    payments = new[] { new { type = payType, sum = order.Amount } },
                    vats     = new[] { new { type = vatType, sum = vatSum } },
                    cashier  = AppConstants.CashierName,
                }
            };
        }
        else
        {
            // sell, sell_refund, buy_refund
            endpoint = checkType;   // URL совпадает с именем операции

            requestBody = new
            {
                timestamp   = DateTime.Now.ToString("dd.MM.yyyy HH:mm:ss"),
                external_id = Guid.NewGuid().ToString("N"),
                receipt     = new
                {
                    client  = new { email = AppConstants.EmailOrg },
                    company = new
                    {
                        email           = AppConstants.EmailOrg,
                        sno             = AppConstants.Sno,
                        inn             = AppConstants.InnOrg,
                        payment_address = AppConstants.PaymentAddress,
                    },
                    items = new[]
                    {
                        new
                        {
                            name           = $"{(order.IsService ? "Услуга" : "Товар")} по заказу {order.OrderNum}",
                            price          = order.Amount,
                            quantity       = 1.0,
                            sum            = order.Amount,
                            payment_method = "full_payment",
                            payment_object = order.IsService ? "service" : "commodity",
                            vat            = new { type = vatType, sum = vatSum },
                        }
                    },
                    payments = new[] { new { type = payType, sum = order.Amount } },
                    vats     = new[] { new { type = vatType, sum = vatSum } },
                    total    = order.Amount,
                    cashier  = AppConstants.CashierName,
                }
            };
        }

        try
        {
            var json    = JsonSerializer.Serialize(requestBody);
            var content = new StringContent(json, Encoding.UTF8, "application/json");
            var resp    = await Http.PostAsync($"{creds.GroupCode}/{endpoint}?token={token}", content);
            var body    = await resp.Content.ReadAsStringAsync();
            Log($"PunchOrder [{checkType}] {order.OrderNum}: HTTP {(int)resp.StatusCode} → {(body.Length > 400 ? body[..400] + "…" : body)}");

            var result = JsonSerializer.Deserialize<AtolReceiptResponse>(body, JsonOpts);

            if (result?.Error is { Code: not 0 } err)
            {
                if (err.Code == 3 || err.Code == 4) InvalidateToken();
                return new AtolPunchResult { Error = $"АТОЛ {err.Code}: {err.Text}" };
            }

            if (string.IsNullOrEmpty(result?.Uuid))
                return new AtolPunchResult { Error = $"UUID не получен. Ответ: {body}" };

            return await PollStatusAsync(creds.GroupCode, endpoint, result.Uuid, token,
                $"{checkType}/{order.OrderNum}");
        }
        catch (Exception ex)
        {
            Log($"PunchOrder [{checkType}] {order.OrderNum}: Exception: {ex.Message}");
            return new AtolPunchResult { Error = ex.Message };
        }
    }

    // ── Опрос статуса ─────────────────────────────────────────────────────────
    private static async Task<AtolPunchResult> PollStatusAsync(
        string groupCode, string operation, string uuid, string token,
        string logLabel = "")
    {
        string tag     = string.IsNullOrEmpty(logLabel) ? uuid[..8] : logLabel;
        string pollOp  = operation.Replace('_', '-');   // АТОЛ: GET использует дефис (sell-refund), POST принимает оба варианта

        for (int i = 0; i < 8; i++)
        {
            await Task.Delay(2500);
            try
            {
                var resp   = await Http.GetAsync(
                    $"{groupCode}/{pollOp}/{uuid}?token={token}");
                var body   = await resp.Content.ReadAsStringAsync();
                if (string.IsNullOrWhiteSpace(body))
                {
                    Log($"Poll [{tag}] attempt {i + 1}: HTTP {(int)resp.StatusCode} — пустой ответ");
                    continue;
                }
                var result = JsonSerializer.Deserialize<AtolReceiptResponse>(body, JsonOpts);

                if (result?.Status == "done")
                {
                    Log($"Poll [{tag}] attempt {i + 1}: done ✓");
                    return new AtolPunchResult { Success = true, Uuid = uuid, Status = "done" };
                }

                if (result?.Status == "fail")
                {
                    var errText = result.Error is { } e ? $"{e.Code}: {e.Text}" : body;
                    Log($"Poll [{tag}] attempt {i + 1}: fail → {errText}");
                    return new AtolPunchResult { Error = errText, Uuid = uuid, Status = "fail" };
                }
                // "wait" — продолжаем ждать
                Log($"Poll [{tag}] attempt {i + 1}: wait…");
            }
            catch (Exception ex) { Log($"Poll [{tag}] attempt {i + 1}: exception HTTP→{ex.Message[..Math.Min(120,ex.Message.Length)]}"); }
        }
        // Статус не получен после 20 секунд — возвращаем предупреждение
        Log($"Poll [{tag}]: timeout — uuid={uuid}, проверьте статус вручную");
        return new AtolPunchResult { Success = false, Uuid = uuid, Status = "wait",
            Error = $"Статус не получен за 20 сек. UUID: {uuid}" };
    }

    // ── Поиск group_code (диагностика) ────────────────────────────────────────
    public static async Task<string> DiscoverGroupCodeAsync(string token)
    {
        var sb = new StringBuilder();
        var candidates = new[] { "cashRegisters", "groups", "report", "" };
        foreach (var path in candidates)
        {
            try
            {
                var url    = string.IsNullOrEmpty(path)
                    ? $"?token={token}" : $"{path}?token={token}";
                var resp   = await Http.GetAsync(url);
                var body   = await resp.Content.ReadAsStringAsync();
                sb.AppendLine($"GET /{path}  →  HTTP {(int)resp.StatusCode}");
                sb.AppendLine(body.Length > 600 ? body[..600] + "…" : body);
                sb.AppendLine();
            }
            catch (Exception ex) { sb.AppendLine($"GET /{path}  →  {ex.Message}"); }
        }
        return sb.ToString();
    }
}
