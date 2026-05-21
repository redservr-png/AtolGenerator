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
    [JsonPropertyName("uuid")]      public string?         Uuid      { get; set; }
    [JsonPropertyName("error")]     public AtolError?      Error     { get; set; }
    [JsonPropertyName("status")]    public string?         Status    { get; set; }
    [JsonPropertyName("timestamp")] public string?         Timestamp { get; set; }
    [JsonPropertyName("payload")]   public AtolPayload?    Payload   { get; set; }
}

public class AtolPayload
{
    [JsonPropertyName("fiscal_document_number")]     public long?   FiscalDocNumber  { get; set; }  // № ФД
    [JsonPropertyName("fiscal_document_attribute")]  public long?   FiscalSign       { get; set; }  // ФПД
    [JsonPropertyName("fiscal_receipt_number")]      public long?   FiscalReceiptNum { get; set; }
    [JsonPropertyName("shift_number")]               public long?   ShiftNumber      { get; set; }
    [JsonPropertyName("receipt_datetime")]           public string? ReceiptDateTime  { get; set; }
    [JsonPropertyName("total")]                      public double? Total            { get; set; }
    [JsonPropertyName("ofd_receipt_url")]            public string? OfdReceiptUrl    { get; set; }
}

public class AtolError
{
    [JsonPropertyName("code")] public int    Code { get; set; }
    [JsonPropertyName("text")] public string Text { get; set; } = string.Empty;
}

public class AtolPunchResult
{
    public bool   Success         { get; set; }
    public string Uuid            { get; set; } = string.Empty;
    public string Error           { get; set; } = string.Empty;
    public string Status          { get; set; } = string.Empty;
    public long?  FiscalDocNumber { get; set; }  // № ФД
    public long?  FiscalSign      { get; set; }  // ФПД
    public string ReceiptDateTime { get; set; } = string.Empty;
    public string OfdReceiptUrl   { get; set; } = string.Empty;
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

    // JSONL: каждая успешно пробитая операция — одной строкой
    public static string PunchedJsonPath { get; } = Path.Combine(
        AppDomain.CurrentDomain.BaseDirectory, "punched_checks.jsonl");

    private static void Log(string msg)
    {
        try { File.AppendAllText(LogPath, $"[{DateTime.Now:dd.MM.yyyy HH:mm:ss}] {msg}{Environment.NewLine}"); }
        catch { }
    }

    /// <summary>
    /// Дописывает успешно пробитый чек в punched_checks.jsonl (одна строка JSON).
    /// Поля: ts, operation, order_num, realization_num, amount, uuid, fiscal_doc, fiscal_sign, receipt_dt, ofd_url, cashier
    /// </summary>
    private static void LogPunch(
        string operation, string orderNum, string realizationNum, double amount,
        AtolPunchResult result, string cashier)
    {
        if (!result.Success) return;
        try
        {
            var entry = new
            {
                ts               = DateTime.Now.ToString("dd.MM.yyyy HH:mm:ss"),
                operation        = operation,
                order_num        = orderNum,
                realization_num  = realizationNum,
                amount           = amount,
                uuid             = result.Uuid,
                fiscal_doc       = result.FiscalDocNumber,
                fiscal_sign      = result.FiscalSign,
                receipt_dt       = result.ReceiptDateTime,
                ofd_url          = result.OfdReceiptUrl,
                cashier          = cashier,
            };
            var json = JsonSerializer.Serialize(entry,
                new JsonSerializerOptions { Encoder = System.Text.Encodings.Web.JavaScriptEncoder.UnsafeRelaxedJsonEscaping });
            File.AppendAllText(PunchedJsonPath, json + Environment.NewLine);
        }
        catch (Exception ex) { Log($"LogPunch error: {ex.Message}"); }
    }

    // ── Проверка group_code через фиктивный запрос статуса ──────────────────
    /// <summary>
    /// Возвращает (true, detail) если group_code отвечает JSON-ом (значит он валиден),
    /// иначе (false, detail) — пустой ответ = неверный group_code.
    /// </summary>
    public static async Task<(bool ok, string detail)> TestGroupCodeAsync(string groupCode, string token)
    {
        try
        {
            var resp = await Http.GetAsync($"{groupCode}/report/00000000000000000000000000000000?token={token}");
            var body = await resp.Content.ReadAsStringAsync();
            var http = (int)resp.StatusCode;

            if (!string.IsNullOrWhiteSpace(body))
            {
                // Любой JSON-ответ = group_code валиден (АТОЛ вернул ошибку "задание не найдено" — это нормально)
                Log($"TestGroupCode [{groupCode}]: HTTP {http} → {body[..Math.Min(200, body.Length)]}");
                return (true, $"Group Code доступен (HTTP {http})");
            }
            Log($"TestGroupCode [{groupCode}]: HTTP {http} — пустой ответ (неверный Group Code?)");
            return (false, $"Group Code недоступен — пустой ответ сервера (HTTP {http}). Проверьте правильность Group Code.");
        }
        catch (Exception ex)
        {
            Log($"TestGroupCode [{groupCode}]: Exception {ex.Message}");
            return (false, ex.Message);
        }
    }

    // ── Пробить возврат прихода из 1С-таблицы (sell_refund) ─────────────────
    // Примечание: sell_correction через АТОЛ Online API не поддерживается (ошибка 31).
    // Для коррекции используйте кнопку «Сформировать XML» — она создаёт пару файлов.
    public static async Task<AtolPunchResult> PunchCorrectionAsync(
        AtolCredentials creds, OneCRealization r, string cashierName = "")
    {
        var (token, tokenErr) = await GetTokenAsync(creds);
        if (token is null)
            return new AtolPunchResult { Error = $"Нет токена: {tokenErr}" };

        if (string.IsNullOrEmpty(cashierName))
            cashierName = AppConstants.CashierName;

        string vatType;
        double vatSum;
        if (r.IsService)
        {
            // НДС по городу реализации: Страхов → vat5, остальные → none
            vatType = AppConstants.GetServiceVatTypeByCity(r.City);
            vatSum  = vatType == "vat5"
                ? Math.Round(r.Amount * 5.0 / 100.0, 2)
                : r.Amount;
        }
        else
        {
            vatType = "vat122";
            vatSum  = Math.Round(r.Amount * 22.0 / 122.0, 2);
        }

        // sell_refund (возврат прихода) — единственная из пары, которую поддерживает API.
        // sell_correction нужно пробивать вручную через сформированный XML-файл.
        var requestBody = new
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
                        name           = $"{(r.IsService ? "Услуга" : "Товар/услуга")} по реализации {r.DocNumber}",
                        price          = r.Amount,
                        quantity       = 1.0,
                        sum            = r.Amount,
                        payment_method = "full_payment",
                        payment_object = r.IsService ? "service" : "commodity",
                        vat            = new { type = vatType, sum = vatSum },
                    }
                },
                payments = new[] { new { type = 2, sum = r.Amount } },  // 2 = аванс/предоплата
                vats     = new[] { new { type = vatType, sum = vatSum } },
                total    = r.Amount,
                cashier  = cashierName,
                // Тег 1192 — ФП исходного чека (если есть)
                additional_check_props = string.IsNullOrEmpty(r.FiscalNumber) ? null : r.FiscalNumber,
                // Тег 1086 — доп.реквизит пользователя: номер реализации 1С
                additional_user_attribute = string.IsNullOrEmpty(r.DocNumber) ? null : new
                {
                    name  = "Номер реализации",
                    value = r.DocNumber,
                },
            }
        };

        try
        {
            var json    = JsonSerializer.Serialize(requestBody, new JsonSerializerOptions
                          { DefaultIgnoreCondition = System.Text.Json.Serialization.JsonIgnoreCondition.WhenWritingNull });
            var content = new StringContent(json, Encoding.UTF8, "application/json");
            var resp    = await Http.PostAsync($"{creds.GroupCode}/sell_refund?token={token}", content);
            var body    = await resp.Content.ReadAsStringAsync();
            Log($"PunchRefund [{r.DocNumber}]: HTTP {(int)resp.StatusCode} → {(body.Length > 400 ? body[..400] + "…" : body)}");

            var result = JsonSerializer.Deserialize<AtolReceiptResponse>(body, JsonOpts);
            if (result?.Error is { Code: not 0 } err)
            {
                if (err.Code == 3 || err.Code == 4) InvalidateToken();
                return new AtolPunchResult { Error = $"АТОЛ {err.Code}: {err.Text}" };
            }
            if (string.IsNullOrEmpty(result?.Uuid))
                return new AtolPunchResult { Error = $"UUID не получен. Ответ: {body}" };

            var poll = await PollStatusAsync(creds.GroupCode, "sell_refund", result.Uuid, token,
                $"sell_refund/{r.DocNumber}");
            // r.DocNumber — это номер реализации 1С, его и пишем как realization_num
            LogPunch("sell_refund", r.DocNumber, r.DocNumber, r.Amount, poll, cashierName);
            return poll;
        }
        catch (Exception ex)
        {
            Log($"PunchRefund [{r.DocNumber}]: Exception: {ex.Message}");
            return new AtolPunchResult { Error = ex.Message };
        }
    }

    // ── Пробить произвольный чек из списка заказов ────────────────────────────
    // checkType: sell | sell_correction | buy_correction | sell_refund | buy_refund
    // paymentType: cash | card | advance
    public static async Task<AtolPunchResult> PunchOrderAsync(
        AtolCredentials creds, OrderEntry order, string checkType, string paymentType,
        string tab = "payment")
    {
        var (token, tokenErr) = await GetTokenAsync(creds);
        if (token is null)
            return new AtolPunchResult { Error = $"Нет токена: {tokenErr}" };

        // НДС
        string vatType;
        double vatSum;
        if (order.IsService)
        {
            vatType = order.AgentInfo?.VatType ?? "none";
            vatSum  = vatType == "vat5"
                ? Math.Round(order.Amount * 5.0 / 100.0, 2)
                : order.Amount;
        }
        else if (checkType is "buy_correction" or "buy_refund")
        {
            vatType = "none";
            vatSum  = order.Amount;
        }
        else
        {
            vatType = "vat122";
            vatSum  = Math.Round(order.Amount * 22.0 / 122.0, 2);
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
                    // Тег 1086 (additional_user_attribute) запрещён в correction по схеме АТОЛ —
                    // номер реализации идёт только в correction_info.base_number (тег 1179).
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
                            // Оплата (аванс от покупателя) → «Платёж»; Реализация → товар/услуга
                            name           = tab == "payment"
                                ? $"Аванс от покупателя по заказу № {order.OrderNum}"
                                : $"{(order.IsService ? "Услуга" : "Товар")} по заказу {order.OrderNum}",
                            price          = order.Amount,
                            quantity       = 1.0,
                            sum            = order.Amount,
                            // payment_method: оплата → advance/full_prepayment; реализация → full_payment
                            payment_method = tab == "payment"
                                ? (order.IsService ? "full_prepayment" : "advance")
                                : "full_payment",
                            // payment_object: оплата → payment; реализация → service/commodity
                            payment_object = tab == "payment"
                                ? "payment"
                                : (order.IsService ? "service" : "commodity"),
                            vat            = new { type = vatType, sum = vatSum },
                        }
                    },
                    payments = new[] { new { type = payType, sum = order.Amount } },
                    vats     = new[] { new { type = vatType, sum = vatSum } },
                    total    = order.Amount,
                    cashier  = AppConstants.CashierName,
                    // Тег 1086 — доп.реквизит пользователя: номер реализации (для sell_refund).
                    // Для обычного sell — не добавляем (номер уже в названии позиции).
                    additional_user_attribute = checkType == "sell_refund" && !string.IsNullOrEmpty(order.OrderNum)
                        ? new { name = "Номер реализации", value = order.OrderNum }
                        : null,
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

            var poll = await PollStatusAsync(creds.GroupCode, endpoint, result.Uuid, token,
                $"{checkType}/{order.OrderNum}");
            // Для коррекций база (base_number) — номер реализации; для обычных — order_num.
            var realNum = checkType is "sell_correction" or "buy_correction"
                ? (string.IsNullOrEmpty(order.CorrectionNumber) ? order.OrderNum : order.CorrectionNumber)
                : order.OrderNum;
            LogPunch(checkType, order.OrderNum, realNum, order.Amount, poll, AppConstants.CashierName);
            return poll;
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
        string tag = string.IsNullOrEmpty(logLabel) ? uuid[..8] : logLabel;

        for (int i = 0; i < 8; i++)
        {
            await Task.Delay(2500);
            try
            {
                // АТОЛ API v4: статус всегда через /report/{uuid}, не через /{operation}/{uuid}
                var resp   = await Http.GetAsync(
                    $"{groupCode}/report/{uuid}?token={token}");
                var body   = await resp.Content.ReadAsStringAsync();
                if (string.IsNullOrWhiteSpace(body))
                {
                    Log($"Poll [{tag}] attempt {i + 1}: HTTP {(int)resp.StatusCode} — пустой ответ");
                    continue;
                }
                var result = JsonSerializer.Deserialize<AtolReceiptResponse>(body, JsonOpts);

                if (result?.Status == "done")
                {
                    var pl = result.Payload;
                    Log($"Poll [{tag}] attempt {i + 1}: done ✓ ФД={pl?.FiscalDocNumber} ФПД={pl?.FiscalSign}");
                    return new AtolPunchResult
                    {
                        Success         = true,
                        Uuid            = uuid,
                        Status          = "done",
                        FiscalDocNumber = pl?.FiscalDocNumber,
                        FiscalSign      = pl?.FiscalSign,
                        ReceiptDateTime = pl?.ReceiptDateTime ?? string.Empty,
                        OfdReceiptUrl   = pl?.OfdReceiptUrl   ?? string.Empty,
                    };
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
