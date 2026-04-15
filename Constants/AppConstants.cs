using AtolGenerator.Models;

namespace AtolGenerator.Constants;

// ── Данные кассира ────────────────────────────────────────────────────────────
public class CashierInfo
{
    public string FullName     { get; init; } = string.Empty;  // в XML и предпросмотре
    public string ShortName    { get; init; } = string.Empty;  // "Фамилия И.О." — подпись
    public string Position     { get; init; } = string.Empty;  // "от Должности" — DOCX шапка
    public string NameGenitive { get; init; } = string.Empty;  // "Фамилии И.О." — DOCX шапка
    public string Display      { get; init; } = string.Empty;  // для выпадающего списка

    public override string ToString() => Display;
}

// ── Данные кассы (ККТ) ────────────────────────────────────────────────────────
public class KktData
{
    public string Model  { get; init; } = string.Empty;
    public string Serial { get; init; } = string.Empty;
    public string RegNum { get; init; } = string.Empty;
    public string Ffd    { get; init; } = "1.05";
}

public static class AppConstants
{
    // ── Кассиры ──────────────────────────────────────────────────────────────
    public static readonly IReadOnlyList<CashierInfo> Cashiers = new List<CashierInfo>
    {
        new()
        {
            FullName     = "Консультант-аналитик 1С Полюшков Константин Николаевич",
            ShortName    = "Полюшков К.Н.",
            Position     = "Консультанта-Аналитика 1С",
            NameGenitive = "Полюшкова Константина Николаевича",
            Display      = "Полюшков К.Н. — Консультант-аналитик 1С",
        },
        new()
        {
            FullName     = "Ст.менеджер отдела лидогенерации Бородкина Татьяна Александровна",
            ShortName    = "Бородкина Т.А.",
            Position     = "Ст.менеджера отдела лидогенерации",
            NameGenitive = "Бородкиной Татьяны Александровны",
            Display      = "Бородкина Т.А. — Ст.менеджер отдела лидогенерации",
        },
    };

    public static CashierInfo DefaultCashier => Cashiers[0];

    // Обратная совместимость
    public const string CashierName    = "Консультант-аналитик 1С Полюшков Константин Николаевич";
    public const string CashierShort   = "Полюшков К.Н.";
    public const string PhoneBuyer     = "+79005005016";
    public const string PaymentAddress = "https://alleyadoma.ru";
    public const string InnOrg         = "352526561274";
    public const string Sno            = "osn";
    public const string EmailOrg       = "info@alleyadoma.ru";
    public const string KktModel       = "АТОЛ 42ФС";
    public const string KktSerial      = "00107700988750";
    public const string KktReg         = "0002978147064584";
    public const string FfdVersion     = "1.05";
    public const string City           = "г. Вологда";
    public const string FromPosition   = "Консультанта-Аналитика 1С";
    public const string FromName       = "Полюшкова Константина Николаевича";
    public const string RecipientName  = "ИП Шевелеву Е.Н.";

    public static readonly IReadOnlyList<ServiceProvider> ServiceProviders = new List<ServiceProvider>
    {
        new("Доставка", "Грязовец",     "ИП Страхов Дмитрий Михайлович", "463224431946", "+79992601825"),
        new("Сборка",   "Иваново",      "Воронков Р.А.",                  "370260964433", "+79611189433"),
        new("Доставка", "Иваново",      "ИП Страхов Дмитрий Михайлович", "463224431946", "+79992601825"),
        new("Сборка",   "Ярославль",    "ИП Страхов Дмитрий Михайлович", "463224431946", "+79992601826"),
        new("Доставка", "Ярославль",    "ИП Страхов Дмитрий Михайлович", "463224431946", "+79992601827"),
        new("Сборка",   "Череповец",    "ИП Страхов Дмитрий Михайлович", "463224431946", "+79992601828"),
        new("Доставка", "Череповец",    "ИП Страхов Дмитрий Михайлович", "463224431946", "+79992601829"),
        new("Сборка",   "Архангельск",  "ИП Страхов Дмитрий Михайлович", "463224431946", "+79992601830"),
        new("Доставка", "Архангельск",  "ИП Драчев А.А.",                 "290132766549", "+79009187847"),
        new("Сборка",   "Розница МИКС", "ИП Страхов Дмитрий Михайлович", "463224431946", "+79992601830"),
        new("Доставка", "Розница МИКС", "ИП Страхов Дмитрий Михайлович", "463224431946", "+79992601831"),
        new("Сборка",   "Котлас",       "Анкудинов И.Н.",                 "290402783932", "+79212929739"),
        new("Доставка", "Котлас",       "Анкудинов И.Н.",                 "290402783932", "+79212929739"),
    };

    // ── ККТ по умолчанию (Интернет-магазин / Вологда) ───────────────────────
    public static KktData DefaultKkt => new()
    {
        Model  = KktModel,
        Serial = KktSerial,
        RegNum = KktReg,
        Ffd    = FfdVersion,
    };

    // ── Словарь: подразделение 1С → ККТ ─────────────────────────────────────
    // Ключи совпадают с Подразделение.Наименование из УТ 10.3 (без учёта регистра).
    // Для подразделений без записи используется DefaultKkt.
    private static readonly Dictionary<string, KktData> _cityKkt =
        new(StringComparer.OrdinalIgnoreCase)
    {
        // ── Крупные точки ──────────────────────────────────────────────────
        ["Архангельск"]          = new() { Model = "АТОЛ 11Ф", Serial = "00106708870749", RegNum = "0009526416044505", Ffd = "1.05" },
        ["Вологда Микс"]         = new() { Model = "АТОЛ 11Ф", Serial = "00106708313058", RegNum = "0000322392053329", Ffd = "1.05" },
        ["Вологда МИКС"]         = new() { Model = "АТОЛ 11Ф", Serial = "00106708313058", RegNum = "0000322392053329", Ffd = "1.05" },
        ["Розница МИКС"]         = new() { Model = "АТОЛ 11Ф", Serial = "00106708313058", RegNum = "0000322392053329", Ffd = "1.05" },
        ["Вологда ЛенТупик"]     = new() { Model = "АТОЛ 11Ф", Serial = "00106701831258", RegNum = "0006496061042658", Ffd = "1.05" },
        ["Вологда Луначарского"] = new() { Model = "АТОЛ 11Ф", Serial = "00106701831258", RegNum = "0006496061042658", Ffd = "1.05" },
        ["Вологда Тест"]         = new() { Model = "АТОЛ 11Ф", Serial = "00106708533813", RegNum = "0000574053033118", Ffd = "1.05" },
        ["Иваново"]              = new() { Model = "АТОЛ 11Ф", Serial = "00106708512105", RegNum = "0009564908033715", Ffd = "1.05" },
        ["Котлас"]               = new() { Model = "АТОЛ 11Ф", Serial = "00106700950507", RegNum = "0009564841005962", Ffd = "1.05" },
        ["Череповец"]            = new() { Model = "АТОЛ 11Ф", Serial = "00106705660844", RegNum = "0009526331003083", Ffd = "1.05" },
        ["Череповец Гипер"]      = new() { Model = "АТОЛ 11Ф", Serial = "00106705660844", RegNum = "0009526331003083", Ffd = "1.05" },
        ["Ярославль"]            = new() { Model = "АТОЛ 11Ф", Serial = "00106707811776", RegNum = "0008771350025876", Ffd = "1.05" },
        ["Ярославль Гипер"]      = new() { Model = "АТОЛ 11Ф", Serial = "00106707811776", RegNum = "0008771350025876", Ffd = "1.05" },
        ["Грязовец"]             = new() { Model = "АТОЛ 11Ф", Serial = "00106708873102", RegNum = "0009564926007298", Ffd = "1.05" },
        // ── Малые точки ───────────────────────────────────────────────────
        ["С.им.Бабушкина"]       = new() { Model = "АТОЛ 11Ф", Serial = "00106700052263", RegNum = "0008771337000614", Ffd = "1.05" },
        ["Белозерск"]            = new() { Model = "АТОЛ 11Ф", Serial = "00106708554456", RegNum = "0009564856036931", Ffd = "1.05" },
        ["Кинешма"]              = new() { Model = "АТОЛ 11Ф", Serial = "00106702358698", RegNum = "0009564900056432", Ffd = "1.05" },
        ["Коряжма"]              = new() { Model = "АТОЛ 11Ф", Serial = "00106704396748", RegNum = "0008957538044537", Ffd = "1.05" },
        ["Тейково"]              = new() { Model = "АТОЛ 11Ф", Serial = "00106702014031", RegNum = "0007231086016208", Ffd = "1.05" },
        ["Тутаев"]               = new() { Model = "АТОЛ 11Ф", Serial = "00106701397667", RegNum = "0008270180064231", Ffd = "1.05" },
        ["Фурманов"]             = new() { Model = "АТОЛ 11Ф", Serial = "00106707872414", RegNum = "0008270159058110", Ffd = "1.05" },
        ["Няндома"]              = new() { Model = "АТОЛ 11Ф", Serial = "00106704040242", RegNum = "0008270147047101", Ffd = "1.05" },
        ["Сокол"]                = new() { Model = "АТОЛ 11Ф", Serial = "00106700173256", RegNum = "0008224938055136", Ffd = "1.05" },
        ["Тотьма"]               = new() { Model = "АТОЛ 11Ф", Serial = "00106708451321", RegNum = "0008224930024067", Ffd = "1.05" },
    };

    /// <summary>
    /// Возвращает ККТ по названию города (подразделения 1С).
    /// Поиск: точное совпадение без учёта регистра → частичное вхождение → DefaultKkt.
    /// </summary>
    public static KktData GetKktByCity(string city)
    {
        if (string.IsNullOrWhiteSpace(city)) return DefaultKkt;

        // 1. Точное совпадение (case-insensitive)
        if (_cityKkt.TryGetValue(city.Trim(), out var kkt)) return kkt;

        // 2. Частичное: ключ содержится в названии или наоборот
        var norm = city.Trim();
        foreach (var (key, data) in _cityKkt)
        {
            if (norm.Contains(key, StringComparison.OrdinalIgnoreCase) ||
                key.Contains(norm,  StringComparison.OrdinalIgnoreCase))
                return data;
        }

        return DefaultKkt;
    }

    public static readonly IReadOnlyDictionary<string, string> OperationDescriptions =
        new Dictionary<string, string>
        {
            ["sell"]             = "поступлении оплаты от покупателя",
            ["sell_refund"]      = "возврате денежных средств покупателю",
            ["buy"]              = "закупке у поставщика",
            ["buy_refund"]       = "возврате товара поставщику",
            ["sell_correction"]  = "коррекции прихода",
            ["buy_correction"]   = "коррекции расхода",
        };

    public static readonly IReadOnlyDictionary<string, string> CorrectionDescriptions =
        new Dictionary<string, string>
        {
            ["sell_correction"] = "чек коррекции с признаком расчёта \"Приход\"",
            ["buy_correction"]  = "чек коррекции с признаком расчёта \"Расход\"",
        };
}
