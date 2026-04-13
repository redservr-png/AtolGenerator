using AtolGenerator.Models;

namespace AtolGenerator.Constants;

public static class AppConstants
{
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
