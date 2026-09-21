using System.Globalization;
using System.Windows.Data;

namespace AtolGenerator.Views.Converters;

/// <summary>
/// Деньги в полях ввода: принимает и запятую, и точку.
/// При незавершённом вводе (например «756,») не сбрасывает текст.
/// </summary>
[ValueConversion(typeof(double), typeof(string))]
[ValueConversion(typeof(double?), typeof(string))]
public class MoneyRuConverter : IValueConverter
{
    private static readonly CultureInfo Ru = CultureInfo.GetCultureInfo("ru-RU");

    public object Convert(object? value, Type targetType, object? parameter, CultureInfo culture)
    {
        if (value is null || value is string { Length: 0 })
            return string.Empty;

        if (value is double d)
            return Format(d);

        if (value is float f)
            return Format(f);

        if (value is decimal m)
            return Format((double)m);

        return value.ToString() ?? string.Empty;
    }

    public object ConvertBack(object? value, Type targetType, object? parameter, CultureInfo culture)
    {
        var raw = (value as string)?.Trim() ?? string.Empty;
        if (raw.Length == 0)
        {
            if (Nullable.GetUnderlyingType(targetType) is not null || !targetType.IsValueType)
                return null!;
            return Binding.DoNothing;
        }

        var normalized = raw.Replace('\u00A0', ' ').Replace(" ", "").Replace(',', '.');
        if (normalized is "." or "-" or "-.")
            return Binding.DoNothing;

        if (!double.TryParse(normalized, NumberStyles.Float, CultureInfo.InvariantCulture, out var parsed))
            return Binding.DoNothing;

        if (Nullable.GetUnderlyingType(targetType) == typeof(double) || targetType == typeof(double?))
            return (double?)parsed;

        return parsed;
    }

    private static string Format(double value)
    {
        // Без лишних нулей, но с копейками когда они есть: 756,5 / 756,50 → 756,5
        return value.ToString("0.##", Ru);
    }
}
