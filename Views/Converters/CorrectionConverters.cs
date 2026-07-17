using System.Globalization;
using System.Windows.Data;
using System.Windows.Media;
using AtolGenerator.Models;

namespace AtolGenerator.Views.Converters;

/// <summary>OrderKind → русское название (для чипа в списке).</summary>
[ValueConversion(typeof(OrderKind), typeof(string))]
public class OrderKindLabelConverter : IValueConverter
{
    public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        => value is OrderKind k ? k switch
        {
            OrderKind.SingleRefund         => "🔧 Возврат прихода",
            OrderKind.SingleCorrection     => "🔧 Коррекция (XML)",
            OrderKind.RefundCorrectionPair => "🔧 Пара: возврат+коррекция (XML)",
            _ => string.Empty,
        } : string.Empty;

    public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
        => throw new NotImplementedException();
}

/// <summary>CorrectionScenario → русское название.</summary>
[ValueConversion(typeof(CorrectionScenario), typeof(string))]
public class CorrectionScenarioLabelConverter : IValueConverter
{
    public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        => value is CorrectionScenario s ? s.ToDisplayString() : string.Empty;

    public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
        => throw new NotImplementedException();
}

/// <summary>OrderKind → цвет полосы слева у карточки.</summary>
[ValueConversion(typeof(OrderKind), typeof(Brush))]
public class OrderKindToBrushConverter : IValueConverter
{
    // Цвета:
    //   Regular              — обычный (прозрачный/полоса не видна)
    //   SingleRefund         — голубой    #4080FF
    //   SingleCorrection     — оранжевый  #FF8800   (требует XML)
    //   RefundCorrectionPair — красно-оранжевый #FF5A1A (пара, XML)
    public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        => value is OrderKind k ? k switch
        {
            OrderKind.SingleRefund         => new SolidColorBrush(Color.FromRgb(0x40, 0x80, 0xFF)),
            OrderKind.SingleCorrection     => new SolidColorBrush(Color.FromRgb(0xFF, 0x88, 0x00)),
            OrderKind.RefundCorrectionPair => new SolidColorBrush(Color.FromRgb(0xFF, 0x5A, 0x1A)),
            _ => Brushes.Transparent,
        } : Brushes.Transparent;

    public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
        => throw new NotImplementedException();
}

/// <summary>OrderKind → видимость (Visible если коррекция).</summary>
[ValueConversion(typeof(OrderKind), typeof(System.Windows.Visibility))]
public class OrderKindToVisibilityConverter : IValueConverter
{
    public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        => value is OrderKind k && k != OrderKind.Regular
            ? System.Windows.Visibility.Visible
            : System.Windows.Visibility.Collapsed;

    public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
        => throw new NotImplementedException();
}

/// <summary>SourceDocumentType → короткая русская подпись (для чипа в карточке).</summary>
[ValueConversion(typeof(SourceDocumentType), typeof(string))]
public class SourceDocumentTypeLabelConverter : IValueConverter
{
    public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        => value is SourceDocumentType t ? t switch
        {
            SourceDocumentType.Realization => "Реализация",
            SourceDocumentType.CardPayment => "Оплата картой",
            SourceDocumentType.CashPayment => "ПКО (нал.)",
            SourceDocumentType.CashExpense => "РКО (расход)",
            SourceDocumentType.BuyerOrder  => "Заказ покупателя",
            SourceDocumentType.KkmCheck    => "Чек ККМ",
            SourceDocumentType.FpOnly      => "Только ФП",
            _ => string.Empty,
        } : string.Empty;

    public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
        => throw new NotImplementedException();
}

/// <summary>SourceDocumentType → Visible если ≠ Unknown (нужен чтобы не показывать пустые чипы).</summary>
[ValueConversion(typeof(SourceDocumentType), typeof(System.Windows.Visibility))]
public class SourceDocumentTypeVisibilityConverter : IValueConverter
{
    public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        => value is SourceDocumentType t && t != SourceDocumentType.Unknown
            ? System.Windows.Visibility.Visible
            : System.Windows.Visibility.Collapsed;

    public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
        => throw new NotImplementedException();
}

/// <summary>CorrectionScenario == Unknown → Visible (для предупреждения «выберите сценарий»).</summary>
[ValueConversion(typeof(CorrectionScenario), typeof(System.Windows.Visibility))]
public class UnknownScenarioVisibilityConverter : IValueConverter
{
    public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        => value is CorrectionScenario s && s == CorrectionScenario.Unknown
            ? System.Windows.Visibility.Visible
            : System.Windows.Visibility.Collapsed;

    public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
        => throw new NotImplementedException();
}

[ValueConversion(typeof(string), typeof(string))]
public class VatRateLabelConverter : IValueConverter
{
    public object Convert(object value, Type targetType, object parameter, CultureInfo culture) =>
        VatRateCatalog.LabelFor(value as string);

    public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture) =>
        throw new NotSupportedException();
}
