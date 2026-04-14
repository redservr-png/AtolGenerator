using System.Windows;
using System.Windows.Controls;

namespace AtolGenerator.Views.Converters;

/// <summary>
/// Attached property that enables two-way binding on PasswordBox.Password.
/// Usage: cv:PasswordHelper.Attach="True" cv:PasswordHelper.Password="{Binding MyPassword}"
/// </summary>
public static class PasswordHelper
{
    public static readonly DependencyProperty AttachProperty =
        DependencyProperty.RegisterAttached("Attach", typeof(bool), typeof(PasswordHelper),
            new PropertyMetadata(false, OnAttachChanged));

    public static readonly DependencyProperty PasswordProperty =
        DependencyProperty.RegisterAttached("Password", typeof(string), typeof(PasswordHelper),
            new FrameworkPropertyMetadata(string.Empty, FrameworkPropertyMetadataOptions.BindsTwoWayByDefault, OnPasswordChanged));

    private static readonly DependencyProperty IsUpdatingProperty =
        DependencyProperty.RegisterAttached("IsUpdating", typeof(bool), typeof(PasswordHelper));

    public static void SetAttach(DependencyObject d, bool value)  => d.SetValue(AttachProperty, value);
    public static bool GetAttach(DependencyObject d)              => (bool)d.GetValue(AttachProperty);

    public static void SetPassword(DependencyObject d, string value) => d.SetValue(PasswordProperty, value);
    public static string GetPassword(DependencyObject d)             => (string)d.GetValue(PasswordProperty);

    private static void SetIsUpdating(DependencyObject d, bool value) => d.SetValue(IsUpdatingProperty, value);
    private static bool GetIsUpdating(DependencyObject d)             => (bool)d.GetValue(IsUpdatingProperty);

    private static void OnAttachChanged(DependencyObject d, DependencyPropertyChangedEventArgs e)
    {
        if (d is not PasswordBox pb) return;
        if ((bool)e.OldValue) pb.PasswordChanged -= PasswordChanged;
        if ((bool)e.NewValue) pb.PasswordChanged += PasswordChanged;
    }

    private static void OnPasswordChanged(DependencyObject d, DependencyPropertyChangedEventArgs e)
    {
        if (d is not PasswordBox pb) return;
        if (GetIsUpdating(pb)) return;
        pb.PasswordChanged -= PasswordChanged;
        pb.Password = (string)e.NewValue;
        pb.PasswordChanged += PasswordChanged;
    }

    private static void PasswordChanged(object sender, RoutedEventArgs e)
    {
        if (sender is not PasswordBox pb) return;
        SetIsUpdating(pb, true);
        SetPassword(pb, pb.Password);
        SetIsUpdating(pb, false);
    }
}
