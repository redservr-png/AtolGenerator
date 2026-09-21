using System.Windows;

namespace AtolGenerator.Helpers;

/// <summary>
/// Подгоняет окно под рабочую область монитора, чтобы большие диалоги
/// не вылезали за край на ноутбуках и не прыгали при появлении панелей.
/// </summary>
internal static class WindowFit
{
    public static void ToWorkArea(Window window)
    {
        if (window.WindowState == WindowState.Maximized)
            return;

        var work = SystemParameters.WorkArea;
        if (work.Width < 1 || work.Height < 1)
            return;

        window.MaxWidth = work.Width;
        window.MaxHeight = work.Height;

        var minW = window.MinWidth > 0 && !double.IsNaN(window.MinWidth)
            ? Math.Min(window.MinWidth, work.Width)
            : 0;
        var minH = window.MinHeight > 0 && !double.IsNaN(window.MinHeight)
            ? Math.Min(window.MinHeight, work.Height)
            : 0;
        if (minW > 0)
            window.MinWidth = minW;
        if (minH > 0)
            window.MinHeight = minH;

        var width = double.IsNaN(window.Width) ? window.ActualWidth : window.Width;
        var height = double.IsNaN(window.Height) ? window.ActualHeight : window.Height;
        if (width < 1) width = minW > 0 ? minW : 800;
        if (height < 1) height = minH > 0 ? minH : 560;

        var maxW = Math.Min(work.Width, Math.Max(window.MinWidth, work.Width * 0.96));
        var maxH = Math.Min(work.Height, Math.Max(window.MinHeight, work.Height * 0.96));
        width = Math.Clamp(width, window.MinWidth, maxW);
        height = Math.Clamp(height, window.MinHeight, maxH);

        window.Width = width;
        window.Height = height;

        if (window.WindowStartupLocation == WindowStartupLocation.Manual)
        {
            Clamp(window, work);
            return;
        }

        if (window.Owner is { } owner)
        {
            window.Left = owner.Left + (owner.ActualWidth - width) / 2;
            window.Top = owner.Top + (owner.ActualHeight - height) / 2;
        }
        else
        {
            window.Left = work.Left + (work.Width - width) / 2;
            window.Top = work.Top + (work.Height - height) / 2;
        }

        Clamp(window, work);
    }

    private static void Clamp(Window window, Rect work)
    {
        if (window.Left < work.Left)
            window.Left = work.Left;
        if (window.Top < work.Top)
            window.Top = work.Top;
        if (window.Left + window.Width > work.Right)
            window.Left = Math.Max(work.Left, work.Right - window.Width);
        if (window.Top + window.Height > work.Bottom)
            window.Top = Math.Max(work.Top, work.Bottom - window.Height);
    }
}
