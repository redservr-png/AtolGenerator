using System.Windows;
using System.Windows.Media;

namespace AtolGenerator.Services;

public static class ThemeService
{
    private static readonly IReadOnlyDictionary<string, string> BrushKeys =
        new Dictionary<string, string>
        {
            ["ThemeBgColor"] = "BrushBg",
            ["ThemeSurface1Color"] = "BrushSurface1",
            ["ThemeSurface2Color"] = "BrushSurface2",
            ["ThemeSurface3Color"] = "BrushSurface3",
            ["ThemeBorderColor"] = "BrushBorder",
            ["ThemeAccentColor"] = "BrushAccent",
            ["ThemeAccent2Color"] = "BrushAccent2",
            ["ThemeTextColor"] = "BrushText",
            ["ThemeText2Color"] = "BrushText2",
            ["ThemeText3Color"] = "BrushText3",
            ["ThemeSidebarBgColor"] = "BrushSidebarBg",
            ["ThemeSidebarBorderColor"] = "BrushSidebarBorder",
            ["ThemeSidebarTextColor"] = "BrushSidebarText",
            ["ThemeSidebarMutedColor"] = "BrushSidebarMuted",
            ["ThemeSidebarStatusColor"] = "BrushSidebarStatus",
            ["ThemeSidebarSeparatorColor"] = "BrushSidebarSeparator",
            ["ThemeSidebarHoverColor"] = "BrushSidebarHover",
        };

    private static readonly IReadOnlyDictionary<string, IReadOnlyDictionary<string, string>> Themes =
        new Dictionary<string, IReadOnlyDictionary<string, string>>(StringComparer.OrdinalIgnoreCase)
        {
            ["light"] = Palette(
                "#F4F7FA", "#FFFFFF", "#EEF3F8", "#E3EBF4", "#D4DEE9",
                "#2563EB", "#0F766E", "#111827", "#4B5563", "#7C8796",
                "#14202B", "#243244", "#F8FAFC", "#8DA0B8", "#7DD3C7", "#314154", "#223244"),
            ["dark"] = Palette(
                "#11161A", "#171E24", "#202A32", "#293641", "#384752",
                "#6F9CFF", "#42B7A8", "#F3F5F7", "#C4CDD5", "#8997A4",
                "#0B1014", "#25313A", "#F7FAFC", "#8796A3", "#65D5C0", "#303E48", "#1A252D"),
            ["warm"] = Palette(
                "#F3F0EA", "#FFFCF7", "#ECE7DE", "#E1D9CE", "#D2C8BA",
                "#C45F3C", "#2F7D6D", "#292622", "#5E5750", "#8A8178",
                "#292723", "#3D3934", "#FFF9F1", "#A99F94", "#7FD1BC", "#4A453F", "#3A3732"),
        };

    public static string NormalizeKey(string? key) =>
        key is not null && Themes.ContainsKey(key) ? key.ToLowerInvariant() : "light";

    public static void ApplyTheme(string? key)
    {
        if (Application.Current is null) return;

        var palette = Themes[NormalizeKey(key)];
        foreach (var (resourceKey, colorValue) in palette)
        {
            var color = (Color)ColorConverter.ConvertFromString(colorValue);
            Application.Current.Resources[resourceKey] = color;
            Application.Current.Resources[BrushKeys[resourceKey]] = new SolidColorBrush(color);
        }
    }

    private static IReadOnlyDictionary<string, string> Palette(
        string bg, string surface1, string surface2, string surface3, string border,
        string accent, string accent2, string text, string text2, string text3,
        string sidebarBg, string sidebarBorder, string sidebarText, string sidebarMuted,
        string sidebarStatus, string sidebarSeparator, string sidebarHover) =>
        new Dictionary<string, string>
        {
            ["ThemeBgColor"] = bg,
            ["ThemeSurface1Color"] = surface1,
            ["ThemeSurface2Color"] = surface2,
            ["ThemeSurface3Color"] = surface3,
            ["ThemeBorderColor"] = border,
            ["ThemeAccentColor"] = accent,
            ["ThemeAccent2Color"] = accent2,
            ["ThemeTextColor"] = text,
            ["ThemeText2Color"] = text2,
            ["ThemeText3Color"] = text3,
            ["ThemeSidebarBgColor"] = sidebarBg,
            ["ThemeSidebarBorderColor"] = sidebarBorder,
            ["ThemeSidebarTextColor"] = sidebarText,
            ["ThemeSidebarMutedColor"] = sidebarMuted,
            ["ThemeSidebarStatusColor"] = sidebarStatus,
            ["ThemeSidebarSeparatorColor"] = sidebarSeparator,
            ["ThemeSidebarHoverColor"] = sidebarHover,
        };
}
