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
                "#F5F7FA", "#FFFFFF", "#F8FAFC", "#EAF0F6", "#D7E0E9",
                "#155EEF", "#0E9384", "#101828", "#475467", "#667085",
                "#101828", "#27364A", "#F8FAFC", "#98A2B3", "#5FE9D0", "#344054", "#1D2939"),
            ["dark"] = Palette(
                "#0C111D", "#161B26", "#1D2939", "#344054", "#344054",
                "#84ADFF", "#5FE9D0", "#F2F4F7", "#D0D5DD", "#98A2B3",
                "#070B12", "#1D2939", "#F9FAFB", "#98A2B3", "#5FE9D0", "#344054", "#182230"),
            ["warm"] = Palette(
                "#F7F5F2", "#FFFFFF", "#F4F1EC", "#EAE4DC", "#D8D0C5",
                "#C55232", "#187E72", "#1F2937", "#4B5563", "#7C746C",
                "#20262E", "#374151", "#FFFDF9", "#A7AFBA", "#75E0CD", "#3D4652", "#303844"),
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
