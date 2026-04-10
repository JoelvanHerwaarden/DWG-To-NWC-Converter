using System;
using System.Runtime.InteropServices;
using System.Windows;
using System.Windows.Interop;
using MediaColor = System.Windows.Media.Color;
using MediaColorConverter = System.Windows.Media.ColorConverter;

namespace DWGToNWCConverter;

internal static class WindowThemeHelper
{
    private const int DwmUseImmersiveDarkModeAttribute = 20;
    private const int DwmBorderColorAttribute = 34;
    private const int DwmCaptionColorAttribute = 35;
    private const int DwmTextColorAttribute = 36;

    private static readonly MediaColor CaptionColor = (MediaColor)MediaColorConverter.ConvertFromString("#252523");
    private static readonly MediaColor BorderColor = (MediaColor)MediaColorConverter.ConvertFromString("#3d3d39");
    private static readonly MediaColor TextColor = (MediaColor)MediaColorConverter.ConvertFromString("#c2c0b6");

    public static void ApplyTheme(Window window)
    {
        if (!OperatingSystem.IsWindows())
        {
            return;
        }

        var handle = new WindowInteropHelper(window).Handle;
        if (handle == IntPtr.Zero)
        {
            return;
        }

        var enabled = 1;
        DwmSetWindowAttribute(handle, DwmUseImmersiveDarkModeAttribute, ref enabled, sizeof(int));

        var borderColor = ToColorRef(BorderColor);
        DwmSetWindowAttribute(handle, DwmBorderColorAttribute, ref borderColor, sizeof(int));

        var captionColor = ToColorRef(CaptionColor);
        DwmSetWindowAttribute(handle, DwmCaptionColorAttribute, ref captionColor, sizeof(int));

        var textColor = ToColorRef(TextColor);
        DwmSetWindowAttribute(handle, DwmTextColorAttribute, ref textColor, sizeof(int));
    }

    private static int ToColorRef(MediaColor color) => color.R | (color.G << 8) | (color.B << 16);

    [DllImport("dwmapi.dll")]
    private static extern int DwmSetWindowAttribute(
        IntPtr hwnd,
        int dwAttribute,
        ref int pvAttribute,
        int cbAttribute);
}
