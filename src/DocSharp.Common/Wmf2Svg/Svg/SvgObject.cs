using System;

namespace DocSharp.Wmf2Svg.Svg;

public abstract class SvgObject
{
    private readonly SvgGdi _gdi;

    protected SvgObject(SvgGdi gdi)
    {
        _gdi = gdi ?? throw new ArgumentNullException(nameof(gdi));
    }

    public SvgGdi Gdi => _gdi;

    public int ToRealSize(int px)
    {
        return Gdi.DC.Dpi * px / 90;
    }

    public static string ToColor(int color)
    {
        var b = (0x00FF0000 & color) >> 16;
        var g = (0x0000FF00 & color) >> 8;
        var r = (0x000000FF & color);

        return $"rgb({r},{g},{b})";
    }
}