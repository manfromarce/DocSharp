using DocSharp.Wmf2Svg.Gdi;

namespace DocSharp.Wmf2Svg.Svg;

public sealed class SvgPatternBrush : SvgObject, IGdiPatternBrush
{
    private readonly byte[] _pattern;

    public SvgPatternBrush(SvgGdi gdi, byte[] pattern) : base(gdi)
    {
        _pattern = pattern;
    }

    public byte[] Pattern => _pattern;
}