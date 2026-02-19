using DocSharp.Wmf2Svg.Gdi;

namespace DocSharp.Wmf2Svg.Svg;

public sealed class SvgPalette : SvgObject, IGdiPalette
{
    private readonly int _version;
    private readonly int[] _entries;

    public SvgPalette(SvgGdi gdi, int version, int[] entries) : base(gdi)
    {
        _version = version;
        _entries = entries;
    }

    public int Version => _version;
    public int[] Entries => _entries;
}