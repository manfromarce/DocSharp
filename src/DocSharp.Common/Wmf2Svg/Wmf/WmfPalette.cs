using DocSharp.Wmf2Svg.Gdi;

namespace DocSharp.Wmf2Svg.Wmf;

public sealed class WmfPalette : WmfObject, IGdiPalette
{
    public int Version { get; set; }
    public int[] Entries { get; set; }

    public WmfPalette(int id, int version, int[] entries) : base(id)
    {
        Version = version;
        Entries = entries;
    }
}