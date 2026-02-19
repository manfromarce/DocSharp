using DocSharp.Wmf2Svg.Gdi;

namespace DocSharp.Wmf2Svg.Wmf;

public sealed class WmfBrush : WmfObject, IGdiBrush
{
    public int Style { get; set; }
    public int Color { get; set; }
    public int Hatch { get; set; }

    public WmfBrush(int id, int style, int color, int hatch) : base(id)
    {
        Style = style;
        Color = color;
        Hatch = hatch;
    }
}