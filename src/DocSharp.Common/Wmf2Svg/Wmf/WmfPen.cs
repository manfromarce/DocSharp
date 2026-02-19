using DocSharp.Wmf2Svg.Gdi;

namespace DocSharp.Wmf2Svg.Wmf;

public sealed class WmfPen : WmfObject, IGdiPen
{
    public int Style { get; set; }
    public int Width { get; set; }
    public int Color { get; set; }

    public WmfPen(int id, int style, int width, int color) : base(id)
    {
        Style = style;
        Width = width;
        Color = color;
    }
}