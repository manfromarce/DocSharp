namespace DocSharp.Wmf2Svg.Gdi;

public interface IGdiPen : IGdiObject
{
    int Style { get; }
    int Width { get; }
    int Color { get; }
}