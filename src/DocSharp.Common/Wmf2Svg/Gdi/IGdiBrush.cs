namespace DocSharp.Wmf2Svg.Gdi;

public interface IGdiBrush : IGdiObject
{
    int Style { get; }
    int Color { get; }
    int Hatch { get; }
}