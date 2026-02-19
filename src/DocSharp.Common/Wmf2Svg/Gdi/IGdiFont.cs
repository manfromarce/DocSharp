namespace DocSharp.Wmf2Svg.Gdi;

public interface IGdiFont : IGdiObject
{
    int Height { get; }
    int Width { get; }
    int Escapement { get; }
    int Orientation { get; }
    int Weight { get; }
    bool IsItalic { get; }
    bool IsUnderlined { get; }
    bool IsStrikedOut { get; }
    int Charset { get; }
    int OutPrecision { get; }
    int ClipPrecision { get; }
    int Quality { get; }
    int PitchAndFamily { get; }
    string FaceName { get; }
}