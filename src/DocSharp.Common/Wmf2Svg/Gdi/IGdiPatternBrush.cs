namespace DocSharp.Wmf2Svg.Gdi;

public interface IGdiPatternBrush : IGdiObject
{
    byte[] Pattern { get; }
}