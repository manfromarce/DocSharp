namespace DocSharp.Wmf2Svg.Gdi;

public sealed class Size
{
    public int Width { get; set; }

    public int Height { get; set; }

    public Size(int width, int height)
    {
        Width = width;
        Height = height;
    }

    public override string ToString()
    {
        return $"Size [Width={Width}, Height={Height}]";
    }
}