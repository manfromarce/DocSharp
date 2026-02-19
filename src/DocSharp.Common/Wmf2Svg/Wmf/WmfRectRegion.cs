namespace DocSharp.Wmf2Svg.Wmf;

public sealed class WmfRectRegion : WmfRegion
{
    public int Left { get; set; }
    public int Top { get; set; }
    public int Right { get; set; }
    public int Bottom { get; set; }

    public WmfRectRegion(int id, int left, int top, int right, int bottom) : base(id)
    {
        Left = left;
        Top = top;
        Right = right;
        Bottom = bottom;
    }
}