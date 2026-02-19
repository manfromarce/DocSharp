using DocSharp.Wmf2Svg.Gdi;

namespace DocSharp.Wmf2Svg.Wmf;

public sealed class WmfPatternBrush : WmfObject, IGdiPatternBrush
{
    public byte[] Pattern { get; set; }

    public WmfPatternBrush(int id, byte[] pattern) : base(id)
    {
        Pattern = pattern;
    }
}