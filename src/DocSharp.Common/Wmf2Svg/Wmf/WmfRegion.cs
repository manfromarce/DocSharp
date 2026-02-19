using DocSharp.Wmf2Svg.Gdi;

namespace DocSharp.Wmf2Svg.Wmf;

public class WmfRegion : WmfObject, IGdiRegion
{
    public WmfRegion(int id) : base(id)
    {
    }
}