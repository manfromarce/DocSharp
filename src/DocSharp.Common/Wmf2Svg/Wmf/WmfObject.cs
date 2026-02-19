using DocSharp.Wmf2Svg.Gdi;

namespace DocSharp.Wmf2Svg.Wmf;

public abstract class WmfObject : IGdiObject
{
    public int ID { get; set; }

    protected WmfObject(int id)
    {
        ID = id;
    }
}