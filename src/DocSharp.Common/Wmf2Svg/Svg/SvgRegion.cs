using System.Xml;
using DocSharp.Wmf2Svg.Gdi;

namespace DocSharp.Wmf2Svg.Svg;

public abstract class SvgRegion : SvgObject, IGdiRegion
{
    protected SvgRegion(SvgGdi gdi) : base(gdi)
    {
    }

    public abstract XmlElement CreateElement();
}