using System.Globalization;
using System.Xml;

namespace DocSharp.Wmf2Svg.Svg;

public sealed class SvgRectRegion : SvgRegion
{
    private readonly int _left;
    private readonly int _top;
    private readonly int _right;
    private readonly int _bottom;

    public SvgRectRegion(SvgGdi gdi, int left, int top, int right, int bottom) : base(gdi)
    {
        _left = left;
        _top = top;
        _right = right;
        _bottom = bottom;
    }

    public int Left => _left;
    public int Top => _top;
    public int Right => _right;
    public int Bottom => _bottom;

    public override XmlElement CreateElement()
    {
        var elem = Gdi.Document.CreateElement("rect");
        elem.SetAttribute("x", ((int)Gdi.DC.ToAbsoluteX(Left)).ToString(CultureInfo.InvariantCulture));
        elem.SetAttribute("y", ((int)Gdi.DC.ToAbsoluteY(Top)).ToString(CultureInfo.InvariantCulture));
        elem.SetAttribute("width", ((int)Gdi.DC.ToRelativeX(Right - Left)).ToString(CultureInfo.InvariantCulture));
        elem.SetAttribute("height", ((int)Gdi.DC.ToRelativeY(Bottom - Top)).ToString(CultureInfo.InvariantCulture));
        return elem;
    }

    public override int GetHashCode()
    {
        const int prime = 31;
        var result = 1;
        result = prime * result + _bottom;
        result = prime * result + _left;
        result = prime * result + _right;
        result = prime * result + _top;
        return result;
    }

    public override bool Equals(object? obj)
    {
        if (this == obj)
        {
            return true;
        }

        if (obj == null)
        {
            return false;
        }

        if (GetType() != obj.GetType())
        {
            return false;
        }

        var other = (SvgRectRegion)obj;
        if (_bottom != other._bottom)
        {
            return false;
        }

        if (_left != other._left)
        {
            return false;
        }

        if (_right != other._right)
        {
            return false;
        }

        if (_top != other._top)
        {
            return false;
        }

        return true;
    }
}