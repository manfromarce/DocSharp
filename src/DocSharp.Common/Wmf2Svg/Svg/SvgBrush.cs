using System.Globalization;
using System.Text;
using System.Xml;
using DocSharp.Wmf2Svg.Gdi;

namespace DocSharp.Wmf2Svg.Svg;

public sealed class SvgBrush : SvgObject, IGdiBrush
{
    private readonly int _style;
    private readonly int _color;
    private readonly int _hatch;

    public SvgBrush(SvgGdi gdi, int style, int color, int hatch) : base(gdi)
    {
        _style = style;
        _color = color;
        _hatch = hatch;
    }

    public int Style => _style;
    public int Color => _color;
    public int Hatch => _hatch;

    public XmlElement? CreateFillPattern(string id)
    {
        XmlElement? pattern = null;

        if (_style == GdiBrushConstants.BS_HATCHED)
        {
            var doc = Gdi.Document;
            pattern = doc.CreateElement("pattern");
            pattern.SetAttribute("id", id);
            pattern.SetAttribute("patternUnits", "userSpaceOnUse");
            pattern.SetAttribute("x", ToRealSize(0).ToString(CultureInfo.InvariantCulture));
            pattern.SetAttribute("y", ToRealSize(0).ToString(CultureInfo.InvariantCulture));
            pattern.SetAttribute("width", ToRealSize(8).ToString(CultureInfo.InvariantCulture));
            pattern.SetAttribute("height", ToRealSize(8).ToString(CultureInfo.InvariantCulture));

            if (Gdi.DC.BkMode == GdiConstants.OPAQUE)
            {
                var rect = doc.CreateElement("rect");
                rect.SetAttribute("fill", ToColor(Gdi.DC.BkColor));
                rect.SetAttribute("x", ToRealSize(0).ToString(CultureInfo.InvariantCulture));
                rect.SetAttribute("y", ToRealSize(0).ToString(CultureInfo.InvariantCulture));
                rect.SetAttribute("width", ToRealSize(8).ToString(CultureInfo.InvariantCulture));
                rect.SetAttribute("height", ToRealSize(8).ToString(CultureInfo.InvariantCulture));
                pattern.AppendChild(rect);
            }

            switch (_hatch)
            {
                case GdiBrushConstants.HS_HORIZONTAL:
                {
                    var path = doc.CreateElement("line");
                    path.SetAttribute("stroke", ToColor(_color));
                    path.SetAttribute("x1", ToRealSize(0).ToString(CultureInfo.InvariantCulture));
                    path.SetAttribute("y1", ToRealSize(4).ToString(CultureInfo.InvariantCulture));
                    path.SetAttribute("x2", ToRealSize(8).ToString(CultureInfo.InvariantCulture));
                    path.SetAttribute("y2", ToRealSize(4).ToString(CultureInfo.InvariantCulture));
                    pattern.AppendChild(path);
                }
                    break;
                case GdiBrushConstants.HS_VERTICAL:
                {
                    var path = doc.CreateElement("line");
                    path.SetAttribute("stroke", ToColor(_color));
                    path.SetAttribute("x1", ToRealSize(4).ToString(CultureInfo.InvariantCulture));
                    path.SetAttribute("y1", ToRealSize(0).ToString(CultureInfo.InvariantCulture));
                    path.SetAttribute("x2", ToRealSize(4).ToString(CultureInfo.InvariantCulture));
                    path.SetAttribute("y2", ToRealSize(8).ToString(CultureInfo.InvariantCulture));
                    pattern.AppendChild(path);
                }
                    break;
                case GdiBrushConstants.HS_FDIAGONAL:
                {
                    var path = doc.CreateElement("line");
                    path.SetAttribute("stroke", ToColor(_color));
                    path.SetAttribute("x1", ToRealSize(0).ToString(CultureInfo.InvariantCulture));
                    path.SetAttribute("y1", ToRealSize(0).ToString(CultureInfo.InvariantCulture));
                    path.SetAttribute("x2", ToRealSize(8).ToString(CultureInfo.InvariantCulture));
                    path.SetAttribute("y2", ToRealSize(8).ToString(CultureInfo.InvariantCulture));
                    pattern.AppendChild(path);
                }
                    break;
                case GdiBrushConstants.HS_BDIAGONAL:
                {
                    var path = doc.CreateElement("line");
                    path.SetAttribute("stroke", ToColor(_color));
                    path.SetAttribute("x1", ToRealSize(0).ToString(CultureInfo.InvariantCulture));
                    path.SetAttribute("y1", ToRealSize(8).ToString(CultureInfo.InvariantCulture));
                    path.SetAttribute("x2", ToRealSize(8).ToString(CultureInfo.InvariantCulture));
                    path.SetAttribute("y2", ToRealSize(0).ToString(CultureInfo.InvariantCulture));
                    pattern.AppendChild(path);
                }
                    break;
                case GdiBrushConstants.HS_CROSS:
                {
                    var path1 = doc.CreateElement("line");
                    path1.SetAttribute("stroke", ToColor(_color));
                    path1.SetAttribute("x1", ToRealSize(0).ToString(CultureInfo.InvariantCulture));
                    path1.SetAttribute("y1", ToRealSize(4).ToString(CultureInfo.InvariantCulture));
                    path1.SetAttribute("x2", ToRealSize(8).ToString(CultureInfo.InvariantCulture));
                    path1.SetAttribute("y2", ToRealSize(4).ToString(CultureInfo.InvariantCulture));
                    pattern.AppendChild(path1);
                    var path2 = doc.CreateElement("line");
                    path2.SetAttribute("stroke", ToColor(_color));
                    path2.SetAttribute("x1", ToRealSize(4).ToString(CultureInfo.InvariantCulture));
                    path2.SetAttribute("y1", ToRealSize(0).ToString(CultureInfo.InvariantCulture));
                    path2.SetAttribute("x2", ToRealSize(4).ToString(CultureInfo.InvariantCulture));
                    path2.SetAttribute("y2", ToRealSize(8).ToString(CultureInfo.InvariantCulture));
                    pattern.AppendChild(path2);
                }
                    break;
                case GdiBrushConstants.HS_DIAGCROSS:
                {
                    var path1 = doc.CreateElement("line");
                    path1.SetAttribute("stroke", ToColor(_color));
                    path1.SetAttribute("x1", ToRealSize(0).ToString(CultureInfo.InvariantCulture));
                    path1.SetAttribute("y1", ToRealSize(0).ToString(CultureInfo.InvariantCulture));
                    path1.SetAttribute("x2", ToRealSize(8).ToString(CultureInfo.InvariantCulture));
                    path1.SetAttribute("y2", ToRealSize(8).ToString(CultureInfo.InvariantCulture));
                    pattern.AppendChild(path1);
                    var path2 = doc.CreateElement("line");
                    path2.SetAttribute("stroke", ToColor(_color));
                    path2.SetAttribute("x1", ToRealSize(0).ToString(CultureInfo.InvariantCulture));
                    path2.SetAttribute("y1", ToRealSize(8).ToString(CultureInfo.InvariantCulture));
                    path2.SetAttribute("x2", ToRealSize(8).ToString(CultureInfo.InvariantCulture));
                    path2.SetAttribute("y2", ToRealSize(0).ToString(CultureInfo.InvariantCulture));
                    pattern.AppendChild(path2);
                }
                    break;
            }
        }

        return pattern;
    }

    public override int GetHashCode()
    {
        const int prime = 31;
        var result = 1;
        result = prime * result + _color;
        result = prime * result + _hatch;
        result = prime * result + _style;
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

        var other = (SvgBrush)obj;
        if (_color != other._color)
        {
            return false;
        }

        if (_hatch != other._hatch)
        {
            return false;
        }

        if (_style != other._style)
        {
            return false;
        }

        return true;
    }

    public XmlText CreateTextNode(string id)
    {
        return Gdi.Document.CreateTextNode("." + id + " { " + ToString() + " }\n");
    }

    public override string ToString()
    {
        var buffer = new StringBuilder();

        // fill
        switch (_style)
        {
            case GdiBrushConstants.BS_SOLID:
                buffer.Append("fill: ").Append(ToColor(_color)).Append("; ");
                break;
            case GdiBrushConstants.BS_HATCHED:
                break;
            default:
                buffer.Append("fill: none; ");
                break;
        }

        if (buffer.Length > 0)
        {
            buffer.Length = buffer.Length - 1;
        }

        return buffer.ToString();
    }
}