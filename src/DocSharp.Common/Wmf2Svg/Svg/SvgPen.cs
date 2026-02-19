using System.Text;
using System.Xml;
using DocSharp.Wmf2Svg.Gdi;

namespace DocSharp.Wmf2Svg.Svg;

public sealed class SvgPen : SvgObject, IGdiPen
{
    private readonly int _style;
    private readonly int _width;
    private readonly int _color;

    public SvgPen(SvgGdi gdi, int style, int width, int color) : base(gdi)
    {
        _style = style;
        _width = (width > 0) ? width : 1;
        _color = color;
    }

    public int Style => _style;
    public int Width => _width;
    public int Color => _color;

    public override int GetHashCode()
    {
        const int prime = 31;
        var result = 1;
        result = prime * result + _color;
        result = prime * result + _style;
        result = prime * result + _width;
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

        var other = (SvgPen)obj;
        if (_color != other._color)
        {
            return false;
        }

        if (_style != other._style)
        {
            return false;
        }

        if (_width != other._width)
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

        if (_style == GdiPenConstants.PS_NULL)
        {
            buffer.Append("stroke: none; ");
        }
        else
        {
            // stroke
            buffer.Append("stroke: " + ToColor(_color) + "; ");

            // stroke-width
            buffer.Append("stroke-width: " + _width + "; ");

            // stroke-linejoin
            buffer.Append("stroke-linejoin: round; ");

            // stroke-dasharray
            if (_width == 1 && GdiPenConstants.PS_DASH <= _style && _style <= GdiPenConstants.PS_DASHDOTDOT)
            {
                buffer.Append("stroke-dasharray: ");
                switch (_style)
                {
                    case GdiPenConstants.PS_DASH:
                        buffer.Append(ToRealSize(18) + "," + ToRealSize(6));
                        break;
                    case GdiPenConstants.PS_DOT:
                        buffer.Append(ToRealSize(3) + "," + ToRealSize(3));
                        break;
                    case GdiPenConstants.PS_DASHDOT:
                        buffer.Append(
                            ToRealSize(9) + "," +
                            ToRealSize(3) + "," +
                            ToRealSize(3) + "," +
                            ToRealSize(3));
                        break;
                    case GdiPenConstants.PS_DASHDOTDOT:
                        buffer.Append(
                            ToRealSize(9) + "," +
                            ToRealSize(3) + "," +
                            ToRealSize(3) + "," +
                            ToRealSize(3) + "," +
                            ToRealSize(3) + "," +
                            ToRealSize(3));
                        break;
                }

                buffer.Append("; ");
            }
        }

        if (buffer.Length > 0)
        {
            buffer.Length = buffer.Length - 1;
        }

        return buffer.ToString();
    }
}