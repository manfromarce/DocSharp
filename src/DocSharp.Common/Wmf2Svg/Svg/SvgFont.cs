using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Xml;
using DocSharp.Wmf2Svg.Gdi;

namespace DocSharp.Wmf2Svg.Svg;

public sealed class SvgFont : SvgObject, IGdiFont
{
    private readonly int _height;
    private readonly int _width;
    private readonly int _escapement;
    private readonly int _orientation;
    private readonly int _weight;
    private readonly bool _italic;
    private readonly bool _underline;
    private readonly bool _strikeout;
    private readonly int _charset;
    private readonly int _outPrecision;
    private readonly int _clipPrecision;
    private readonly int _quality;
    private readonly int _pitchAndFamily;

    private readonly string _faceName;
    private readonly double _heightMultiply;
    private readonly string? _lang;

    public SvgFont(
        SvgGdi gdi,
        int height,
        int width,
        int escapement,
        int orientation,
        int weight,
        bool italic,
        bool underline,
        bool strikeout,
        int charset,
        int outPrecision,
        int clipPrecision,
        int quality,
        int pitchAndFamily,
        byte[] faceName) : base(gdi)
    {
        _height = height;
        _width = width;
        _escapement = escapement;
        _orientation = orientation;
        _weight = weight;
        _italic = italic;
        _underline = underline;
        _strikeout = strikeout;
        _outPrecision = outPrecision;
        _clipPrecision = clipPrecision;
        _quality = quality;
        _pitchAndFamily = pitchAndFamily;
        _faceName = Helper.ConvertString(faceName, charset);

        var altCharset = gdi.GetProperty("font-charset." + _faceName);
        if (altCharset != null)
        {
            _charset = int.Parse(altCharset, NumberStyles.Integer, CultureInfo.InvariantCulture);
        }
        else
        {
            _charset = charset;
        }

        // xml:lang
        _lang = Helper.GetLanguage(_charset);

        var heightMultiply = 1.0;
        var emheight = gdi.GetProperty("font-emheight." + _faceName);
        if (emheight == null)
        {
            var alter = gdi.GetProperty("alternative-font." + _faceName);
            if (alter != null)
            {
                emheight = gdi.GetProperty("font-emheight." + alter);
            }
        }

        if (emheight != null)
        {
            heightMultiply = double.Parse(emheight, CultureInfo.InvariantCulture);
        }

        _heightMultiply = heightMultiply;
    }

    public int Height => _height;
    public int Width => _width;
    public int Escapement => _escapement;
    public int Orientation => _orientation;
    public int Weight => _weight;
    public bool IsItalic => _italic;
    public bool IsUnderlined => _underline;
    public bool IsStrikedOut => _strikeout;
    public int Charset => _charset;
    public int OutPrecision => _outPrecision;
    public int ClipPrecision => _clipPrecision;
    public int Quality => _quality;
    public int PitchAndFamily => _pitchAndFamily;
    public string FaceName => _faceName;
    public string? Lang => _lang;

    public int FontSize => Math.Abs((int)Gdi.DC.ToRelativeY(_height * _heightMultiply));

    public override int GetHashCode()
    {
        const int prime = 31;
        var result = 1;
        result = prime * result + _charset;
        result = prime * result + _clipPrecision;
        result = prime * result + _escapement;
        result = prime * result + (_faceName?.GetHashCode() ?? 0);
        // result = prime * result + (_faceName?.GetHashCode(StringComparison.Ordinal) ?? 0);
        result = prime * result + _height;
        result = prime * result + (_italic ? 1231 : 1237);
        result = prime * result + _orientation;
        result = prime * result + _outPrecision;
        result = prime * result + _pitchAndFamily;
        result = prime * result + _quality;
        result = prime * result + (_strikeout ? 1231 : 1237);
        result = prime * result + (_underline ? 1231 : 1237);
        result = prime * result + _weight;
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

        var other = (SvgFont)obj;
        if (_charset != other._charset)
        {
            return false;
        }

        if (_clipPrecision != other._clipPrecision)
        {
            return false;
        }

        if (_escapement != other._escapement)
        {
            return false;
        }

        if (_faceName == null)
        {
            if (other._faceName != null)
            {
                return false;
            }
        }
        else if (!_faceName.Equals(other._faceName, StringComparison.OrdinalIgnoreCase))
        {
            return false;
        }

        if (_height != other._height)
        {
            return false;
        }

        if (_italic != other._italic)
        {
            return false;
        }

        if (_orientation != other._orientation)
        {
            return false;
        }

        if (_outPrecision != other._outPrecision)
        {
            return false;
        }

        if (_pitchAndFamily != other._pitchAndFamily)
        {
            return false;
        }

        if (_quality != other._quality)
        {
            return false;
        }

        if (_strikeout != other._strikeout)
        {
            return false;
        }

        if (_underline != other._underline)
        {
            return false;
        }

        if (_weight != other._weight)
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
        var weight = _weight;

        // font-style
        if (_italic)
        {
            buffer.Append("font-style: italic; ");
        }

        // font-weight
        if (weight != GdiFontConstants.FW_DONTCARE && weight != GdiFontConstants.FW_NORMAL)
        {
            if (weight < 100)
            {
                weight = 100;
            }
            else if (weight > 900)
            {
                weight = 900;
            }
            else
            {
                weight = (weight / 100) * 100;
            }

            if (weight == GdiFontConstants.FW_BOLD)
            {
                buffer.Append("font-weight: bold; ");
            }
            else
            {
                buffer.Append("font-weight: " + weight + "; ");
            }
        }

        var fontSize = FontSize;
        if (fontSize != 0)
        {
            buffer.Append("font-size: ").Append(fontSize).Append("px; ");
        }

        // font-family
        var fontList = new List<string>();
        if (_faceName.Length != 0)
        {
            var fontFamily = _faceName;
            if (_faceName[0] == '@')
            {
                fontFamily = _faceName.Substring(1);
            }

            fontList.Add(fontFamily);

            var altfont = Gdi.GetProperty("alternative-font." + fontFamily);
            if (altfont != null && altfont.Length != 0)
            {
                fontList.Add(altfont);
            }
        }

        // int pitch = pitchAndFamily & 0x00000003;
        var family = _pitchAndFamily & 0x000000F0;
        switch (family)
        {
            case GdiFontConstants.FF_DECORATIVE:
                fontList.Add("fantasy");
                break;
            case GdiFontConstants.FF_MODERN:
                fontList.Add("monospace");
                break;
            case GdiFontConstants.FF_ROMAN:
                fontList.Add("serif");
                break;
            case GdiFontConstants.FF_SCRIPT:
                fontList.Add("cursive");
                break;
            case GdiFontConstants.FF_SWISS:
                fontList.Add("sans-serif");
                break;
        }

        if (fontList.Count > 0)
        {
            buffer.Append("font-family:");
            for (var i = 0; i < fontList.Count; i++)
            {
                var font = fontList[i];
                if (font.Contains(" ", StringComparison.Ordinal))
                {
                    buffer.Append(" \"" + font + "\"");
                }
                else
                {
                    buffer.Append(" " + font);
                }

                if (i < fontList.Count - 1)
                {
                    buffer.Append(',');
                }
            }

            buffer.Append("; ");
        }

        // text-decoration
        if (_underline || _strikeout)
        {
            buffer.Append("text-decoration:");
            if (_underline)
            {
                buffer.Append(" underline");
            }

            if (_strikeout)
            {
                buffer.Append(" overline");
            }

            buffer.Append("; ");
        }

        if (buffer.Length > 0)
        {
            buffer.Length = buffer.Length - 1;
        }

        return buffer.ToString();
    }
}