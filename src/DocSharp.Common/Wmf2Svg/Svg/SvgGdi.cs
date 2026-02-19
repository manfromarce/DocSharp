using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Reflection;
using System.Text;
using System.Text.RegularExpressions;
using System.Xml;
using DocSharp;
using DocSharp.Wmf2Svg.Gdi;

namespace DocSharp.Wmf2Svg.Svg;

public sealed class SvgGdi : IGdi
{
    private bool _compatible;
    private readonly IImageConverter? _imageConverter;
    private bool _replaceSymbolFont;
    private SvgDc _dc = null!;
    private readonly List<SvgDc> _saveDC = new();
    private XmlDocument _doc = null!;
    private XmlElement _parentNode = null!;
    private XmlElement _styleNode = null!;
    private XmlElement _defsNode = null!;
    private int _brushNo;
    private int _fontNo;
    private int _penNo;
    private int _patternNo;
    private int _rgnNo;
    private int _clipPathNo;
    private int _maskNo;
    private readonly Dictionary<IGdiObject, string> _nameMap = new();
    private readonly StringBuilder _buffer = new(1000);
    private SvgBrush _defaultBrush = null!;
    private SvgPen _defaultPen = null!;
    private SvgFont? _defaultFont;

    public SvgGdi() : this(false)
    {
    }

    public SvgGdi(bool compatible, IImageConverter? imageConverter = null)
    {
        _compatible = compatible;
        _imageConverter = imageConverter;
        _doc = new XmlDocument();
        var root = _doc.CreateElement("svg", "http://www.w3.org/2000/svg");
        _doc.AppendChild(root);
    }

    public void Write(Stream output)
    {
        var settings = new XmlWriterSettings
        {
            Indent = true,
            Encoding = Encoding.UTF8,
            OmitXmlDeclaration = false
        };

        using var writer = XmlWriter.Create(output, settings);
        _doc.Save(writer);
        writer.Flush();
    }

    public bool Compatible
    {
        get => _compatible;
        set => _compatible = value;
    }

    public bool ReplaceSymbolFont
    {
        get => _replaceSymbolFont;
        set => _replaceSymbolFont = value;
    }

    public SvgDc DC => _dc;

    public string? GetProperty(string key) => Properties.Values.TryGetValue(key, out var value) ? value : null;

    public XmlDocument Document => _doc;

    public XmlElement DefsElement => _defsNode;

    public XmlElement StyleElement => _styleNode;

    public void PlaceableHeader(int wsx, int wsy, int wex, int wey, int dpi)
    {
        if (_parentNode == null)
        {
            Init();
        }

        _dc.SetWindowExtEx(Math.Abs(wex - wsx), Math.Abs(wey - wsy), null);
        _dc.Dpi = dpi;

        var root = _doc.DocumentElement!;
        root.SetAttribute("width", (Math.Abs(wex - wsx) / (double)_dc.Dpi) + "in");
        root.SetAttribute("height", (Math.Abs(wey - wsy) / (double)_dc.Dpi) + "in");
    }

    public void Header()
    {
        if (_parentNode == null)
        {
            Init();
        }
    }

    private void Init()
    {
        _dc = new SvgDc(this);

        var root = _doc.DocumentElement!;
        root.SetAttribute("xmlns", "http://www.w3.org/2000/svg");
        root.SetAttribute("xmlns:xlink", "http://www.w3.org/1999/xlink");

        _defsNode = _doc.CreateElement("defs", "http://www.w3.org/2000/svg");
        root.AppendChild(_defsNode);

        _styleNode = _doc.CreateElement("style", "http://www.w3.org/2000/svg");
        _styleNode.SetAttribute("type", "text/css");
        root.AppendChild(_styleNode);

        _parentNode = _doc.CreateElement("g", "http://www.w3.org/2000/svg");
        root.AppendChild(_parentNode);

        _defaultBrush = (SvgBrush)CreateBrushIndirect(GdiBrushConstants.BS_SOLID, 0x00FFFFFF, 0);
        _defaultPen = (SvgPen)CreatePenIndirect(GdiPenConstants.PS_SOLID, 1, 0x00000000);
        _defaultFont = null;

        _dc.Brush = _defaultBrush;
        _dc.Pen = _defaultPen;
        _dc.Font = _defaultFont;
    }

    public void AnimatePalette(IGdiPalette palette, int startIndex, int[] entries)
    {
        // Not implemented
    }

    public void Arc(int sxr, int syr, int exr, int eyr, int sxa, int sya, int exa, int eya)
    {
        var rx = Math.Abs(exr - sxr) / 2.0;
        var ry = Math.Abs(eyr - syr) / 2.0;
        if (rx <= 0 || ry <= 0)
        {
            return;
        }

        var cx = Math.Min(sxr, exr) + rx;
        var cy = Math.Min(syr, eyr) + ry;

        XmlElement elem;
        if (sxa == exa && sya == eya)
        {
            if (Math.Abs(rx - ry) < 0.0001)
            {
                elem = _doc.CreateElement("circle", "http://www.w3.org/2000/svg");
                elem.SetAttribute("cx", _dc.ToAbsoluteX(cx).ToString(CultureInfo.InvariantCulture));
                elem.SetAttribute("cy", _dc.ToAbsoluteY(cy).ToString(CultureInfo.InvariantCulture));
                elem.SetAttribute("r", _dc.ToRelativeX(rx).ToString(CultureInfo.InvariantCulture));
            }
            else
            {
                elem = _doc.CreateElement("ellipse", "http://www.w3.org/2000/svg");
                elem.SetAttribute("cx", _dc.ToAbsoluteX(cx).ToString(CultureInfo.InvariantCulture));
                elem.SetAttribute("cy", _dc.ToAbsoluteY(cy).ToString(CultureInfo.InvariantCulture));
                elem.SetAttribute("rx", _dc.ToRelativeX(rx).ToString(CultureInfo.InvariantCulture));
                elem.SetAttribute("ry", _dc.ToRelativeY(ry).ToString(CultureInfo.InvariantCulture));
            }
        }
        else
        {
            var sa = Math.Atan2((sya - cy) * rx, (sxa - cx) * ry);
            var sx = rx * Math.Cos(sa);
            var sy = ry * Math.Sin(sa);

            var ea = Math.Atan2((eya - cy) * rx, (exa - cx) * ry);
            var ex = rx * Math.Cos(ea);
            var ey = ry * Math.Sin(ea);

            var a = Math.Atan2((ex - sx) * (-sy) - (ey - sy) * (-sx), (ex - sx) * (-sx) + (ey - sy) * (-sy));

            elem = _doc.CreateElement("path", "http://www.w3.org/2000/svg");
            elem.SetAttribute("d", $"M {_dc.ToAbsoluteX(sx + cx)},{_dc.ToAbsoluteY(sy + cy)} A {_dc.ToRelativeX(rx)},{_dc.ToRelativeY(ry)} 0 {(a > 0 ? "1" : "0")} 0 {_dc.ToAbsoluteX(ex + cx)},{_dc.ToAbsoluteY(ey + cy)}");
        }

        if (_dc.Pen != null)
        {
            elem.SetAttribute("class", GetClassString(_dc.Pen));
        }

        elem.SetAttribute("fill", "none");
        _parentNode.AppendChild(elem);
    }

    public void BitBlt(byte[] image, int dx, int dy, int dw, int dh, int sx, int sy, long rop)
    {
        BmpToSvg(image, dx, dy, dw, dh, sx, sy, dw, dh, GdiConstants.DIB_RGB_COLORS, rop);
    }

    public void Chord(int sxr, int syr, int exr, int eyr, int sxa, int sya, int exa, int eya)
    {
        var rx = Math.Abs(exr - sxr) / 2.0;
        var ry = Math.Abs(eyr - syr) / 2.0;
        if (rx <= 0 || ry <= 0)
        {
            return;
        }

        var cx = Math.Min(sxr, exr) + rx;
        var cy = Math.Min(syr, eyr) + ry;

        XmlElement elem;
        if (sxa == exa && sya == eya)
        {
            if (Math.Abs(rx - ry) < 0.0001)
            {
                elem = _doc.CreateElement("circle", "http://www.w3.org/2000/svg");
                elem.SetAttribute("cx", _dc.ToAbsoluteX(cx).ToString(CultureInfo.InvariantCulture));
                elem.SetAttribute("cy", _dc.ToAbsoluteY(cy).ToString(CultureInfo.InvariantCulture));
                elem.SetAttribute("r", _dc.ToRelativeX(rx).ToString(CultureInfo.InvariantCulture));
            }
            else
            {
                elem = _doc.CreateElement("ellipse", "http://www.w3.org/2000/svg");
                elem.SetAttribute("cx", _dc.ToAbsoluteX(cx).ToString(CultureInfo.InvariantCulture));
                elem.SetAttribute("cy", _dc.ToAbsoluteY(cy).ToString(CultureInfo.InvariantCulture));
                elem.SetAttribute("rx", _dc.ToRelativeX(rx).ToString(CultureInfo.InvariantCulture));
                elem.SetAttribute("ry", _dc.ToRelativeY(ry).ToString(CultureInfo.InvariantCulture));
            }
        }
        else
        {
            var sa = Math.Atan2((sya - cy) * rx, (sxa - cx) * ry);
            var sx = rx * Math.Cos(sa);
            var sy = ry * Math.Sin(sa);

            var ea = Math.Atan2((eya - cy) * rx, (exa - cx) * ry);
            var ex = rx * Math.Cos(ea);
            var ey = ry * Math.Sin(ea);

            var a = Math.Atan2((ex - sx) * (-sy) - (ey - sy) * (-sx), (ex - sx) * (-sx) + (ey - sy) * (-sy));

            elem = _doc.CreateElement("path", "http://www.w3.org/2000/svg");
            elem.SetAttribute("d", $"M {_dc.ToAbsoluteX(sx + cx)},{_dc.ToAbsoluteY(sy + cy)} A {_dc.ToRelativeX(rx)},{_dc.ToRelativeY(ry)} 0 {(a > 0 ? "1" : "0")} 0 {_dc.ToAbsoluteX(ex + cx)},{_dc.ToAbsoluteY(ey + cy)} Z");
        }

        if (_dc.Pen != null || _dc.Brush != null)
        {
            elem.SetAttribute("class", GetClassString(_dc.Pen, _dc.Brush));
            if (_dc.Brush?.Style == GdiBrushConstants.BS_HATCHED)
            {
                var id = "pattern" + (_patternNo++);
                elem.SetAttribute("fill", $"url(#{id})");
                _defsNode.AppendChild(_dc.Brush.CreateFillPattern(id)!);
            }
        }

        _parentNode.AppendChild(elem);
    }

    public IGdiBrush CreateBrushIndirect(int style, int color, int hatch)
    {
        var brush = new SvgBrush(this, style, color, hatch);
        if (!_nameMap.ContainsKey(brush))
        {
            var name = "brush" + (_brushNo++);
            _nameMap[brush] = name;
            _styleNode.AppendChild(brush.CreateTextNode(name));
        }

        return brush;
    }

    public IGdiFont CreateFontIndirect(int height, int width, int escapement, int orientation, int weight,
        bool italic, bool underline, bool strikeout, int charset, int outPrecision, int clipPrecision,
        int quality, int pitchAndFamily, byte[] faceName)
    {
        var font = new SvgFont(this, height, width, escapement, orientation, weight, italic, underline,
            strikeout, charset, outPrecision, clipPrecision, quality, pitchAndFamily, faceName);
        if (!_nameMap.ContainsKey(font))
        {
            var name = "font" + (_fontNo++);
            _nameMap[font] = name;
            _styleNode.AppendChild(font.CreateTextNode(name));
        }

        return font;
    }

    public IGdiPalette CreatePalette(int version, int[] palEntry)
    {
        return new SvgPalette(this, version, palEntry);
    }

    public IGdiPatternBrush CreatePatternBrush(byte[] image)
    {
        return new SvgPatternBrush(this, image);
    }

    public IGdiPen CreatePenIndirect(int style, int width, int color)
    {
        var pen = new SvgPen(this, style, width, color);
        if (!_nameMap.ContainsKey(pen))
        {
            var name = "pen" + (_penNo++);
            _nameMap[pen] = name;
            _styleNode.AppendChild(pen.CreateTextNode(name));
        }

        return pen;
    }

    public IGdiRegion CreateRectRgn(int left, int top, int right, int bottom)
    {
        var rgn = new SvgRectRegion(this, left, top, right, bottom);
        if (!_nameMap.ContainsKey(rgn))
        {
            _nameMap[rgn] = "rgn" + (_rgnNo++);
            _defsNode.AppendChild(rgn.CreateElement());
        }

        return rgn;
    }

    public void DeleteObject(IGdiObject obj)
    {
        if (_dc.Brush == obj)
        {
            _dc.Brush = _defaultBrush;
        }
        else if (_dc.Font == obj)
        {
            _dc.Font = _defaultFont;
        }
        else if (_dc.Pen == obj)
        {
            _dc.Pen = _defaultPen;
        }
    }

    public void DibBitBlt(byte[] image, int dx, int dy, int dw, int dh, int sx, int sy, long rop)
    {
        BitBlt(image, dx, dy, dw, dh, sx, sy, rop);
    }

    public IGdiPatternBrush DibCreatePatternBrush(byte[] image, int usage)
    {
        return new SvgPatternBrush(this, image);
    }

    public void DibStretchBlt(byte[] image, int dx, int dy, int dw, int dh, int sx, int sy, int sw, int sh, long rop)
    {
        StretchDIBits(dx, dy, dw, dh, sx, sy, sw, sh, image, GdiConstants.DIB_RGB_COLORS, rop);
    }

    public void Ellipse(int sx, int sy, int ex, int ey)
    {
        var elem = _doc.CreateElement("ellipse", "http://www.w3.org/2000/svg");

        if (_dc.Pen != null || _dc.Brush != null)
        {
            elem.SetAttribute("class", GetClassString(_dc.Pen, _dc.Brush));
            if (_dc.Brush?.Style == GdiBrushConstants.BS_HATCHED)
            {
                var id = "pattern" + (_patternNo++);
                elem.SetAttribute("fill", $"url(#{id})");
                _defsNode.AppendChild(_dc.Brush.CreateFillPattern(id)!);
            }
        }

        elem.SetAttribute("cx", ((int)_dc.ToAbsoluteX((sx + ex) / 2.0)).ToString(CultureInfo.InvariantCulture));
        elem.SetAttribute("cy", ((int)_dc.ToAbsoluteY((sy + ey) / 2.0)).ToString(CultureInfo.InvariantCulture));
        elem.SetAttribute("rx", ((int)_dc.ToRelativeX((ex - sx) / 2.0)).ToString(CultureInfo.InvariantCulture));
        elem.SetAttribute("ry", ((int)_dc.ToRelativeY((ey - sy) / 2.0)).ToString(CultureInfo.InvariantCulture));
        _parentNode.AppendChild(elem);
    }

    public void Escape(byte[] data)
    {
    }

    public int ExcludeClipRect(int left, int top, int right, int bottom)
    {
        var mask = _dc.Mask;
        if (mask != null)
        {
            mask = (XmlElement)mask.CloneNode(true);
            var name = "mask" + (_maskNo++);
            mask.SetAttribute("id", name);
            _defsNode.AppendChild(mask);

            var unclip = _doc.CreateElement("rect", "http://www.w3.org/2000/svg");
            unclip.SetAttribute("x", ((int)_dc.ToAbsoluteX(left)).ToString(CultureInfo.InvariantCulture));
            unclip.SetAttribute("y", ((int)_dc.ToAbsoluteY(top)).ToString(CultureInfo.InvariantCulture));
            unclip.SetAttribute("width", ((int)_dc.ToRelativeX(right - left)).ToString(CultureInfo.InvariantCulture));
            unclip.SetAttribute("height", ((int)_dc.ToRelativeY(bottom - top)).ToString(CultureInfo.InvariantCulture));
            unclip.SetAttribute("fill", "black");
            mask.AppendChild(unclip);
            _dc.Mask = mask;

            return GdiRegionConstants.COMPLEXREGION;
        }

        return GdiRegionConstants.NULLREGION;
    }

    public void ExtFloodFill(int x, int y, int color, int type)
    {
        // Not implemented
    }

    public void ExtTextOut(int x, int y, int options, int[]? rect, byte[] text, int[]? lpdx)
    {
        var elem = _doc.CreateElement("text", "http://www.w3.org/2000/svg");

        var escapement = 0;
        var vertical = false;
        if (_dc.Font != null)
        {
            elem.SetAttribute("class", GetClassString(_dc.Font));
            if (_dc.Font.FaceName.StartsWith("@"))
            {
                vertical = true;
                escapement = _dc.Font.Escapement - 2700;
            }
            else
            {
                escapement = _dc.Font.Escapement;
            }
        }

        elem.SetAttribute("fill", SvgObject.ToColor(_dc.TextColor));

        // style
        _buffer.Clear();
        var align = _dc.TextAlign;

        if ((align & (GdiConstants.TA_LEFT | GdiConstants.TA_CENTER | GdiConstants.TA_RIGHT)) == GdiConstants.TA_RIGHT)
        {
            _buffer.Append("text-anchor: end; ");
        }
        else if ((align & (GdiConstants.TA_LEFT | GdiConstants.TA_CENTER | GdiConstants.TA_RIGHT)) == GdiConstants.TA_CENTER)
        {
            _buffer.Append("text-anchor: middle; ");
        }

        if (_compatible)
        {
            _buffer.Append("dominant-baseline: alphabetic; ");
        }
        else
        {
            if (vertical)
            {
                elem.SetAttribute("writing-mode", "tb");
            }
            else
            {
                if ((align & (GdiConstants.TA_BOTTOM | GdiConstants.TA_TOP | GdiConstants.TA_BASELINE)) == GdiConstants.TA_BASELINE)
                {
                    _buffer.Append("dominant-baseline: alphabetic; ");
                }
                else
                {
                    _buffer.Append("dominant-baseline: text-before-edge; ");
                }
            }
        }

        if ((align & GdiConstants.TA_RTLREADING) == GdiConstants.TA_RTLREADING || (options & GdiConstants.ETO_RTLREADING) > 0)
        {
            _buffer.Append("unicode-bidi: bidi-override; direction: rtl; ");
        }

        if (_dc.TextSpace > 0)
        {
            _buffer.Append($"word-spacing: ").Append(_dc.TextSpace).Append("; ");
        }

        if (_buffer.Length > 0)
        {
            _buffer.Length--;
            elem.SetAttribute("style", _buffer.ToString());
        }

        elem.SetAttribute("stroke", "none");

        if ((align & (GdiConstants.TA_NOUPDATECP | GdiConstants.TA_UPDATECP)) == GdiConstants.TA_UPDATECP)
        {
            x = _dc.CurrentX;
            y = _dc.CurrentY;
        }

        // x
        var ax = (int)_dc.ToAbsoluteX(x);
        var width = 0;
        if (vertical)
        {
            elem.SetAttribute("x", ax.ToString(CultureInfo.InvariantCulture));
            if (_dc.Font != null)
            {
                width = Math.Abs(_dc.Font.FontSize);
            }
        }
        else
        {
            if (_dc.Font != null)
            {
                lpdx = Helper.FixTextDx(_dc.Font.Charset, text, lpdx);
            }

            if (lpdx?.Length > 0)
            {
                for (var i = 0; i < lpdx.Length; i++)
                {
                    width += lpdx[i];
                }

                var tx = x;

                if ((align & (GdiConstants.TA_LEFT | GdiConstants.TA_CENTER | GdiConstants.TA_RIGHT)) == GdiConstants.TA_RIGHT)
                {
                    tx -= (width - lpdx[lpdx.Length - 1]);
                }
                else if ((align & (GdiConstants.TA_LEFT | GdiConstants.TA_CENTER | GdiConstants.TA_RIGHT)) == GdiConstants.TA_CENTER)
                {
                    tx -= (width - lpdx[lpdx.Length - 1]) / 2;
                }

                _buffer.Clear();
                for (var i = 0; i < lpdx.Length; i++)
                {
                    if (i > 0)
                    {
                        _buffer.Append(' ');
                    }

                    _buffer.Append((int)_dc.ToAbsoluteX(tx));
                    tx += lpdx[i];
                }

                if ((align & (GdiConstants.TA_NOUPDATECP | GdiConstants.TA_UPDATECP)) == GdiConstants.TA_UPDATECP)
                {
                    _dc.MoveToEx(tx, y, null);
                }

                elem.SetAttribute("x", _buffer.ToString());
            }
            else
            {
                if (_dc.Font != null)
                {
                    width = Math.Abs(_dc.Font.FontSize * text.Length) / 2;
                }

                elem.SetAttribute("x", ax.ToString(CultureInfo.InvariantCulture));
            }
        }

        // y
        var ay = (int)_dc.ToAbsoluteY(y);
        var height = 0;
        if (vertical)
        {
            if (_dc.Font != null)
            {
                lpdx = Helper.FixTextDx(_dc.Font.Charset, text, lpdx);
            }

            _buffer.Clear();
            if (align == 0 && _dc.Font != null)
            {
                _buffer.Append(ay + (int)_dc.ToRelativeY(Math.Abs(_dc.Font.Height)));
            }
            else
            {
                _buffer.Append(ay);
            }

            if (lpdx?.Length > 0)
            {
                for (var i = 0; i < lpdx.Length - 1; i++)
                {
                    height += lpdx[i];
                }

                var ty = y;

                if ((align & (GdiConstants.TA_LEFT | GdiConstants.TA_CENTER | GdiConstants.TA_RIGHT)) == GdiConstants.TA_RIGHT)
                {
                    ty -= (height - lpdx[lpdx.Length - 1]);
                }
                else if ((align & (GdiConstants.TA_LEFT | GdiConstants.TA_CENTER | GdiConstants.TA_RIGHT)) == GdiConstants.TA_CENTER)
                {
                    ty -= (height - lpdx[lpdx.Length - 1]) / 2;
                }

                for (var i = 0; i < lpdx.Length; i++)
                {
                    _buffer.Append(' ');
                    _buffer.Append((int)_dc.ToAbsoluteY(ty));
                    ty += lpdx[i];
                }

                if ((align & (GdiConstants.TA_NOUPDATECP | GdiConstants.TA_UPDATECP)) == GdiConstants.TA_UPDATECP)
                {
                    _dc.MoveToEx(x, ty, null);
                }
            }
            else
            {
                if (_dc.Font != null)
                {
                    height = Math.Abs(_dc.Font.FontSize * text.Length) / 2;
                }
            }

            elem.SetAttribute("y", _buffer.ToString());
        }
        else
        {
            if (_dc.Font != null)
            {
                height = Math.Abs(_dc.Font.FontSize);
            }

            if (_compatible)
            {
                if ((align & (GdiConstants.TA_BOTTOM | GdiConstants.TA_TOP | GdiConstants.TA_BASELINE)) == GdiConstants.TA_TOP)
                {
                    elem.SetAttribute("y", (ay + (int)_dc.ToRelativeY(height * 0.88)).ToString(CultureInfo.InvariantCulture));
                }
                else if ((align & (GdiConstants.TA_BOTTOM | GdiConstants.TA_TOP | GdiConstants.TA_BASELINE)) == GdiConstants.TA_BOTTOM && rect != null)
                {
                    elem.SetAttribute("y", (ay + rect[3] - rect[1] + (int)_dc.ToRelativeY(height * 0.88)).ToString(CultureInfo.InvariantCulture));
                }
                else
                {
                    elem.SetAttribute("y", ay.ToString(CultureInfo.InvariantCulture));
                }
            }
            else
            {
                if ((align & (GdiConstants.TA_BOTTOM | GdiConstants.TA_TOP | GdiConstants.TA_BASELINE)) == GdiConstants.TA_BOTTOM && rect != null)
                {
                    elem.SetAttribute("y", (ay + rect[3] - rect[1] - (int)_dc.ToRelativeY(height)).ToString(CultureInfo.InvariantCulture));
                }
                else
                {
                    elem.SetAttribute("y", ay.ToString(CultureInfo.InvariantCulture));
                }
            }
        }

        XmlElement? bk = null;
        if (_dc.BkMode == GdiConstants.OPAQUE || (options & GdiConstants.ETO_OPAQUE) > 0)
        {
            if (rect == null && _dc.Font != null)
            {
                rect = new int[4];
                if (vertical)
                {
                    if ((align & (GdiConstants.TA_BOTTOM | GdiConstants.TA_TOP | GdiConstants.TA_BASELINE)) == GdiConstants.TA_BOTTOM)
                    {
                        rect[0] = x - width;
                    }
                    else if ((align & (GdiConstants.TA_BOTTOM | GdiConstants.TA_TOP | GdiConstants.TA_BASELINE)) == GdiConstants.TA_BASELINE)
                    {
                        rect[0] = x - (int)(width * 0.85);
                    }
                    else
                    {
                        rect[0] = x;
                    }

                    if ((align & (GdiConstants.TA_LEFT | GdiConstants.TA_RIGHT | GdiConstants.TA_CENTER)) == GdiConstants.TA_RIGHT)
                    {
                        rect[1] = y - height;
                    }
                    else if ((align & (GdiConstants.TA_LEFT | GdiConstants.TA_RIGHT | GdiConstants.TA_CENTER)) == GdiConstants.TA_CENTER)
                    {
                        rect[1] = y - height / 2;
                    }
                    else
                    {
                        rect[1] = y;
                    }
                }
                else
                {
                    if ((align & (GdiConstants.TA_LEFT | GdiConstants.TA_RIGHT | GdiConstants.TA_CENTER)) == GdiConstants.TA_RIGHT)
                    {
                        rect[0] = x - width;
                    }
                    else if ((align & (GdiConstants.TA_LEFT | GdiConstants.TA_RIGHT | GdiConstants.TA_CENTER)) == GdiConstants.TA_CENTER)
                    {
                        rect[0] = x - width / 2;
                    }
                    else
                    {
                        rect[0] = x;
                    }

                    if ((align & (GdiConstants.TA_BOTTOM | GdiConstants.TA_TOP | GdiConstants.TA_BASELINE)) == GdiConstants.TA_BOTTOM)
                    {
                        rect[1] = y - height;
                    }
                    else if ((align & (GdiConstants.TA_BOTTOM | GdiConstants.TA_TOP | GdiConstants.TA_BASELINE)) == GdiConstants.TA_BASELINE)
                    {
                        rect[1] = y - (int)(height * 0.85);
                    }
                    else
                    {
                        rect[1] = y;
                    }
                }

                rect[2] = rect[0] + width;
                rect[3] = rect[1] + height;
            }

            if (rect != null)
            {
                bk = _doc.CreateElement("rect", "http://www.w3.org/2000/svg");
                bk.SetAttribute("x", ((int)_dc.ToAbsoluteX(rect[0])).ToString(CultureInfo.InvariantCulture));
                bk.SetAttribute("y", ((int)_dc.ToAbsoluteY(rect[1])).ToString(CultureInfo.InvariantCulture));
                bk.SetAttribute("width", ((int)_dc.ToRelativeX(rect[2] - rect[0])).ToString(CultureInfo.InvariantCulture));
                bk.SetAttribute("height", ((int)_dc.ToRelativeY(rect[3] - rect[1])).ToString(CultureInfo.InvariantCulture));
                bk.SetAttribute("fill", SvgObject.ToColor(_dc.BkColor));
            }
        }

        XmlElement? clip = null;
        if ((options & GdiConstants.ETO_CLIPPED) > 0 && rect != null)
        {
            var name = "clipPath" + (_clipPathNo++);
            clip = _doc.CreateElement("clipPath", "http://www.w3.org/2000/svg");
            clip.SetAttribute("id", name);

            var clipRect = _doc.CreateElement("rect", "http://www.w3.org/2000/svg");
            clipRect.SetAttribute("x", ((int)_dc.ToAbsoluteX(rect[0])).ToString(CultureInfo.InvariantCulture));
            clipRect.SetAttribute("y", ((int)_dc.ToAbsoluteY(rect[1])).ToString(CultureInfo.InvariantCulture));
            clipRect.SetAttribute("width", ((int)_dc.ToRelativeX(rect[2] - rect[0])).ToString(CultureInfo.InvariantCulture));
            clipRect.SetAttribute("height", ((int)_dc.ToRelativeY(rect[3] - rect[1])).ToString(CultureInfo.InvariantCulture));

            clip.AppendChild(clipRect);
            elem.SetAttribute("clip-path", $"url(#{name})");
        }

        string str;
        if (_dc.Font != null)
        {
            str = Helper.ConvertString(text, _dc.Font.Charset);
        }
        else
        {
            str = Helper.ConvertString(text, GdiFontConstants.DEFAULT_CHARSET);
        }

        if (_dc.Font?.Lang != null)
        {
            elem.SetAttribute("xml:lang", _dc.Font.Lang);
        }

        elem.SetAttribute("xml:space", "preserve");
        AppendText(elem, str);

        if (bk != null || clip != null)
        {
            var g = _doc.CreateElement("g", "http://www.w3.org/2000/svg");
            if (bk != null)
            {
                g.AppendChild(bk);
            }

            if (clip != null)
            {
                g.AppendChild(clip);
            }

            g.AppendChild(elem);
            elem = g;
        }

        if (escapement != 0)
        {
            elem.SetAttribute("transform", $"rotate({-escapement / 10.0}, {ax}, {ay})");
        }

        _parentNode.AppendChild(elem);
    }

    public void FillRgn(IGdiRegion rgn, IGdiBrush brush)
    {
        if (rgn == null)
        {
            return;
        }

        var elem = _doc.CreateElement("use", "http://www.w3.org/2000/svg");
        elem.SetAttribute("href", $"url(#{_nameMap[rgn]})", "http://www.w3.org/1999/xlink");
        elem.SetAttribute("class", GetClassString(brush));
        var sbrush = (SvgBrush)brush;
        if (sbrush.Style == GdiBrushConstants.BS_HATCHED)
        {
            var id = "pattern" + (_patternNo++);
            elem.SetAttribute("fill", $"url(#{id})");
            _defsNode.AppendChild(sbrush.CreateFillPattern(id)!);
        }

        _parentNode.AppendChild(elem);
    }

    public void FloodFill(int x, int y, int color)
    {
        // Not implemented
    }

    public void FrameRgn(IGdiRegion rgn, IGdiBrush brush, int w, int h)
    {
        // Not implemented
    }

    public void IntersectClipRect(int left, int top, int right, int bottom)
    {
        // Not implemented
    }

    public void InvertRgn(IGdiRegion rgn)
    {
        if (rgn == null)
        {
            return;
        }

        var elem = _doc.CreateElement("use", "http://www.w3.org/2000/svg");
        elem.SetAttribute("href", $"url(#{_nameMap[rgn]})", "http://www.w3.org/1999/xlink");
        var ropFilter = _dc.GetRopFilter(GdiConstants.DSTINVERT);
        if (ropFilter != null)
        {
            elem.SetAttribute("filter", ropFilter);
        }

        _parentNode.AppendChild(elem);
    }

    public void LineTo(int ex, int ey)
    {
        var elem = _doc.CreateElement("line", "http://www.w3.org/2000/svg");
        if (_dc.Pen != null)
        {
            elem.SetAttribute("class", GetClassString(_dc.Pen));
        }

        elem.SetAttribute("fill", "none");

        elem.SetAttribute("x1", ((int)_dc.ToAbsoluteX(_dc.CurrentX)).ToString(CultureInfo.InvariantCulture));
        elem.SetAttribute("y1", ((int)_dc.ToAbsoluteY(_dc.CurrentY)).ToString(CultureInfo.InvariantCulture));
        elem.SetAttribute("x2", ((int)_dc.ToAbsoluteX(ex)).ToString(CultureInfo.InvariantCulture));
        elem.SetAttribute("y2", ((int)_dc.ToAbsoluteY(ey)).ToString(CultureInfo.InvariantCulture));
        _parentNode.AppendChild(elem);

        _dc.MoveToEx(ex, ey, null);
    }

    public void MoveToEx(int x, int y, Point? old)
    {
        _dc.MoveToEx(x, y, old);
    }

    public void OffsetClipRgn(int x, int y)
    {
        _dc.OffsetClipRgn(x, y);
        var mask = _dc.Mask;
        if (mask != null)
        {
            mask = (XmlElement)mask.CloneNode(true);
            var name = "mask" + (_maskNo++);
            mask.SetAttribute("id", name);
            if (_dc.OffsetClipX != 0 || _dc.OffsetClipY != 0)
            {
                mask.SetAttribute("transform", $"translate({_dc.OffsetClipX},{_dc.OffsetClipY})");
            }

            _defsNode.AppendChild(mask);

            if (!_parentNode.HasChildNodes)
            {
                _doc.DocumentElement!.RemoveChild(_parentNode);
            }

            _parentNode = _doc.CreateElement("g", "http://www.w3.org/2000/svg");
            _parentNode.SetAttribute("mask", name);
            _doc.DocumentElement!.AppendChild(_parentNode);

            _dc.Mask = mask;
        }
    }

    public void OffsetViewportOrgEx(int x, int y, Point? point)
    {
        _dc.OffsetViewportOrgEx(x, y, point);
    }

    public void OffsetWindowOrgEx(int x, int y, Point? point)
    {
        _dc.OffsetWindowOrgEx(x, y, point);
    }

    public void PaintRgn(IGdiRegion rgn)
    {
        FillRgn(rgn, _dc.Brush!);
    }

    public void PatBlt(int x, int y, int width, int height, long rop)
    {
        var elem = _doc.CreateElement("rect", "http://www.w3.org/2000/svg");

        var brush = _dc.Brush;
        if (brush != null)
        {
            elem.SetAttribute("class", GetClassString(brush));
            if (brush.Style == GdiBrushConstants.BS_HATCHED)
            {
                var id = "pattern" + (_patternNo++);
                elem.SetAttribute("fill", $"url(#{id})");
                _defsNode.AppendChild(brush.CreateFillPattern(id)!);
            }
        }
        else
        {
            elem.SetAttribute("fill", "none");
        }

        elem.SetAttribute("stroke", "none");
        elem.SetAttribute("x", ((int)_dc.ToAbsoluteX(x)).ToString(CultureInfo.InvariantCulture));
        elem.SetAttribute("y", ((int)_dc.ToAbsoluteY(y)).ToString(CultureInfo.InvariantCulture));
        elem.SetAttribute("width", ((int)_dc.ToRelativeX(width)).ToString(CultureInfo.InvariantCulture));
        elem.SetAttribute("height", ((int)_dc.ToRelativeY(height)).ToString(CultureInfo.InvariantCulture));

        var ropFilter = _dc.GetRopFilter(rop);
        if (ropFilter != null)
        {
            elem.SetAttribute("filter", ropFilter);
        }

        _parentNode.AppendChild(elem);
    }

    public void Pie(int sxr, int syr, int exr, int eyr, int sxa, int sya, int exa, int eya)
    {
        var rx = Math.Abs(exr - sxr) / 2.0;
        var ry = Math.Abs(eyr - syr) / 2.0;
        if (rx <= 0 || ry <= 0)
        {
            return;
        }

        var cx = Math.Min(sxr, exr) + rx;
        var cy = Math.Min(syr, eyr) + ry;

        XmlElement elem;
        if (sxa == exa && sya == eya)
        {
            if (Math.Abs(rx - ry) < 0.0001)
            {
                elem = _doc.CreateElement("circle", "http://www.w3.org/2000/svg");
                elem.SetAttribute("cx", _dc.ToAbsoluteX(cx).ToString(CultureInfo.InvariantCulture));
                elem.SetAttribute("cy", _dc.ToAbsoluteY(cy).ToString(CultureInfo.InvariantCulture));
                elem.SetAttribute("r", _dc.ToRelativeX(rx).ToString(CultureInfo.InvariantCulture));
            }
            else
            {
                elem = _doc.CreateElement("ellipse", "http://www.w3.org/2000/svg");
                elem.SetAttribute("cx", _dc.ToAbsoluteX(cx).ToString(CultureInfo.InvariantCulture));
                elem.SetAttribute("cy", _dc.ToAbsoluteY(cy).ToString(CultureInfo.InvariantCulture));
                elem.SetAttribute("rx", _dc.ToRelativeX(rx).ToString(CultureInfo.InvariantCulture));
                elem.SetAttribute("ry", _dc.ToRelativeY(ry).ToString(CultureInfo.InvariantCulture));
            }
        }
        else
        {
            var sa = Math.Atan2((sya - cy) * rx, (sxa - cx) * ry);
            var sx = rx * Math.Cos(sa);
            var sy = ry * Math.Sin(sa);

            var ea = Math.Atan2((eya - cy) * rx, (exa - cx) * ry);
            var ex = rx * Math.Cos(ea);
            var ey = ry * Math.Sin(ea);

            var a = Math.Atan2((ex - sx) * (-sy) - (ey - sy) * (-sx), (ex - sx) * (-sx) + (ey - sy) * (-sy));

            elem = _doc.CreateElement("path", "http://www.w3.org/2000/svg");
            elem.SetAttribute("d",
                $"M {_dc.ToAbsoluteX(cx)},{_dc.ToAbsoluteY(cy)} L {_dc.ToAbsoluteX(sx + cx)},{_dc.ToAbsoluteY(sy + cy)} A {_dc.ToRelativeX(rx)},{_dc.ToRelativeY(ry)} 0 {(a > 0 ? "1" : "0")} 0 {_dc.ToAbsoluteX(ex + cx)},{_dc.ToAbsoluteY(ey + cy)} Z");
        }

        if (_dc.Pen != null || _dc.Brush != null)
        {
            elem.SetAttribute("class", GetClassString(_dc.Pen, _dc.Brush));
            if (_dc.Brush?.Style == GdiBrushConstants.BS_HATCHED)
            {
                var id = "pattern" + (_patternNo++);
                elem.SetAttribute("fill", $"url(#{id})");
                _defsNode.AppendChild(_dc.Brush.CreateFillPattern(id)!);
            }
        }

        _parentNode.AppendChild(elem);
    }

    public void Polygon(Point[] points)
    {
        var elem = _doc.CreateElement("polygon", "http://www.w3.org/2000/svg");

        if (_dc.Pen != null || _dc.Brush != null)
        {
            elem.SetAttribute("class", GetClassString(_dc.Pen, _dc.Brush));
            if (_dc.Brush?.Style == GdiBrushConstants.BS_HATCHED)
            {
                var id = "pattern" + (_patternNo++);
                elem.SetAttribute("fill", $"url(#{id})");
                _defsNode.AppendChild(_dc.Brush.CreateFillPattern(id)!);
            }

            if (_dc.PolyFillMode == GdiConstants.WINDING)
            {
                elem.SetAttribute("fill-rule", "nonzero");
            }
        }

        _buffer.Clear();
        for (var i = 0; i < points.Length; i++)
        {
            if (i != 0)
            {
                _buffer.Append(' ');
            }

            _buffer.Append((int)_dc.ToAbsoluteX(points[i].X)).Append(',');
            _buffer.Append((int)_dc.ToAbsoluteY(points[i].Y));
        }

        elem.SetAttribute("points", _buffer.ToString());
        _parentNode.AppendChild(elem);
    }

    public void Polyline(Point[] points)
    {
        var elem = _doc.CreateElement("polyline", "http://www.w3.org/2000/svg");
        if (_dc.Pen != null)
        {
            elem.SetAttribute("class", GetClassString(_dc.Pen));
        }

        elem.SetAttribute("fill", "none");

        _buffer.Clear();
        for (var i = 0; i < points.Length; i++)
        {
            if (i != 0)
            {
                _buffer.Append(' ');
            }

            _buffer.Append((int)_dc.ToAbsoluteX(points[i].X)).Append(',');
            _buffer.Append((int)_dc.ToAbsoluteY(points[i].Y));
        }

        elem.SetAttribute("points", _buffer.ToString());
        _parentNode.AppendChild(elem);
    }

    public void PolyPolygon(Point[][] points)
    {
        var elem = _doc.CreateElement("path", "http://www.w3.org/2000/svg");

        if (_dc.Pen != null || _dc.Brush != null)
        {
            elem.SetAttribute("class", GetClassString(_dc.Pen, _dc.Brush));
            if (_dc.Brush?.Style == GdiBrushConstants.BS_HATCHED)
            {
                var id = "pattern" + (_patternNo++);
                elem.SetAttribute("fill", $"url(#{id})");
                _defsNode.AppendChild(_dc.Brush.CreateFillPattern(id)!);
            }

            if (_dc.PolyFillMode == GdiConstants.WINDING)
            {
                elem.SetAttribute("fill-rule", "nonzero");
            }
        }

        _buffer.Clear();
        for (var i = 0; i < points.Length; i++)
        {
            if (i != 0)
            {
                _buffer.Append(' ');
            }

            for (var j = 0; j < points[i].Length; j++)
            {
                if (j == 0)
                {
                    _buffer.Append("M ");
                }
                else if (j == 1)
                {
                    _buffer.Append(" L ");
                }

                _buffer.Append((int)_dc.ToAbsoluteX(points[i][j].X)).Append(',');
                _buffer.Append((int)_dc.ToAbsoluteY(points[i][j].Y)).Append(' ');
                if (j == points[i].Length - 1)
                {
                    _buffer.Append('z');
                }
            }
        }

        elem.SetAttribute("d", _buffer.ToString());
        _parentNode.AppendChild(elem);
    }

    public void RealizePalette()
    {
        // Not implemented
    }

    public void RestoreDC(int savedDC)
    {
        var limit = savedDC < 0 ? -savedDC : _saveDC.Count - savedDC;
        for (var i = 0; i < limit && _saveDC.Count > 0; i++)
        {
            _dc = _saveDC[^1];
            _saveDC.RemoveAt(_saveDC.Count - 1);
        }

        if (!_parentNode.HasChildNodes)
        {
            _doc.DocumentElement!.RemoveChild(_parentNode);
        }

        _parentNode = _doc.CreateElement("g", "http://www.w3.org/2000/svg");
        var mask = _dc.Mask;
        if (mask != null)
        {
            _parentNode.SetAttribute("mask", $"url(#{mask.GetAttribute("id")})");
        }

        _doc.DocumentElement!.AppendChild(_parentNode);
    }

    public void Rectangle(int sx, int sy, int ex, int ey)
    {
        var elem = _doc.CreateElement("rect", "http://www.w3.org/2000/svg");

        if (_dc.Pen != null || _dc.Brush != null)
        {
            elem.SetAttribute("class", GetClassString(_dc.Pen, _dc.Brush));
            if (_dc.Brush?.Style == GdiBrushConstants.BS_HATCHED)
            {
                var id = "pattern" + (_patternNo++);
                elem.SetAttribute("fill", $"url(#{id})");
                _defsNode.AppendChild(_dc.Brush.CreateFillPattern(id)!);
            }
        }

        elem.SetAttribute("x", ((int)_dc.ToAbsoluteX(sx)).ToString(CultureInfo.InvariantCulture));
        elem.SetAttribute("y", ((int)_dc.ToAbsoluteY(sy)).ToString(CultureInfo.InvariantCulture));
        elem.SetAttribute("width", ((int)_dc.ToRelativeX(ex - sx)).ToString(CultureInfo.InvariantCulture));
        elem.SetAttribute("height", ((int)_dc.ToRelativeY(ey - sy)).ToString(CultureInfo.InvariantCulture));
        _parentNode.AppendChild(elem);
    }

    public void ResizePalette(IGdiPalette palette)
    {
        // Not implemented
    }

    public void RoundRect(int sx, int sy, int ex, int ey, int rw, int rh)
    {
        var elem = _doc.CreateElement("rect", "http://www.w3.org/2000/svg");

        if (_dc.Pen != null || _dc.Brush != null)
        {
            elem.SetAttribute("class", GetClassString(_dc.Pen, _dc.Brush));
            if (_dc.Brush?.Style == GdiBrushConstants.BS_HATCHED)
            {
                var id = "pattern" + (_patternNo++);
                elem.SetAttribute("fill", $"url(#{id})");
                _defsNode.AppendChild(_dc.Brush.CreateFillPattern(id)!);
            }
        }

        elem.SetAttribute("x", ((int)_dc.ToAbsoluteX(sx)).ToString(CultureInfo.InvariantCulture));
        elem.SetAttribute("y", ((int)_dc.ToAbsoluteY(sy)).ToString(CultureInfo.InvariantCulture));
        elem.SetAttribute("width", ((int)_dc.ToRelativeX(ex - sx)).ToString(CultureInfo.InvariantCulture));
        elem.SetAttribute("height", ((int)_dc.ToRelativeY(ey - sy)).ToString(CultureInfo.InvariantCulture));
        elem.SetAttribute("rx", ((int)_dc.ToRelativeX(rw)).ToString(CultureInfo.InvariantCulture));
        elem.SetAttribute("ry", ((int)_dc.ToRelativeY(rh)).ToString(CultureInfo.InvariantCulture));
        _parentNode.AppendChild(elem);
    }

    public void SaveDC()
    {
        _saveDC.Add((SvgDc)_dc.Clone());
    }

    public void ScaleViewportExtEx(int x, int xd, int y, int yd, Size? old)
    {
        _dc.ScaleViewportExtEx(x, xd, y, yd, old);
    }

    public void ScaleWindowExtEx(int x, int xd, int y, int yd, Size? old)
    {
        _dc.ScaleWindowExtEx(x, xd, y, yd, old);
    }

    public void SelectClipRgn(IGdiRegion? rgn)
    {
        if (!_parentNode.HasChildNodes)
        {
            _doc.DocumentElement!.RemoveChild(_parentNode);
        }

        _parentNode = _doc.CreateElement("g", "http://www.w3.org/2000/svg");

        if (rgn != null)
        {
            var mask = _doc.CreateElement("mask", "http://www.w3.org/2000/svg");
            mask.SetAttribute("id", "mask" + (_maskNo++));

            if (_dc.OffsetClipX != 0 || _dc.OffsetClipY != 0)
            {
                mask.SetAttribute("transform", $"translate({_dc.OffsetClipX},{_dc.OffsetClipY})");
            }

            _defsNode.AppendChild(mask);

            var clip = _doc.CreateElement("use", "http://www.w3.org/2000/svg");
            clip.SetAttribute("href", $"url(#{_nameMap[rgn]})", "http://www.w3.org/1999/xlink");
            clip.SetAttribute("fill", "white");

            mask.AppendChild(clip);

            _parentNode.SetAttribute("mask", $"url(#{mask.GetAttribute("id")})");
        }

        _doc.DocumentElement!.AppendChild(_parentNode);
    }

    public void SelectObject(IGdiObject obj)
    {
        if (obj is SvgBrush brush)
        {
            _dc.Brush = brush;
        }
        else if (obj is SvgFont font)
        {
            _dc.Font = font;
        }
        else if (obj is SvgPen pen)
        {
            _dc.Pen = pen;
        }
    }

    public void SelectPalette(IGdiPalette palette, bool mode)
    {
        // Not implemented
    }

    public void SetBkColor(int color)
    {
        _dc.BkColor = color;
    }

    public void SetBkMode(int mode)
    {
        _dc.BkMode = mode;
    }

    public void SetDIBitsToDevice(int dx, int dy, int dw, int dh, int sx, int sy,
        int startscan, int scanlines, byte[] image, int colorUse)
    {
        StretchDIBits(dx, dy, dw, dh, sx, sy, dw, dh, image, colorUse, GdiConstants.SRCCOPY);
    }

    public void SetLayout(long layout)
    {
        _dc.Layout = layout;
    }

    public void SetMapMode(int mode)
    {
        _dc.SetMapMode(mode);
    }

    public void SetMapperFlags(long flags)
    {
        _dc.MapperFlags = flags;
    }

    public void SetPaletteEntries(IGdiPalette palette, int startIndex, int[] entries)
    {
        // Not implemented
    }

    public void SetPixel(int x, int y, int color)
    {
        var elem = _doc.CreateElement("rect", "http://www.w3.org/2000/svg");
        elem.SetAttribute("stroke", "none");
        elem.SetAttribute("fill", SvgPen.ToColor(color));
        elem.SetAttribute("x", ((int)_dc.ToAbsoluteX(x)).ToString(CultureInfo.InvariantCulture));
        elem.SetAttribute("y", ((int)_dc.ToAbsoluteY(y)).ToString(CultureInfo.InvariantCulture));
        elem.SetAttribute("width", ((int)_dc.ToRelativeX(1)).ToString(CultureInfo.InvariantCulture));
        elem.SetAttribute("height", ((int)_dc.ToRelativeY(1)).ToString(CultureInfo.InvariantCulture));
        _parentNode.AppendChild(elem);
    }

    public void SetPolyFillMode(int mode)
    {
        _dc.PolyFillMode = mode;
    }

    public void SetRelAbs(int mode)
    {
        _dc.RelAbs = mode;
    }

    public void SetROP2(int mode)
    {
        _dc.ROP2 = mode;
    }

    public void SetStretchBltMode(int mode)
    {
        _dc.StretchBltMode = mode;
    }

    public void SetTextAlign(int align)
    {
        _dc.TextAlign = align;
    }

    public void SetTextCharacterExtra(int extra)
    {
        _dc.TextCharacterExtra = extra;
    }

    public void SetTextColor(int color)
    {
        _dc.TextColor = color;
    }

    public void SetTextJustification(int breakExtra, int breakCount)
    {
        if (breakCount > 0)
        {
            _dc.TextSpace = Math.Abs((int)_dc.ToRelativeX(breakExtra)) / breakCount;
        }
    }

    public void SetViewportExtEx(int x, int y, Size? old)
    {
        _dc.SetViewportExtEx(x, y, old);
    }

    public void SetViewportOrgEx(int x, int y, Point? old)
    {
        _dc.SetViewportOrgEx(x, y, old);
    }

    public void SetWindowExtEx(int width, int height, Size? old)
    {
        _dc.SetWindowExtEx(width, height, old);
    }

    public void SetWindowOrgEx(int x, int y, Point? old)
    {
        _dc.SetWindowOrgEx(x, y, old);
    }

    public void StretchBlt(byte[] image, int dx, int dy, int dw, int dh, int sx, int sy,
        int sw, int sh, long rop)
    {
        DibStretchBlt(image, dx, dy, dw, dh, sx, sy, sw, sh, rop);
    }

    public void StretchDIBits(int dx, int dy, int dw, int dh, int sx, int sy,
        int sw, int sh, byte[] image, int usage, long rop)
    {
        BmpToSvg(image, dx, dy, dw, dh, sx, sy, sw, sh, usage, rop);
    }

    public void TextOut(int x, int y, byte[] text)
    {
        ExtTextOut(x, y, 0, null, text, null);
    }

    public void Footer()
    {
        var root = _doc.DocumentElement!;
        if (!root.HasAttribute("width") && _dc.WindowWidth != 0)
        {
            root.SetAttribute("width", Math.Abs(_dc.WindowWidth).ToString(CultureInfo.InvariantCulture));
        }

        if (!root.HasAttribute("height") && _dc.WindowHeight != 0)
        {
            root.SetAttribute("height", Math.Abs(_dc.WindowHeight).ToString(CultureInfo.InvariantCulture));
        }

        if (_dc.WindowWidth != 0 && _dc.WindowHeight != 0)
        {
            root.SetAttribute("viewBox", $"0 0 {Math.Abs(_dc.WindowWidth)} {Math.Abs(_dc.WindowHeight)}");
            root.SetAttribute("preserveAspectRatio", "none");
        }

        root.SetAttribute("stroke-linecap", "round");
        root.SetAttribute("fill-rule", "evenodd");

        if (!_styleNode.HasChildNodes)
        {
            root.RemoveChild(_styleNode);
        }
        else
        {
            _styleNode.InsertBefore(_doc.CreateTextNode("\n"), _styleNode.FirstChild);
        }

        if (!_defsNode.HasChildNodes)
        {
            root.RemoveChild(_defsNode);
        }
    }

    private string GetClassString(IGdiObject? obj1, IGdiObject? obj2)
    {
        var name1 = GetClassString(obj1);
        var name2 = GetClassString(obj2);
        if (!string.IsNullOrEmpty(name1) && !string.IsNullOrEmpty(name2))
        {
            return name1 + " " + name2;
        }

        if (!string.IsNullOrEmpty(name1))
        {
            return name1;
        }

        if (!string.IsNullOrEmpty(name2))
        {
            return name2;
        }

        return "";
    }

    private string GetClassString(IGdiObject? style)
    {
        if (style == null)
        {
            return "";
        }

        return _nameMap.TryGetValue(style, out var name) ? name : "";
    }

    private void AppendText(XmlElement elem, string str)
    {
        if (_compatible)
        {
            str = Regex.Replace(str, @"\r\n|[\t\r\n ]", "\u00A0");
        }

        var font = _dc.Font;

        if (ReplaceSymbolFont && font != null)
        {
            if (string.Equals("Symbol", font.FaceName, StringComparison.OrdinalIgnoreCase))
            {
                var state = 0; // 0: default, 1: serif, 2: sans-serif
                var start = 0;
                var ca = str.ToCharArray();

                for (var i = 0; i < ca.Length; i++)
                {
                    var nstate = state;
                    switch (ca[i])
                    {
                        case '"':
                            ca[i] = '\u2200';
                            nstate = 1;
                            break;
                        case '$':
                            ca[i] = '\u2203';
                            nstate = 1;
                            break;
                        case '\'':
                            ca[i] = '\u220D';
                            nstate = 1;
                            break;
                        case '*':
                            ca[i] = '\u2217';
                            nstate = 1;
                            break;
                        case '-':
                            ca[i] = '\u2212';
                            nstate = 1;
                            break;
                        case '@':
                            ca[i] = '\u2245';
                            nstate = 1;
                            break;
                        case 'A':
                            ca[i] = '\u0391';
                            nstate = 1;
                            break;
                        case 'B':
                            ca[i] = '\u0392';
                            nstate = 1;
                            break;
                        case 'C':
                            ca[i] = '\u03A7';
                            nstate = 1;
                            break;
                        case 'D':
                            ca[i] = '\u0394';
                            nstate = 1;
                            break;
                        case 'E':
                            ca[i] = '\u0395';
                            nstate = 1;
                            break;
                        case 'F':
                            ca[i] = '\u03A6';
                            nstate = 1;
                            break;
                        case 'G':
                            ca[i] = '\u0393';
                            nstate = 1;
                            break;
                        case 'H':
                            ca[i] = '\u0397';
                            nstate = 1;
                            break;
                        case 'I':
                            ca[i] = '\u0399';
                            nstate = 1;
                            break;
                        case 'J':
                            ca[i] = '\u03D1';
                            nstate = 1;
                            break;
                        case 'K':
                            ca[i] = '\u039A';
                            nstate = 1;
                            break;
                        case 'L':
                            ca[i] = '\u039B';
                            nstate = 1;
                            break;
                        case 'M':
                            ca[i] = '\u039C';
                            nstate = 1;
                            break;
                        case 'N':
                            ca[i] = '\u039D';
                            nstate = 1;
                            break;
                        case 'O':
                            ca[i] = '\u039F';
                            nstate = 1;
                            break;
                        case 'P':
                            ca[i] = '\u03A0';
                            nstate = 1;
                            break;
                        case 'Q':
                            ca[i] = '\u0398';
                            nstate = 1;
                            break;
                        case 'R':
                            ca[i] = '\u03A1';
                            nstate = 1;
                            break;
                        case 'S':
                            ca[i] = '\u03A3';
                            nstate = 1;
                            break;
                        case 'T':
                            ca[i] = '\u03A4';
                            nstate = 1;
                            break;
                        case 'U':
                            ca[i] = '\u03A5';
                            nstate = 1;
                            break;
                        case 'V':
                            ca[i] = '\u03C3';
                            nstate = 1;
                            break;
                        case 'W':
                            ca[i] = '\u03A9';
                            nstate = 1;
                            break;
                        case 'X':
                            ca[i] = '\u039E';
                            nstate = 1;
                            break;
                        case 'Y':
                            ca[i] = '\u03A8';
                            nstate = 1;
                            break;
                        case 'Z':
                            ca[i] = '\u0396';
                            nstate = 1;
                            break;
                        case '\\':
                            ca[i] = '\u2234';
                            nstate = 1;
                            break;
                        case '^':
                            ca[i] = '\u22A5';
                            nstate = 1;
                            break;
                        case '`':
                            ca[i] = '\uF8E5';
                            nstate = 1;
                            break;
                        case 'a':
                            ca[i] = '\u03B1';
                            nstate = 1;
                            break;
                        case 'b':
                            ca[i] = '\u03B2';
                            nstate = 1;
                            break;
                        case 'c':
                            ca[i] = '\u03C7';
                            nstate = 1;
                            break;
                        case 'd':
                            ca[i] = '\u03B4';
                            nstate = 1;
                            break;
                        case 'e':
                            ca[i] = '\u03B5';
                            nstate = 1;
                            break;
                        case 'f':
                            ca[i] = '\u03C6';
                            nstate = 1;
                            break;
                        case 'g':
                            ca[i] = '\u03B3';
                            nstate = 1;
                            break;
                        case 'h':
                            ca[i] = '\u03B7';
                            nstate = 1;
                            break;
                        case 'i':
                            ca[i] = '\u03B9';
                            nstate = 1;
                            break;
                        case 'j':
                            ca[i] = '\u03D5';
                            nstate = 1;
                            break;
                        case 'k':
                            ca[i] = '\u03BA';
                            nstate = 1;
                            break;
                        case 'l':
                            ca[i] = '\u03BB';
                            nstate = 1;
                            break;
                        case 'm':
                            ca[i] = '\u03BC';
                            nstate = 1;
                            break;
                        case 'n':
                            ca[i] = '\u03BD';
                            nstate = 1;
                            break;
                        case 'o':
                            ca[i] = '\u03BF';
                            nstate = 1;
                            break;
                        case 'p':
                            ca[i] = '\u03C0';
                            nstate = 1;
                            break;
                        case 'q':
                            ca[i] = '\u03B8';
                            nstate = 1;
                            break;
                        case 'r':
                            ca[i] = '\u03C1';
                            nstate = 1;
                            break;
                        case 's':
                            ca[i] = '\u03C3';
                            nstate = 1;
                            break;
                        case 't':
                            ca[i] = '\u03C4';
                            nstate = 1;
                            break;
                        case 'u':
                            ca[i] = '\u03C5';
                            nstate = 1;
                            break;
                        case 'v':
                            ca[i] = '\u03D6';
                            nstate = 1;
                            break;
                        case 'w':
                            ca[i] = '\u03C9';
                            nstate = 1;
                            break;
                        case 'x':
                            ca[i] = '\u03BE';
                            nstate = 1;
                            break;
                        case 'y':
                            ca[i] = '\u03C8';
                            nstate = 1;
                            break;
                        case 'z':
                            ca[i] = '\u03B6';
                            nstate = 1;
                            break;
                        case '~':
                            ca[i] = '\u223C';
                            nstate = 1;
                            break;
                        case '\u00A0':
                            ca[i] = '\u20AC';
                            nstate = 1;
                            break;
                        case '\u00A1':
                            ca[i] = '\u03D2';
                            nstate = 1;
                            break;
                        case '\u00A2':
                            ca[i] = '\u2032';
                            nstate = 1;
                            break;
                        case '\u00A3':
                            ca[i] = '\u2264';
                            nstate = 1;
                            break;
                        case '\u00A4':
                            ca[i] = '\u2044';
                            nstate = 1;
                            break;
                        case '\u00A5':
                            ca[i] = '\u221E';
                            nstate = 1;
                            break;
                        case '\u00A6':
                            ca[i] = '\u0192';
                            nstate = 1;
                            break;
                        case '\u00A7':
                            ca[i] = '\u2663';
                            nstate = 1;
                            break;
                        case '\u00A8':
                            ca[i] = '\u2666';
                            nstate = 1;
                            break;
                        case '\u00A9':
                            ca[i] = '\u2665';
                            nstate = 1;
                            break;
                        case '\u00AA':
                            ca[i] = '\u2660';
                            nstate = 1;
                            break;
                        case '\u00AB':
                            ca[i] = '\u2194';
                            nstate = 1;
                            break;
                        case '\u00AC':
                            ca[i] = '\u2190';
                            nstate = 1;
                            break;
                        case '\u00AD':
                            ca[i] = '\u2191';
                            nstate = 1;
                            break;
                        case '\u00AE':
                            ca[i] = '\u2192';
                            nstate = 1;
                            break;
                        case '\u00AF':
                            ca[i] = '\u2193';
                            nstate = 1;
                            break;
                        case '\u00B2':
                            ca[i] = '\u2033';
                            nstate = 1;
                            break;
                        case '\u00B3':
                            ca[i] = '\u2265';
                            nstate = 1;
                            break;
                        case '\u00B4':
                            ca[i] = '\u00D7';
                            nstate = 1;
                            break;
                        case '\u00B5':
                            ca[i] = '\u221D';
                            nstate = 1;
                            break;
                        case '\u00B6':
                            ca[i] = '\u2202';
                            nstate = 1;
                            break;
                        case '\u00B7':
                            ca[i] = '\u2022';
                            nstate = 1;
                            break;
                        case '\u00B8':
                            ca[i] = '\u00F7';
                            nstate = 1;
                            break;
                        case '\u00B9':
                            ca[i] = '\u2260';
                            nstate = 1;
                            break;
                        case '\u00BA':
                            ca[i] = '\u2261';
                            nstate = 1;
                            break;
                        case '\u00BB':
                            ca[i] = '\u2248';
                            nstate = 1;
                            break;
                        case '\u00BC':
                            ca[i] = '\u2026';
                            nstate = 1;
                            break;
                        case '\u00BD':
                            ca[i] = '\u23D0';
                            nstate = 1;
                            break;
                        case '\u00BE':
                            ca[i] = '\u23AF';
                            nstate = 1;
                            break;
                        case '\u00BF':
                            ca[i] = '\u21B5';
                            nstate = 1;
                            break;
                        case '\u00C0':
                            ca[i] = '\u2135';
                            nstate = 1;
                            break;
                        case '\u00C1':
                            ca[i] = '\u2111';
                            nstate = 1;
                            break;
                        case '\u00C2':
                            ca[i] = '\u211C';
                            nstate = 1;
                            break;
                        case '\u00C3':
                            ca[i] = '\u2118';
                            nstate = 1;
                            break;
                        case '\u00C4':
                            ca[i] = '\u2297';
                            nstate = 1;
                            break;
                        case '\u00C5':
                            ca[i] = '\u2295';
                            nstate = 1;
                            break;
                        case '\u00C6':
                            ca[i] = '\u2205';
                            nstate = 1;
                            break;
                        case '\u00C7':
                            ca[i] = '\u2229';
                            nstate = 1;
                            break;
                        case '\u00C8':
                            ca[i] = '\u222A';
                            nstate = 1;
                            break;
                        case '\u00C9':
                            ca[i] = '\u2283';
                            nstate = 1;
                            break;
                        case '\u00CA':
                            ca[i] = '\u2287';
                            nstate = 1;
                            break;
                        case '\u00CB':
                            ca[i] = '\u2284';
                            nstate = 1;
                            break;
                        case '\u00CC':
                            ca[i] = '\u2282';
                            nstate = 1;
                            break;
                        case '\u00CD':
                            ca[i] = '\u2286';
                            nstate = 1;
                            break;
                        case '\u00CE':
                            ca[i] = '\u2208';
                            nstate = 1;
                            break;
                        case '\u00CF':
                            ca[i] = '\u2209';
                            nstate = 1;
                            break;
                        case '\u00D0':
                            ca[i] = '\u2220';
                            nstate = 1;
                            break;
                        case '\u00D1':
                            ca[i] = '\u2207';
                            nstate = 1;
                            break;
                        case '\u00D2':
                            ca[i] = '\u00AE';
                            nstate = 1;
                            break;
                        case '\u00D3':
                            ca[i] = '\u00A9';
                            nstate = 1;
                            break;
                        case '\u00D4':
                            ca[i] = '\u2122';
                            nstate = 1;
                            break;
                        case '\u00D5':
                            ca[i] = '\u220F';
                            nstate = 1;
                            break;
                        case '\u00D6':
                            ca[i] = '\u221A';
                            nstate = 1;
                            break;
                        case '\u00D7':
                            ca[i] = '\u22C5';
                            nstate = 1;
                            break;
                        case '\u00D8':
                            ca[i] = '\u00AC';
                            nstate = 1;
                            break;
                        case '\u00D9':
                            ca[i] = '\u2227';
                            nstate = 1;
                            break;
                        case '\u00DA':
                            ca[i] = '\u2228';
                            nstate = 1;
                            break;
                        case '\u00DB':
                            ca[i] = '\u21D4';
                            nstate = 1;
                            break;
                        case '\u00DC':
                            ca[i] = '\u21D0';
                            nstate = 1;
                            break;
                        case '\u00DD':
                            ca[i] = '\u21D1';
                            nstate = 1;
                            break;
                        case '\u00DE':
                            ca[i] = '\u21D2';
                            nstate = 1;
                            break;
                        case '\u00DF':
                            ca[i] = '\u21D3';
                            nstate = 1;
                            break;
                        case '\u00E0':
                            ca[i] = '\u25CA';
                            nstate = 1;
                            break;
                        case '\u00E1':
                            ca[i] = '\u3008';
                            nstate = 1;
                            break;
                        case '\u00E2':
                            ca[i] = '\u00AE';
                            nstate = 2;
                            break;
                        case '\u00E3':
                            ca[i] = '\u00A9';
                            nstate = 2;
                            break;
                        case '\u00E4':
                            ca[i] = '\u2122';
                            nstate = 2;
                            break;
                        case '\u00E5':
                            ca[i] = '\u2211';
                            nstate = 1;
                            break;
                        case '\u00E6':
                            ca[i] = '\u239B';
                            nstate = 1;
                            break;
                        case '\u00E7':
                            ca[i] = '\u239C';
                            nstate = 1;
                            break;
                        case '\u00E8':
                            ca[i] = '\u239D';
                            nstate = 1;
                            break;
                        case '\u00E9':
                            ca[i] = '\u23A1';
                            nstate = 1;
                            break;
                        case '\u00EA':
                            ca[i] = '\u23A2';
                            nstate = 1;
                            break;
                        case '\u00EB':
                            ca[i] = '\u23A3';
                            nstate = 1;
                            break;
                        case '\u00EC':
                            ca[i] = '\u23A7';
                            nstate = 1;
                            break;
                        case '\u00ED':
                            ca[i] = '\u23A8';
                            nstate = 1;
                            break;
                        case '\u00EE':
                            ca[i] = '\u23A9';
                            nstate = 1;
                            break;
                        case '\u00EF':
                            ca[i] = '\u23AA';
                            nstate = 1;
                            break;
                        case '\u00F0':
                            ca[i] = '\uF8FF';
                            nstate = 1;
                            break;
                        case '\u00F1':
                            ca[i] = '\u3009';
                            nstate = 1;
                            break;
                        case '\u00F2':
                            ca[i] = '\u222B';
                            nstate = 1;
                            break;
                        case '\u00F3':
                            ca[i] = '\u2320';
                            nstate = 1;
                            break;
                        case '\u00F4':
                            ca[i] = '\u23AE';
                            nstate = 1;
                            break;
                        case '\u00F5':
                            ca[i] = '\u2321';
                            nstate = 1;
                            break;
                        case '\u00F6':
                            ca[i] = '\u239E';
                            nstate = 1;
                            break;
                        case '\u00F7':
                            ca[i] = '\u239F';
                            nstate = 1;
                            break;
                        case '\u00F8':
                            ca[i] = '\u23A0';
                            nstate = 1;
                            break;
                        case '\u00F9':
                            ca[i] = '\u23A4';
                            nstate = 1;
                            break;
                        case '\u00FA':
                            ca[i] = '\u23A5';
                            nstate = 1;
                            break;
                        case '\u00FB':
                            ca[i] = '\u23A6';
                            nstate = 1;
                            break;
                        case '\u00FC':
                            ca[i] = '\u23AB';
                            nstate = 1;
                            break;
                        case '\u00FD':
                            ca[i] = '\u23AC';
                            nstate = 1;
                            break;
                        case '\u00FE':
                            ca[i] = '\u23AD';
                            nstate = 1;
                            break;
                        case '\u00FF':
                            ca[i] = '\u2192';
                            nstate = 1;
                            break;
                        default:
                            nstate = 0;
                            break;
                    }

                    if (nstate != state)
                    {
                        if (start < i)
                        {
                            var text = _doc.CreateTextNode(new string(ca, start, i - start));
                            
                            if (state == 0)
                            {
                                elem.AppendChild(text);
                            }
                            else if (state == 1)
                            {
                                var span = _doc.CreateElement("tspan");
                                span.SetAttribute("font-family", "serif");
                                span.AppendChild(text);
                                elem.AppendChild(span);
                            }
                            else if (state == 2)
                            {
                                var span = _doc.CreateElement("tspan");
                                span.SetAttribute("font-family", "sans-serif");
                                span.AppendChild(text);
                                elem.AppendChild(span);
                            }

                            start = i;
                        }

                        state = nstate;
                    }
                }

                if (start < ca.Length)
                {
                    var text = _doc.CreateTextNode(new string(ca, start, ca.Length - start));
                    if (state == 0)
                    {
                        elem.AppendChild(text);
                    }
                    else if (state == 1)
                    {
                        var span = _doc.CreateElement("tspan");
                        span.SetAttribute("font-family", "serif");
                        span.AppendChild(text);
                        elem.AppendChild(span);
                    }
                    else if (state == 2)
                    {
                        var span = _doc.CreateElement("tspan");
                        span.SetAttribute("font-family", "sans-serif");
                        span.AppendChild(text);
                        elem.AppendChild(span);
                    }
                }

                return;
            }
        }

        elem.AppendChild(_doc.CreateTextNode(str));
    }

    private void BmpToSvg(byte[] image, int dx, int dy, int dw, int dh, int sx, int sy,
        int sw, int sh, int usage, long rop)
    {
        if (image == null || image.Length == 0)
        {
            return;
        }

        var convertedImage = _imageConverter?.BmpToPng(DibToBmp(image), dh < 0);
        if (convertedImage == null || convertedImage.Length == 0)
        {
            return;
        }

        var data = "data:image/png;base64," + Convert.ToBase64String(convertedImage);

        var elem = _doc.CreateElement("image", "http://www.w3.org/2000/svg");
        var x = (int)_dc.ToAbsoluteX(dx);
        var y = (int)_dc.ToAbsoluteY(dy);
        var width = (int)_dc.ToRelativeX(dw);
        var height = (int)_dc.ToRelativeY(dh);

        if (width < 0 && height < 0)
        {
            elem.SetAttribute("transform", $"scale(-1, -1) translate({-x}, {-y})");
        }
        else if (width < 0)
        {
            elem.SetAttribute("transform", $"scale(-1, 1) translate({-x}, {y})");
        }
        else if (height < 0)
        {
            elem.SetAttribute("transform", $"scale(1, -1) translate({x}, {-y})");
        }
        else
        {
            elem.SetAttribute("x", x.ToString(CultureInfo.InvariantCulture));
            elem.SetAttribute("y", y.ToString(CultureInfo.InvariantCulture));
        }

        elem.SetAttribute("width", Math.Abs(width).ToString(CultureInfo.InvariantCulture));
        elem.SetAttribute("height", Math.Abs(height).ToString(CultureInfo.InvariantCulture));

        if (sx != 0 || sy != 0 || sw != dw || sh != dh)
        {
            elem.SetAttribute("viewBox", $"{sx} {sy} {sw} {sh}");
            elem.SetAttribute("preserveAspectRatio", "none");
        }

        var ropFilter = _dc.GetRopFilter(rop);
        if (ropFilter != null)
        {
            elem.SetAttribute("filter", ropFilter);
        }

        elem.SetAttribute("href", data, "http://www.w3.org/1999/xlink");
        _parentNode.AppendChild(elem);
    }

    private byte[] DibToBmp(byte[] dib)
    {
        var data = new byte[14 + dib.Length];

        // BitmapFileHeader
        data[0] = 0x42; // 'B'
        data[1] = 0x4d; // 'M'

        long bfSize = data.Length;
        data[2] = (byte)(bfSize & 0xff);
        data[3] = (byte)((bfSize >> 8) & 0xff);
        data[4] = (byte)((bfSize >> 16) & 0xff);
        data[5] = (byte)((bfSize >> 24) & 0xff);

        // reserved 1
        data[6] = 0x00;
        data[7] = 0x00;

        // reserved 2
        data[8] = 0x00;
        data[9] = 0x00;

        // offset
        long bfOffBits = 14;

        // BitmapInfoHeader
        long biSize = (dib[0] & 0xff) + ((dib[1] & 0xff) << 8)
                                      + ((dib[2] & 0xff) << 16) + ((dib[3] & 0xff) << 24);
        bfOffBits += biSize;

        var biBitCount = (dib[14] & 0xff) + ((dib[15] & 0xff) << 8);

        long clrUsed = (dib[32] & 0xff) + ((dib[33] & 0xff) << 8)
                                        + ((dib[34] & 0xff) << 16) + ((dib[35] & 0xff) << 24);

        switch (biBitCount)
        {
            case 1:
                bfOffBits += (clrUsed == 0L ? 2 : clrUsed) * 4;
                break;
            case 4:
                bfOffBits += (clrUsed == 0L ? 16 : clrUsed) * 4;
                break;
            case 8:
                bfOffBits += (clrUsed == 0L ? 256 : clrUsed) * 4;
                break;
        }

        data[10] = (byte)(bfOffBits & 0xff);
        data[11] = (byte)((bfOffBits >> 8) & 0xff);
        data[12] = (byte)((bfOffBits >> 16) & 0xff);
        data[13] = (byte)((bfOffBits >> 24) & 0xff);

        Array.Copy(dib, 0, data, 14, dib.Length);

        return data;
    }
}