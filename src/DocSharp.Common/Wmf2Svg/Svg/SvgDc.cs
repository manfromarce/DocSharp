using System;
using System.Xml;
using DocSharp.Wmf2Svg.Gdi;

namespace DocSharp.Wmf2Svg.Svg;

public sealed class SvgDc : ICloneable
{
    private readonly SvgGdi _gdi;

    private int _dpi = 1440;

    // window
    private int _wx;
    private int _wy;
    private int _ww;
    private int _wh;

    // window offset
    private int _wox;
    private int _woy;

    // window scale
    private double _wsx = 1.0;
    private double _wsy = 1.0;

    // mapping scale
    private double _mx = 1.0;
    private double _my = 1.0;

    // viewport
    private int _vx;
    private int _vy;
    private int _vw;
    private int _vh;

    // viewport offset
    private int _vox;
    private int _voy;

    // viewport scale
    private double _vsx = 1.0;
    private double _vsy = 1.0;

    // current location
    private int _cx;
    private int _cy;

    // clip offset
    private int _cox;
    private int _coy;

    private int _mapMode = GdiConstants.MM_TEXT;
    private int _bkColor = 0x00FFFFFF;
    private int _bkMode = GdiConstants.OPAQUE;
    private int _textColor;
    private int _textSpace;
    private int _textAlign = GdiConstants.TA_TOP | GdiConstants.TA_LEFT;
    private int _textDx;
    private int _polyFillMode = GdiConstants.ALTERNATE;
    private int _relAbsMode;
    private int _rop2Mode = GdiConstants.R2_COPYPEN;
    private int _stretchBltMode = GdiConstants.STRETCH_ANDSCANS;
    private long _layout;
    private long _mapperFlags;

    private SvgBrush? _brush;
    private SvgFont? _font;
    private SvgPen? _pen;

    private XmlElement? _mask;

    public SvgDc(SvgGdi gdi)
    {
        _gdi = gdi ?? throw new ArgumentNullException(nameof(gdi));
    }

    // Private copy constructor for cloning
    private SvgDc(SvgDc other)
    {
        _gdi = other._gdi;
        _dpi = other._dpi;
        _wx = other._wx;
        _wy = other._wy;
        _ww = other._ww;
        _wh = other._wh;
        _wox = other._wox;
        _woy = other._woy;
        _wsx = other._wsx;
        _wsy = other._wsy;
        _mx = other._mx;
        _my = other._my;
        _vx = other._vx;
        _vy = other._vy;
        _vw = other._vw;
        _vh = other._vh;
        _vox = other._vox;
        _voy = other._voy;
        _vsx = other._vsx;
        _vsy = other._vsy;
        _cx = other._cx;
        _cy = other._cy;
        _cox = other._cox;
        _coy = other._coy;
        _mapMode = other._mapMode;
        _bkColor = other._bkColor;
        _bkMode = other._bkMode;
        _textColor = other._textColor;
        _textSpace = other._textSpace;
        _textAlign = other._textAlign;
        _textDx = other._textDx;
        _polyFillMode = other._polyFillMode;
        _relAbsMode = other._relAbsMode;
        _rop2Mode = other._rop2Mode;
        _stretchBltMode = other._stretchBltMode;
        _layout = other._layout;
        _mapperFlags = other._mapperFlags;
        _brush = other._brush;
        _font = other._font;
        _pen = other._pen;
        _mask = other._mask;
    }

    public int Dpi
    {
        get => _dpi;
        set => _dpi = value;
    }

    public void SetWindowOrgEx(int x, int y, Point? old)
    {
        if (old != null)
        {
            old.X = _wx;
            old.Y = _wy;
        }

        _wx = x;
        _wy = y;
    }

    public void SetWindowExtEx(int width, int height, Size? old)
    {
        if (old != null)
        {
            old.Width = _ww;
            old.Height = _wh;
        }

        _ww = width;
        _wh = height;
    }

    public void OffsetWindowOrgEx(int x, int y, Point? old)
    {
        if (old != null)
        {
            old.X = _wox;
            old.Y = _woy;
        }

        _wox += x;
        _woy += y;
    }

    public void ScaleWindowExtEx(int x, int xd, int y, int yd, Size? old)
    {
        // TODO
        _wsx = (_wsx * x) / xd;
        _wsy = (_wsy * y) / yd;
    }

    public int WindowX => _wx;
    public int WindowY => _wy;
    public int WindowWidth => _ww;
    public int WindowHeight => _wh;

    public void SetViewportOrgEx(int x, int y, Point? old)
    {
        if (old != null)
        {
            old.X = _vx;
            old.Y = _vy;
        }

        _vx = x;
        _vy = y;
    }

    public void SetViewportExtEx(int width, int height, Size? old)
    {
        if (old != null)
        {
            old.Width = _vw;
            old.Height = _vh;
        }

        _vw = width;
        _vh = height;
    }

    public void OffsetViewportOrgEx(int x, int y, Point? old)
    {
        if (old != null)
        {
            old.X = _vox;
            old.Y = _voy;
        }

        _vox = x;
        _voy = y;
    }

    public void ScaleViewportExtEx(int x, int xd, int y, int yd, Size? old)
    {
        // TODO
        _vsx = (_vsx * x) / xd;
        _vsy = (_vsy * y) / yd;
    }

    public void OffsetClipRgn(int x, int y)
    {
        _cox = x;
        _coy = y;
    }

    public int MapMode => _mapMode;

    public void SetMapMode(int mode)
    {
        _mapMode = mode;
        switch (mode)
        {
            case GdiConstants.MM_HIENGLISH:
                _mx = 0.09;
                _my = -0.09;
                break;
            case GdiConstants.MM_LOENGLISH:
                _mx = 0.9;
                _my = -0.9;
                break;
            case GdiConstants.MM_HIMETRIC:
                _mx = 0.03543307;
                _my = -0.03543307;
                break;
            case GdiConstants.MM_LOMETRIC:
                _mx = 0.3543307;
                _my = -0.3543307;
                break;
            case GdiConstants.MM_TWIPS:
                _mx = 0.0625;
                _my = -0.0625;
                break;
            default:
                _mx = 1.0;
                _my = 1.0;
                break;
        }
    }

    public int CurrentX => _cx;
    public int CurrentY => _cy;
    public int OffsetClipX => _cox;
    public int OffsetClipY => _coy;

    public void MoveToEx(int x, int y, Point? old)
    {
        if (old != null)
        {
            old.X = _cx;
            old.Y = _cy;
        }

        _cx = x;
        _cy = y;
    }

    public double ToAbsoluteX(double x)
    {
        // TODO Handle Viewport
        return ((_ww >= 0) ? 1 : -1) * (_mx * x - (_wx + _wox)) / _wsx;
    }

    public double ToAbsoluteY(double y)
    {
        // TODO Handle Viewport
        return ((_wh >= 0) ? 1 : -1) * (_my * y - (_wy + _woy)) / _wsy;
    }

    public double ToRelativeX(double x)
    {
        // TODO Handle Viewport
        return ((_ww >= 0) ? 1 : -1) * (_mx * x) / _wsx;
    }

    public double ToRelativeY(double y)
    {
        // TODO Handle Viewport
        return ((_wh >= 0) ? 1 : -1) * (_my * y) / _wsy;
    }

    public void SetDpi(int dpi)
    {
        _dpi = (dpi > 0) ? dpi : 1440;
    }

    public int BkColor
    {
        get => _bkColor;
        set => _bkColor = value;
    }

    public int BkMode
    {
        get => _bkMode;
        set => _bkMode = value;
    }

    public int TextColor
    {
        get => _textColor;
        set => _textColor = value;
    }

    public int PolyFillMode
    {
        get => _polyFillMode;
        set => _polyFillMode = value;
    }

    public int RelAbs
    {
        get => _relAbsMode;
        set => _relAbsMode = value;
    }

    public int ROP2
    {
        get => _rop2Mode;
        set => _rop2Mode = value;
    }

    public int StretchBltMode
    {
        get => _stretchBltMode;
        set => _stretchBltMode = value;
    }

    public int TextSpace
    {
        get => _textSpace;
        set => _textSpace = value;
    }

    public int TextAlign
    {
        get => _textAlign;
        set => _textAlign = value;
    }

    public int TextCharacterExtra
    {
        get => _textDx;
        set => _textDx = value;
    }

    public long Layout
    {
        get => _layout;
        set => _layout = value;
    }

    public long MapperFlags
    {
        get => _mapperFlags;
        set => _mapperFlags = value;
    }

    public SvgBrush? Brush
    {
        get => _brush;
        set => _brush = value;
    }

    public SvgFont? Font
    {
        get => _font;
        set => _font = value;
    }

    public SvgPen? Pen
    {
        get => _pen;
        set => _pen = value;
    }

    public XmlElement? Mask
    {
        get => _mask;
        set => _mask = value;
    }

    public string? GetRopFilter(long rop)
    {
        string? name = null;
        var doc = _gdi.Document;

        if (rop == GdiConstants.BLACKNESS)
        {
            name = "BLACKNESS_FILTER";
            var filter = doc.GetElementById(name);
            if (filter == null)
            {
                filter = _gdi.Document.CreateElement("filter");
                filter.SetAttribute("id", name);

                var feColorMatrix = doc.CreateElement("feColorMatrix");
                feColorMatrix.SetAttribute("type", "matrix");
                feColorMatrix.SetAttribute("in", "SourceGraphic");
                feColorMatrix.SetAttribute("values", "0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 1 0");
                filter.AppendChild(feColorMatrix);

                _gdi.DefsElement?.AppendChild(filter);
            }
        }
        else if (rop == GdiConstants.NOTSRCERASE)
        {
            name = "NOTSRCERASE_FILTER";
            var filter = doc.GetElementById(name);
            if (filter == null)
            {
                filter = _gdi.Document.CreateElement("filter");
                filter.SetAttribute("id", name);

                var feComposite = doc.CreateElement("feComposite");
                feComposite.SetAttribute("in", "SourceGraphic");
                feComposite.SetAttribute("in2", "BackgroundImage");
                feComposite.SetAttribute("operator", "arithmetic");
                feComposite.SetAttribute("k1", "1");
                feComposite.SetAttribute("result", "result0");
                filter.AppendChild(feComposite);

                var feColorMatrix = doc.CreateElement("feColorMatrix");
                feColorMatrix.SetAttribute("in", "result0");
                feColorMatrix.SetAttribute("values", "-1 0 0 0 1 0 -1 0 0 1 0 0 -1 0 1 0 0 0 1 0");
                filter.AppendChild(feColorMatrix);

                _gdi.DefsElement?.AppendChild(filter);
            }
        }
        else if (rop == GdiConstants.NOTSRCCOPY)
        {
            name = "NOTSRCCOPY_FILTER";
            var filter = doc.GetElementById(name);
            if (filter == null)
            {
                filter = _gdi.Document.CreateElement("filter");
                filter.SetAttribute("id", name);

                var feColorMatrix = doc.CreateElement("feColorMatrix");
                feColorMatrix.SetAttribute("type", "matrix");
                feColorMatrix.SetAttribute("in", "SourceGraphic");
                feColorMatrix.SetAttribute("values", "-1 0 0 0 1 0 -1 0 0 1 0 0 -1 0 1 0 0 0 1 0");
                filter.AppendChild(feColorMatrix);

                _gdi.DefsElement?.AppendChild(filter);
            }
        }
        else if (rop == GdiConstants.SRCERASE)
        {
            name = "SRCERASE_FILTER";
            var filter = doc.GetElementById(name);
            if (filter == null)
            {
                filter = _gdi.Document.CreateElement("filter");
                filter.SetAttribute("id", name);

                var feColorMatrix = doc.CreateElement("feColorMatrix");
                feColorMatrix.SetAttribute("type", "matrix");
                feColorMatrix.SetAttribute("in", "BackgroundImage");
                feColorMatrix.SetAttribute("values", "-1 0 0 0 1 0 -1 0 0 1 0 0 -1 0 1 0 0 0 1 0");
                feColorMatrix.SetAttribute("result", "result0");
                filter.AppendChild(feColorMatrix);

                var feComposite = doc.CreateElement("feComposite");
                feComposite.SetAttribute("in", "SourceGraphic");
                feComposite.SetAttribute("in2", "result0");
                feComposite.SetAttribute("operator", "arithmetic");
                feComposite.SetAttribute("k2", "1");
                feComposite.SetAttribute("k3", "1");
                filter.AppendChild(feComposite);

                _gdi.DefsElement?.AppendChild(filter);
            }
        }
        else if (rop == GdiConstants.PATINVERT)
        {
            // TODO
        }
        else if (rop == GdiConstants.SRCINVERT)
        {
            // TODO
        }
        else if (rop == GdiConstants.DSTINVERT)
        {
            name = "DSTINVERT_FILTER";
            var filter = doc.GetElementById(name);
            if (filter == null)
            {
                filter = _gdi.Document.CreateElement("filter");
                filter.SetAttribute("id", name);

                var feColorMatrix = doc.CreateElement("feColorMatrix");
                feColorMatrix.SetAttribute("type", "matrix");
                feColorMatrix.SetAttribute("in", "BackgroundImage");
                feColorMatrix.SetAttribute("values", "-1 0 0 0 1 0 -1 0 0 1 0 0 -1 0 1 0 0 0 1 0");
                filter.AppendChild(feColorMatrix);

                _gdi.DefsElement?.AppendChild(filter);
            }
        }
        else if (rop == GdiConstants.SRCAND)
        {
            name = "SRCAND_FILTER";
            var filter = doc.GetElementById(name);
            if (filter == null)
            {
                filter = _gdi.Document.CreateElement("filter");
                filter.SetAttribute("id", name);

                var feComposite = doc.CreateElement("feComposite");
                feComposite.SetAttribute("in", "SourceGraphic");
                feComposite.SetAttribute("in2", "BackgroundImage");
                feComposite.SetAttribute("operator", "arithmetic");
                feComposite.SetAttribute("k1", "1");
                filter.AppendChild(feComposite);

                _gdi.DefsElement?.AppendChild(filter);
            }
        }
        else if (rop == GdiConstants.MERGEPAINT)
        {
            name = "MERGEPAINT_FILTER";
            var filter = doc.GetElementById(name);
            if (filter == null)
            {
                filter = _gdi.Document.CreateElement("filter");
                filter.SetAttribute("id", name);

                var feColorMatrix = doc.CreateElement("feColorMatrix");
                feColorMatrix.SetAttribute("type", "matrix");
                feColorMatrix.SetAttribute("in", "SourceGraphic");
                feColorMatrix.SetAttribute("values", "-1 0 0 0 1 0 -1 0 0 1 0 0 -1 0 1 0 0 0 1 0");
                feColorMatrix.SetAttribute("result", "result0");
                filter.AppendChild(feColorMatrix);

                var feComposite = doc.CreateElement("feComposite");
                feComposite.SetAttribute("in", "result0");
                feComposite.SetAttribute("in2", "BackgroundImage");
                feComposite.SetAttribute("operator", "arithmetic");
                feComposite.SetAttribute("k1", "1");
                filter.AppendChild(feComposite);

                _gdi.DefsElement?.AppendChild(filter);
            }
        }
        else if (rop == GdiConstants.MERGECOPY)
        {
            // TODO
        }
        else if (rop == GdiConstants.SRCPAINT)
        {
            name = "SRCPAINT_FILTER";
            var filter = doc.GetElementById(name);
            if (filter == null)
            {
                filter = _gdi.Document.CreateElement("filter");
                filter.SetAttribute("id", name);

                var feComposite = doc.CreateElement("feComposite");
                feComposite.SetAttribute("in", "SourceGraphic");
                feComposite.SetAttribute("in2", "BackgroundImage");
                feComposite.SetAttribute("operator", "arithmetic");
                feComposite.SetAttribute("k2", "1");
                feComposite.SetAttribute("k3", "1");
                filter.AppendChild(feComposite);

                _gdi.DefsElement?.AppendChild(filter);
            }
        }
        else if (rop == GdiConstants.PATCOPY)
        {
            // TODO
        }
        else if (rop == GdiConstants.PATPAINT)
        {
            // TODO
        }
        else if (rop == GdiConstants.WHITENESS)
        {
            name = "WHITENESS_FILTER";
            var filter = doc.GetElementById(name);
            if (filter == null)
            {
                filter = _gdi.Document.CreateElement("filter");
                filter.SetAttribute("id", name);

                var feColorMatrix = doc.CreateElement("feColorMatrix");
                feColorMatrix.SetAttribute("type", "matrix");
                feColorMatrix.SetAttribute("in", "SourceGraphic");
                feColorMatrix.SetAttribute("values", "1 0 0 0 1 0 1 0 0 1 0 0 1 0 1 0 0 0 1 0");
                filter.AppendChild(feColorMatrix);

                _gdi.DefsElement?.AppendChild(filter);
            }
        }

        if (name != null)
        {
            if (doc.DocumentElement?.GetAttribute("enable-background") != "new")
            {
                doc.DocumentElement?.SetAttribute("enable-background", "new");
            }

            return "url(#" + name + ")";
        }

        return null;
    }

    public object Clone()
    {
        return new SvgDc(this);
    }

    public override string ToString()
    {
        return $"SvgDc [gdi={_gdi}, dpi={_dpi}, wx={_wx}, wy={_wy}, ww={_ww}, wh={_wh}, " +
               $"wox={_wox}, woy={_woy}, wsx={_wsx}, wsy={_wsy}, mx={_mx}, my={_my}, " +
               $"vx={_vx}, vy={_vy}, vw={_vw}, vh={_vh}, vox={_vox}, voy={_voy}, " +
               $"vsx={_vsx}, vsy={_vsy}, cx={_cx}, cy={_cy}, mapMode={_mapMode}, " +
               $"bkColor={_bkColor}, bkMode={_bkMode}, textColor={_textColor}, " +
               $"textSpace={_textSpace}, textAlign={_textAlign}, textDx={_textDx}, " +
               $"polyFillMode={_polyFillMode}, relAbsMode={_relAbsMode}, rop2Mode={_rop2Mode}, " +
               $"stretchBltMode={_stretchBltMode}, brush={_brush}, font={_font}, pen={_pen}]";
    }
}