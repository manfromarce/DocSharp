using System;
using DocSharp.Wmf2Svg.Gdi;

namespace DocSharp.Wmf2Svg.Wmf;

public sealed class WmfDc : ICloneable
{
    // window offset
    private int _wox;
    private int _woy;

    // window scale
    private double _wsx = 1.0;
    private double _wsy = 1.0;

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

    private int _textAlign = GdiConstants.TA_TOP | GdiConstants.TA_LEFT;

    private WmfBrush? _brush;
    private WmfFont? _font;
    private WmfPen? _pen;

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

    public int TextAlign
    {
        get => _textAlign;
        set => _textAlign = value;
    }

    public WmfBrush? Brush
    {
        get => _brush;
        set => _brush = value;
    }

    public WmfFont? Font
    {
        get => _font;
        set => _font = value;
    }

    public WmfPen? Pen
    {
        get => _pen;
        set => _pen = value;
    }

    public object Clone()
    {
        return MemberwiseClone();
    }
}