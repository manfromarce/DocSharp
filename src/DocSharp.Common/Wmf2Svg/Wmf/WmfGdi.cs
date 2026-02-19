using System;
using System.Collections.Generic;
using System.IO;
using DocSharp.Wmf2Svg.Gdi;

namespace DocSharp.Wmf2Svg.Wmf;

public sealed class WmfGdi : IGdi
{
    private byte[]? _placeableHeader;
    private byte[]? _header;

    private readonly List<IGdiObject?> _objects = new();
    private readonly List<byte[]> _records = new();

    private readonly WmfDc _dc = new();

    private WmfBrush? _defaultBrush;
    private WmfPen? _defaultPen;
    private WmfFont? _defaultFont;

    public WmfGdi()
    {
        _defaultBrush = (WmfBrush)CreateBrushIndirect(GdiBrushConstants.BS_SOLID, 0x00FFFFFF, 0);
        _defaultPen = (WmfPen)CreatePenIndirect(GdiPenConstants.PS_SOLID, 1, 0x00000000);
        _defaultFont = null;

        _dc.Brush = _defaultBrush;
        _dc.Pen = _defaultPen;
        _dc.Font = _defaultFont;
    }

    public void Write(Stream output)
    {
        Footer();
        if (_placeableHeader != null)
        {
            output.Write(_placeableHeader, 0, _placeableHeader.Length);
        }

        if (_header != null)
        {
            output.Write(_header, 0, _header.Length);
        }

        foreach (var record in _records)
        {
            output.Write(record, 0, record.Length);
        }

        output.Flush();
    }

    public void PlaceableHeader(int wsx, int wsy, int wex, int wey, int dpi)
    {
        var record = new byte[22];
        var pos = 0;
        pos = SetUint32(record, pos, 0x9AC6CDD7);
        pos = SetInt16(record, pos, 0x0000);
        pos = SetInt16(record, pos, wsx);
        pos = SetInt16(record, pos, wsy);
        pos = SetInt16(record, pos, wex);
        pos = SetInt16(record, pos, wey);
        pos = SetUint16(record, pos, dpi);
        pos = SetUint32(record, pos, 0x00000000);

        var checksum = 0;
        for (var i = 0; i < record.Length - 2; i += 2)
        {
            checksum ^= (0xFF & record[i]) | ((0xFF & record[i + 1]) << 8);
        }

        pos = SetUint16(record, pos, checksum);
        _placeableHeader = record;
    }

    public void Header()
    {
        var record = new byte[18];
        var pos = 0;
        pos = SetUint16(record, pos, 0x0001);
        pos = SetUint16(record, pos, 0x0009);
        pos = SetUint16(record, pos, 0x0300);
        pos = SetUint32(record, pos, 0x0000); // dummy size
        pos = SetUint16(record, pos, 0x0000); // dummy noObjects
        pos = SetUint32(record, pos, 0x0000); // dummy maxRecords
        pos = SetUint16(record, pos, 0x0000);
        _header = record;
    }

    public void AnimatePalette(IGdiPalette palette, int startIndex, int[] entries)
    {
        var record = new byte[22];
        var pos = 0;
        pos = SetUint32(record, pos, record.Length / 2);
        pos = SetUint16(record, pos, WmfConstants.RECORD_ANIMATE_PALETTE);
        pos = SetUint16(record, pos, entries.Length);
        pos = SetUint16(record, pos, startIndex);
        pos = SetUint16(record, pos, ((WmfPalette)palette).ID);
        for (var i = 0; i < entries.Length; i++)
        {
            pos = SetInt32(record, pos, entries[i]);
        }

        _records.Add(record);
    }

    public void Arc(int sxr, int syr, int exr, int eyr, int sxa, int sya, int exa, int eya)
    {
        var record = new byte[22];
        var pos = 0;
        pos = SetUint32(record, pos, record.Length / 2);
        pos = SetUint16(record, pos, WmfConstants.RECORD_ARC);
        pos = SetInt16(record, pos, eya);
        pos = SetInt16(record, pos, exa);
        pos = SetInt16(record, pos, sya);
        pos = SetInt16(record, pos, sxa);
        pos = SetInt16(record, pos, eyr);
        pos = SetInt16(record, pos, exr);
        pos = SetInt16(record, pos, syr);
        pos = SetInt16(record, pos, sxr);
        _records.Add(record);
    }

    public void BitBlt(byte[] image, int dx, int dy, int dw, int dh, int sx, int sy, long rop)
    {
        var record = new byte[22 + (image.Length + image.Length % 2)];
        var pos = 0;
        pos = SetUint32(record, pos, record.Length / 2);
        pos = SetUint16(record, pos, WmfConstants.RECORD_BIT_BLT);
        pos = SetUint32(record, pos, rop);
        pos = SetInt16(record, pos, sy);
        pos = SetInt16(record, pos, sx);
        pos = SetInt16(record, pos, dw);
        pos = SetInt16(record, pos, dh);
        pos = SetInt16(record, pos, dy);
        pos = SetInt16(record, pos, dx);
        pos = SetBytes(record, pos, image);
        if (image.Length % 2 == 1)
        {
            pos = SetByte(record, pos, 0);
        }

        _records.Add(record);
    }

    public void Chord(int sxr, int syr, int exr, int eyr, int sxa, int sya, int exa, int eya)
    {
        var record = new byte[22];
        var pos = 0;
        pos = SetUint32(record, pos, record.Length / 2);
        pos = SetUint16(record, pos, WmfConstants.RECORD_CHORD);
        pos = SetInt16(record, pos, eya);
        pos = SetInt16(record, pos, exa);
        pos = SetInt16(record, pos, sya);
        pos = SetInt16(record, pos, sxa);
        pos = SetInt16(record, pos, eyr);
        pos = SetInt16(record, pos, exr);
        pos = SetInt16(record, pos, syr);
        pos = SetInt16(record, pos, sxr);
        _records.Add(record);
    }

    public IGdiBrush CreateBrushIndirect(int style, int color, int hatch)
    {
        var record = new byte[14];
        var pos = 0;
        pos = SetUint32(record, pos, record.Length / 2);
        pos = SetUint16(record, pos, WmfConstants.RECORD_CREATE_BRUSH_INDIRECT);
        pos = SetUint16(record, pos, style);
        pos = SetInt32(record, pos, color);
        pos = SetUint16(record, pos, hatch);
        _records.Add(record);

        var brush = new WmfBrush(_objects.Count, style, color, hatch);
        _objects.Add(brush);
        return brush;
    }

    public IGdiFont CreateFontIndirect(int height, int width, int escapement,
        int orientation, int weight, bool italic, bool underline,
        bool strikeout, int charset, int outPrecision,
        int clipPrecision, int quality, int pitchAndFamily, byte[] faceName)
    {
        var record = new byte[24 + (faceName.Length + faceName.Length % 2)];
        var pos = 0;
        pos = SetUint32(record, pos, record.Length / 2);
        pos = SetUint16(record, pos, WmfConstants.RECORD_CREATE_FONT_INDIRECT);
        pos = SetInt16(record, pos, height);
        pos = SetInt16(record, pos, width);
        pos = SetInt16(record, pos, escapement);
        pos = SetInt16(record, pos, orientation);
        pos = SetInt16(record, pos, weight);
        pos = SetByte(record, pos, italic ? 0x01 : 0x00);
        pos = SetByte(record, pos, underline ? 0x01 : 0x00);
        pos = SetByte(record, pos, strikeout ? 0x01 : 0x00);
        pos = SetByte(record, pos, charset);
        pos = SetByte(record, pos, outPrecision);
        pos = SetByte(record, pos, clipPrecision);
        pos = SetByte(record, pos, quality);
        pos = SetByte(record, pos, pitchAndFamily);
        pos = SetBytes(record, pos, faceName);
        if (faceName.Length % 2 == 1)
        {
            pos = SetByte(record, pos, 0);
        }

        _records.Add(record);

        var font = new WmfFont(_objects.Count, height, width, escapement,
            orientation, weight, italic, underline, strikeout, charset, outPrecision,
            clipPrecision, quality, pitchAndFamily, faceName);
        _objects.Add(font);
        return font;
    }

    public IGdiPalette CreatePalette(int version, int[] palEntry)
    {
        var record = new byte[10 + palEntry.Length * 4];
        var pos = 0;
        pos = SetUint32(record, pos, record.Length / 2);
        pos = SetUint16(record, pos, WmfConstants.RECORD_CREATE_PALETTE);
        pos = SetUint16(record, pos, version);
        pos = SetUint16(record, pos, palEntry.Length);
        for (var i = 0; i < palEntry.Length; i++)
        {
            pos = SetInt32(record, pos, palEntry[i]);
        }

        _records.Add(record);

        var palette = new WmfPalette(_objects.Count, version, palEntry);
        _objects.Add(palette);
        return palette;
    }

    public IGdiPatternBrush CreatePatternBrush(byte[] image)
    {
        var record = new byte[6 + (image.Length + image.Length % 2)];
        var pos = 0;
        pos = SetUint32(record, pos, record.Length / 2);
        pos = SetUint16(record, pos, WmfConstants.RECORD_CREATE_PATTERN_BRUSH);
        pos = SetBytes(record, pos, image);
        if (image.Length % 2 == 1)
        {
            pos = SetByte(record, pos, 0);
        }

        _records.Add(record);

        var brush = new WmfPatternBrush(_objects.Count, image);
        _objects.Add(brush);
        return brush;
    }

    public IGdiPen CreatePenIndirect(int style, int width, int color)
    {
        var record = new byte[16];
        var pos = 0;
        pos = SetUint32(record, pos, record.Length / 2);
        pos = SetUint16(record, pos, WmfConstants.RECORD_CREATE_PEN_INDIRECT);
        pos = SetUint16(record, pos, style);
        pos = SetInt16(record, pos, width);
        pos = SetInt16(record, pos, 0);
        pos = SetInt32(record, pos, color);
        _records.Add(record);

        var pen = new WmfPen(_objects.Count, style, width, color);
        _objects.Add(pen);
        return pen;
    }

    public IGdiRegion CreateRectRgn(int left, int top, int right, int bottom)
    {
        var record = new byte[14];
        var pos = 0;
        pos = SetUint32(record, pos, record.Length / 2);
        pos = SetUint16(record, pos, WmfConstants.RECORD_CREATE_RECT_RGN);
        pos = SetInt16(record, pos, bottom);
        pos = SetInt16(record, pos, right);
        pos = SetInt16(record, pos, top);
        pos = SetInt16(record, pos, left);
        _records.Add(record);

        var rgn = new WmfRectRegion(_objects.Count, left, top, right, bottom);
        _objects.Add(rgn);
        return rgn;
    }

    public void DeleteObject(IGdiObject obj)
    {
        var record = new byte[8];
        var pos = 0;
        pos = SetUint32(record, pos, record.Length / 2);
        pos = SetUint16(record, pos, WmfConstants.RECORD_DELETE_OBJECT);
        pos = SetUint16(record, pos, ((WmfObject)obj).ID);
        _records.Add(record);

        _objects[((WmfObject)obj).ID] = null;

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
        var record = new byte[22 + (image.Length + image.Length % 2)];
        var pos = 0;
        pos = SetUint32(record, pos, record.Length / 2);
        pos = SetUint16(record, pos, WmfConstants.RECORD_DIB_BIT_BLT);
        pos = SetUint32(record, pos, rop);
        pos = SetInt16(record, pos, sy);
        pos = SetInt16(record, pos, sx);
        pos = SetInt16(record, pos, dw);
        pos = SetInt16(record, pos, dh);
        pos = SetInt16(record, pos, dy);
        pos = SetInt16(record, pos, dx);
        pos = SetBytes(record, pos, image);
        if (image.Length % 2 == 1)
        {
            pos = SetByte(record, pos, 0);
        }

        _records.Add(record);
    }

    public IGdiPatternBrush DibCreatePatternBrush(byte[] image, int usage)
    {
        var record = new byte[10 + (image.Length + image.Length % 2)];
        var pos = 0;
        pos = SetUint32(record, pos, record.Length / 2);
        pos = SetUint16(record, pos, WmfConstants.RECORD_DIB_CREATE_PATTERN_BRUSH);
        pos = SetInt32(record, pos, usage);
        pos = SetBytes(record, pos, image);
        if (image.Length % 2 == 1)
        {
            pos = SetByte(record, pos, 0);
        }

        _records.Add(record);

        // TODO usage
        var brush = new WmfPatternBrush(_objects.Count, image);
        _objects.Add(brush);
        return brush;
    }

    public void DibStretchBlt(byte[] image, int dx, int dy, int dw, int dh,
        int sx, int sy, int sw, int sh, long rop)
    {
        var record = new byte[26 + (image.Length + image.Length % 2)];
        var pos = 0;
        pos = SetUint32(record, pos, record.Length / 2);
        pos = SetUint16(record, pos, WmfConstants.RECORD_DIB_STRETCH_BLT);
        pos = SetUint32(record, pos, rop);
        pos = SetInt16(record, pos, sh);
        pos = SetInt16(record, pos, sw);
        pos = SetInt16(record, pos, sy);
        pos = SetInt16(record, pos, sx);
        pos = SetInt16(record, pos, dw);
        pos = SetInt16(record, pos, dh);
        pos = SetInt16(record, pos, dy);
        pos = SetInt16(record, pos, dx);
        pos = SetBytes(record, pos, image);
        if (image.Length % 2 == 1)
        {
            pos = SetByte(record, pos, 0);
        }

        _records.Add(record);
    }

    public void Ellipse(int sx, int sy, int ex, int ey)
    {
        var record = new byte[14];
        var pos = 0;
        pos = SetUint32(record, pos, record.Length / 2);
        pos = SetUint16(record, pos, WmfConstants.RECORD_ELLIPSE);
        pos = SetInt16(record, pos, ey);
        pos = SetInt16(record, pos, ex);
        pos = SetInt16(record, pos, sy);
        pos = SetInt16(record, pos, sx);
        _records.Add(record);
    }

    public void Escape(byte[] data)
    {
        var record = new byte[10 + (data.Length + data.Length % 2)];
        var pos = 0;
        pos = SetUint32(record, pos, record.Length / 2);
        pos = SetUint16(record, pos, WmfConstants.RECORD_ESCAPE);
        pos = SetBytes(record, pos, data);
        if (data.Length % 2 == 1)
        {
            pos = SetByte(record, pos, 0);
        }

        _records.Add(record);
    }

    public int ExcludeClipRect(int left, int top, int right, int bottom)
    {
        var record = new byte[14];
        var pos = 0;
        pos = SetUint32(record, pos, record.Length / 2);
        pos = SetUint16(record, pos, WmfConstants.RECORD_EXCLUDE_CLIP_RECT);
        pos = SetInt16(record, pos, bottom);
        pos = SetInt16(record, pos, right);
        pos = SetInt16(record, pos, top);
        pos = SetInt16(record, pos, left);
        _records.Add(record);

        // TODO
        return GdiRegionConstants.COMPLEXREGION;
    }

    public void ExtFloodFill(int x, int y, int color, int type)
    {
        var record = new byte[16];
        var pos = 0;
        pos = SetUint32(record, pos, record.Length / 2);
        pos = SetUint16(record, pos, WmfConstants.RECORD_EXT_FLOOD_FILL);
        pos = SetUint16(record, pos, type);
        pos = SetInt32(record, pos, color);
        pos = SetInt16(record, pos, y);
        pos = SetInt16(record, pos, x);
        _records.Add(record);
    }

    public void ExtTextOut(int x, int y, int options, int[]? rect, byte[] text, int[]? lpdx)
    {
        if (rect != null && rect.Length != 4)
        {
            throw new ArgumentException("rect must be 4 length.");
        }

        var dxLength = lpdx?.Length ?? 0;
        var record = new byte[14 + ((rect != null) ? 8 : 0) + (text.Length + text.Length % 2) + (dxLength * 2)];
        var pos = 0;
        pos = SetUint32(record, pos, record.Length / 2);
        pos = SetUint16(record, pos, WmfConstants.RECORD_EXT_TEXT_OUT);
        pos = SetInt16(record, pos, y);
        pos = SetInt16(record, pos, x);
        pos = SetInt16(record, pos, text.Length);
        pos = SetInt16(record, pos, options);
        if (rect != null)
        {
            pos = SetInt16(record, pos, rect[0]);
            pos = SetInt16(record, pos, rect[1]);
            pos = SetInt16(record, pos, rect[2]);
            pos = SetInt16(record, pos, rect[3]);
        }

        pos = SetBytes(record, pos, text);
        if (text.Length % 2 == 1)
        {
            pos = SetByte(record, pos, 0);
        }

        if (lpdx != null)
        {
            for (var i = 0; i < lpdx.Length; i++)
            {
                pos = SetInt16(record, pos, lpdx[i]);
            }
        }

        _records.Add(record);

        var vertical = false;
        if (_dc.Font != null)
        {
            if (_dc.Font.FaceName.StartsWith("@"))
            {
                vertical = true;
            }
        }

        var align = _dc.TextAlign;
        var width = 0;
        if (!vertical)
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

                for (var i = 0; i < lpdx.Length; i++)
                {
                    tx += lpdx[i];
                }

                if ((align & (GdiConstants.TA_NOUPDATECP | GdiConstants.TA_UPDATECP)) == GdiConstants.TA_UPDATECP)
                {
                    _dc.MoveToEx(tx, y, null);
                }
            }
        }
    }

    public void FillRgn(IGdiRegion rgn, IGdiBrush brush)
    {
        var record = new byte[10];
        var pos = 0;
        pos = SetUint32(record, pos, record.Length / 2);
        pos = SetUint16(record, pos, WmfConstants.RECORD_FLOOD_FILL);
        pos = SetUint16(record, pos, ((WmfBrush)brush).ID);
        pos = SetUint16(record, pos, ((WmfRegion)rgn).ID);
        _records.Add(record);
    }

    public void FloodFill(int x, int y, int color)
    {
        var record = new byte[16];
        var pos = 0;
        pos = SetUint32(record, pos, record.Length / 2);
        pos = SetUint16(record, pos, WmfConstants.RECORD_FLOOD_FILL);
        pos = SetInt32(record, pos, color);
        pos = SetInt16(record, pos, y);
        pos = SetInt16(record, pos, x);
        _records.Add(record);
    }

    public void FrameRgn(IGdiRegion rgn, IGdiBrush brush, int w, int h)
    {
        var record = new byte[14];
        var pos = 0;
        pos = SetUint32(record, pos, record.Length / 2);
        pos = SetUint16(record, pos, WmfConstants.RECORD_FRAME_RGN);
        pos = SetInt16(record, pos, h);
        pos = SetInt16(record, pos, w);
        pos = SetUint16(record, pos, ((WmfBrush)brush).ID);
        pos = SetUint16(record, pos, ((WmfRegion)rgn).ID);
        _records.Add(record);
    }

    public void IntersectClipRect(int left, int top, int right, int bottom)
    {
        var record = new byte[16];
        var pos = 0;
        pos = SetUint32(record, pos, record.Length / 2);
        pos = SetUint16(record, pos, WmfConstants.RECORD_INTERSECT_CLIP_RECT);
        pos = SetInt16(record, pos, bottom);
        pos = SetInt16(record, pos, right);
        pos = SetInt16(record, pos, top);
        pos = SetInt16(record, pos, left);
        _records.Add(record);
    }

    public void InvertRgn(IGdiRegion rgn)
    {
        var record = new byte[8];
        var pos = 0;
        pos = SetUint32(record, pos, record.Length / 2);
        pos = SetUint16(record, pos, WmfConstants.RECORD_INVERT_RGN);
        pos = SetUint16(record, pos, ((WmfRegion)rgn).ID);
        _records.Add(record);
    }

    public void LineTo(int ex, int ey)
    {
        var record = new byte[10];
        var pos = 0;
        pos = SetUint32(record, pos, record.Length / 2);
        pos = SetUint16(record, pos, WmfConstants.RECORD_LINE_TO);
        pos = SetInt16(record, pos, ey);
        pos = SetInt16(record, pos, ex);
        _records.Add(record);

        _dc.MoveToEx(ex, ey, null);
    }

    public void MoveToEx(int x, int y, Point? old)
    {
        var record = new byte[10];
        var pos = 0;
        pos = SetUint32(record, pos, record.Length / 2);
        pos = SetUint16(record, pos, WmfConstants.RECORD_MOVE_TO_EX);
        pos = SetInt16(record, pos, y);
        pos = SetInt16(record, pos, x);
        _records.Add(record);

        _dc.MoveToEx(x, y, old);
    }

    public void OffsetClipRgn(int x, int y)
    {
        var record = new byte[10];
        var pos = 0;
        pos = SetUint32(record, pos, record.Length / 2);
        pos = SetUint16(record, pos, WmfConstants.RECORD_OFFSET_CLIP_RGN);
        pos = SetInt16(record, pos, y);
        pos = SetInt16(record, pos, x);
        _records.Add(record);
    }

    public void OffsetViewportOrgEx(int x, int y, Point? point)
    {
        var record = new byte[10];
        var pos = 0;
        pos = SetUint32(record, pos, record.Length / 2);
        pos = SetUint16(record, pos, WmfConstants.RECORD_OFFSET_VIEWPORT_ORG_EX);
        pos = SetInt16(record, pos, y);
        pos = SetInt16(record, pos, x);
        _records.Add(record);

        _dc.OffsetViewportOrgEx(x, y, point);
    }

    public void OffsetWindowOrgEx(int x, int y, Point? point)
    {
        var record = new byte[10];
        var pos = 0;
        pos = SetUint32(record, pos, record.Length / 2);
        pos = SetUint16(record, pos, WmfConstants.RECORD_OFFSET_WINDOW_ORG_EX);
        pos = SetInt16(record, pos, y);
        pos = SetInt16(record, pos, x);
        _records.Add(record);

        _dc.OffsetWindowOrgEx(x, y, point);
    }

    public void PaintRgn(IGdiRegion rgn)
    {
        var record = new byte[8];
        var pos = 0;
        pos = SetUint32(record, pos, record.Length / 2);
        pos = SetUint16(record, pos, WmfConstants.RECORD_PAINT_RGN);
        pos = SetUint16(record, pos, ((WmfRegion)rgn).ID);
        _records.Add(record);
    }

    public void PatBlt(int x, int y, int width, int height, long rop)
    {
        var record = new byte[18];
        var pos = 0;
        pos = SetUint32(record, pos, record.Length / 2);
        pos = SetUint16(record, pos, WmfConstants.RECORD_PAT_BLT);
        pos = SetUint32(record, pos, rop);
        pos = SetInt16(record, pos, height);
        pos = SetInt16(record, pos, width);
        pos = SetInt16(record, pos, y);
        pos = SetInt16(record, pos, x);
        _records.Add(record);
    }

    public void Pie(int sxr, int syr, int exr, int eyr, int sxa, int sya, int exa, int eya)
    {
        var record = new byte[22];
        var pos = 0;
        pos = SetUint32(record, pos, record.Length / 2);
        pos = SetUint16(record, pos, WmfConstants.RECORD_PIE);
        pos = SetInt16(record, pos, eya);
        pos = SetInt16(record, pos, exa);
        pos = SetInt16(record, pos, sya);
        pos = SetInt16(record, pos, sxa);
        pos = SetInt16(record, pos, eyr);
        pos = SetInt16(record, pos, exr);
        pos = SetInt16(record, pos, syr);
        pos = SetInt16(record, pos, sxr);
        _records.Add(record);
    }

    public void Polygon(Point[] points)
    {
        var record = new byte[8 + points.Length * 4];
        var pos = 0;
        pos = SetUint32(record, pos, record.Length / 2);
        pos = SetUint16(record, pos, WmfConstants.RECORD_POLYGON);
        pos = SetInt16(record, pos, points.Length);
        for (var i = 0; i < points.Length; i++)
        {
            pos = SetInt16(record, pos, points[i].X);
            pos = SetInt16(record, pos, points[i].Y);
        }

        _records.Add(record);
    }

    public void Polyline(Point[] points)
    {
        var record = new byte[8 + points.Length * 4];
        var pos = 0;
        pos = SetUint32(record, pos, record.Length / 2);
        pos = SetUint16(record, pos, WmfConstants.RECORD_POLYLINE);
        pos = SetInt16(record, pos, points.Length);
        for (var i = 0; i < points.Length; i++)
        {
            pos = SetInt16(record, pos, points[i].X);
            pos = SetInt16(record, pos, points[i].Y);
        }

        _records.Add(record);
    }

    public void PolyPolygon(Point[][] points)
    {
        var length = 8;
        for (var i = 0; i < points.Length; i++)
        {
            length += 2 + points[i].Length * 4;
        }

        var record = new byte[length];
        var pos = 0;
        pos = SetUint32(record, pos, record.Length / 2);
        pos = SetUint16(record, pos, WmfConstants.RECORD_POLYLINE);
        pos = SetInt16(record, pos, points.Length);
        for (var i = 0; i < points.Length; i++)
        {
            pos = SetInt16(record, pos, points[i].Length);
        }

        for (var i = 0; i < points.Length; i++)
        {
            for (var j = 0; j < points[i].Length; j++)
            {
                pos = SetInt16(record, pos, points[i][j].X);
                pos = SetInt16(record, pos, points[i][j].Y);
            }
        }

        _records.Add(record);
    }

    public void RealizePalette()
    {
        var record = new byte[6];
        var pos = 0;
        pos = SetUint32(record, pos, record.Length / 2);
        pos = SetUint16(record, pos, WmfConstants.RECORD_REALIZE_PALETTE);
        _records.Add(record);
    }

    public void RestoreDC(int savedDC)
    {
        var record = new byte[8];
        var pos = 0;
        pos = SetUint32(record, pos, record.Length / 2);
        pos = SetUint16(record, pos, WmfConstants.RECORD_RESTORE_DC);
        pos = SetInt16(record, pos, savedDC);
        _records.Add(record);
    }

    public void Rectangle(int sx, int sy, int ex, int ey)
    {
        var record = new byte[14];
        var pos = 0;
        pos = SetUint32(record, pos, record.Length / 2);
        pos = SetUint16(record, pos, WmfConstants.RECORD_RECTANGLE);
        pos = SetInt16(record, pos, ey);
        pos = SetInt16(record, pos, ex);
        pos = SetInt16(record, pos, sy);
        pos = SetInt16(record, pos, sx);
        _records.Add(record);
    }

    public void ResizePalette(IGdiPalette palette)
    {
        var record = new byte[8];
        var pos = 0;
        pos = SetUint32(record, pos, record.Length / 2);
        pos = SetUint16(record, pos, WmfConstants.RECORD_REALIZE_PALETTE);
        pos = SetUint16(record, pos, ((WmfPalette)palette).ID);
        _records.Add(record);
    }

    public void RoundRect(int sx, int sy, int ex, int ey, int rw, int rh)
    {
        var record = new byte[18];
        var pos = 0;
        pos = SetUint32(record, pos, record.Length / 2);
        pos = SetUint16(record, pos, WmfConstants.RECORD_ROUND_RECT);
        pos = SetInt16(record, pos, rh);
        pos = SetInt16(record, pos, rw);
        pos = SetInt16(record, pos, ey);
        pos = SetInt16(record, pos, ex);
        pos = SetInt16(record, pos, sy);
        pos = SetInt16(record, pos, sx);
        _records.Add(record);
    }

    public void SaveDC()
    {
        var record = new byte[6];
        var pos = 0;
        pos = SetUint32(record, pos, record.Length / 2);
        pos = SetUint16(record, pos, WmfConstants.RECORD_SAVE_DC);
        _records.Add(record);
    }

    public void ScaleViewportExtEx(int x, int xd, int y, int yd, Size? old)
    {
        var record = new byte[14];
        var pos = 0;
        pos = SetUint32(record, pos, record.Length / 2);
        pos = SetUint16(record, pos, WmfConstants.RECORD_SCALE_VIEWPORT_EXT_EX);
        pos = SetInt16(record, pos, yd);
        pos = SetInt16(record, pos, y);
        pos = SetInt16(record, pos, xd);
        pos = SetInt16(record, pos, x);
        _records.Add(record);

        _dc.ScaleViewportExtEx(x, xd, y, yd, old);
    }

    public void ScaleWindowExtEx(int x, int xd, int y, int yd, Size? old)
    {
        var record = new byte[14];
        var pos = 0;
        pos = SetUint32(record, pos, record.Length / 2);
        pos = SetUint16(record, pos, WmfConstants.RECORD_SCALE_WINDOW_EXT_EX);
        pos = SetInt16(record, pos, yd);
        pos = SetInt16(record, pos, y);
        pos = SetInt16(record, pos, xd);
        pos = SetInt16(record, pos, x);
        _records.Add(record);

        _dc.ScaleWindowExtEx(x, xd, y, yd, old);
    }

    public void SelectClipRgn(IGdiRegion? rgn)
    {
        if (rgn == null)
        {
            return;
        }

        var record = new byte[8];
        var pos = 0;
        pos = SetUint32(record, pos, record.Length / 2);
        pos = SetUint16(record, pos, WmfConstants.RECORD_SELECT_CLIP_RGN);
        pos = SetUint16(record, pos, ((WmfRegion)rgn).ID);
        _records.Add(record);
    }

    public void SelectObject(IGdiObject obj)
    {
        var record = new byte[8];
        var pos = 0;
        pos = SetUint32(record, pos, record.Length / 2);
        pos = SetUint16(record, pos, WmfConstants.RECORD_SELECT_OBJECT);
        pos = SetUint16(record, pos, ((WmfObject)obj).ID);
        _records.Add(record);

        if (obj is WmfBrush brush)
        {
            _dc.Brush = brush;
        }
        else if (obj is WmfFont font)
        {
            _dc.Font = font;
        }
        else if (obj is WmfPen pen)
        {
            _dc.Pen = pen;
        }
    }

    public void SelectPalette(IGdiPalette palette, bool mode)
    {
        var record = new byte[8];
        var pos = 0;
        pos = SetUint32(record, pos, record.Length / 2);
        pos = SetUint16(record, pos, WmfConstants.RECORD_SELECT_PALETTE);
        pos = SetInt16(record, pos, mode ? 1 : 0);
        pos = SetUint16(record, pos, ((WmfPalette)palette).ID);
        _records.Add(record);
    }

    public void SetBkColor(int color)
    {
        var record = new byte[10];
        var pos = 0;
        pos = SetUint32(record, pos, record.Length / 2);
        pos = SetUint16(record, pos, WmfConstants.RECORD_SET_BK_COLOR);
        pos = SetInt32(record, pos, color);
        _records.Add(record);
    }

    public void SetBkMode(int mode)
    {
        var record = new byte[8];
        var pos = 0;
        pos = SetUint32(record, pos, record.Length / 2);
        pos = SetUint16(record, pos, WmfConstants.RECORD_SET_BK_MODE);
        pos = SetInt16(record, pos, mode);
        _records.Add(record);
    }

    public void SetDIBitsToDevice(int dx, int dy, int dw, int dh, int sx,
        int sy, int startscan, int scanlines, byte[] image, int colorUse)
    {
        var record = new byte[24 + (image.Length + image.Length % 2)];
        var pos = 0;
        pos = SetUint32(record, pos, record.Length / 2);
        pos = SetUint16(record, pos, WmfConstants.RECORD_SET_DIBITS_TO_DEVICE);
        pos = SetUint16(record, pos, colorUse);
        pos = SetUint16(record, pos, scanlines);
        pos = SetUint16(record, pos, startscan);
        pos = SetInt16(record, pos, sy);
        pos = SetInt16(record, pos, sx);
        pos = SetInt16(record, pos, dw);
        pos = SetInt16(record, pos, dh);
        pos = SetInt16(record, pos, dy);
        pos = SetInt16(record, pos, dx);
        pos = SetBytes(record, pos, image);
        if (image.Length % 2 == 1)
        {
            pos = SetByte(record, pos, 0);
        }

        _records.Add(record);
    }

    public void SetLayout(long layout)
    {
        var record = new byte[10];
        var pos = 0;
        pos = SetUint32(record, pos, record.Length / 2);
        pos = SetUint16(record, pos, WmfConstants.RECORD_SET_LAYOUT);
        pos = SetUint32(record, pos, layout);
        _records.Add(record);
    }

    public void SetMapMode(int mode)
    {
        var record = new byte[8];
        var pos = 0;
        pos = SetUint32(record, pos, record.Length / 2);
        pos = SetUint16(record, pos, WmfConstants.RECORD_SET_MAP_MODE);
        pos = SetInt16(record, pos, mode);
        _records.Add(record);
    }

    public void SetMapperFlags(long flags)
    {
        var record = new byte[10];
        var pos = 0;
        pos = SetUint32(record, pos, record.Length / 2);
        pos = SetUint16(record, pos, WmfConstants.RECORD_SET_MAPPER_FLAGS);
        pos = SetUint32(record, pos, flags);
        _records.Add(record);
    }

    public void SetPaletteEntries(IGdiPalette palette, int startIndex, int[] entries)
    {
        var record = new byte[6 + entries.Length * 4];
        var pos = 0;
        pos = SetUint32(record, pos, record.Length / 2);
        pos = SetUint16(record, pos, WmfConstants.RECORD_SET_PALETTE_ENTRIES);
        pos = SetUint16(record, pos, ((WmfPalette)palette).ID);
        pos = SetUint16(record, pos, entries.Length);
        pos = SetUint16(record, pos, startIndex);
        for (var i = 0; i < entries.Length; i++)
        {
            pos = SetInt32(record, pos, entries[i]);
        }

        _records.Add(record);
    }

    public void SetPixel(int x, int y, int color)
    {
        var record = new byte[14];
        var pos = 0;
        pos = SetUint32(record, pos, record.Length / 2);
        pos = SetUint16(record, pos, WmfConstants.RECORD_SET_PIXEL);
        pos = SetInt32(record, pos, color);
        pos = SetInt16(record, pos, y);
        pos = SetInt16(record, pos, x);
        _records.Add(record);
    }

    public void SetPolyFillMode(int mode)
    {
        var record = new byte[8];
        var pos = 0;
        pos = SetUint32(record, pos, record.Length / 2);
        pos = SetUint16(record, pos, WmfConstants.RECORD_SET_POLY_FILL_MODE);
        pos = SetInt16(record, pos, mode);
        _records.Add(record);
    }

    public void SetRelAbs(int mode)
    {
        var record = new byte[8];
        var pos = 0;
        pos = SetUint32(record, pos, record.Length / 2);
        pos = SetUint16(record, pos, WmfConstants.RECORD_SET_REL_ABS);
        pos = SetInt16(record, pos, mode);
        _records.Add(record);
    }

    public void SetROP2(int mode)
    {
        var record = new byte[8];
        var pos = 0;
        pos = SetUint32(record, pos, record.Length / 2);
        pos = SetUint16(record, pos, WmfConstants.RECORD_SET_ROP2);
        pos = SetInt16(record, pos, mode);
        _records.Add(record);
    }

    public void SetStretchBltMode(int mode)
    {
        var record = new byte[8];
        var pos = 0;
        pos = SetUint32(record, pos, record.Length / 2);
        pos = SetUint16(record, pos, WmfConstants.RECORD_SET_STRETCH_BLT_MODE);
        pos = SetInt16(record, pos, mode);
        _records.Add(record);
    }

    public void SetTextAlign(int align)
    {
        var record = new byte[8];
        var pos = 0;
        pos = SetUint32(record, pos, record.Length / 2);
        pos = SetUint16(record, pos, WmfConstants.RECORD_SET_TEXT_ALIGN);
        pos = SetInt16(record, pos, align);
        _records.Add(record);

        _dc.TextAlign = align;
    }

    public void SetTextCharacterExtra(int extra)
    {
        var record = new byte[8];
        var pos = 0;
        pos = SetUint32(record, pos, record.Length / 2);
        pos = SetUint16(record, pos, WmfConstants.RECORD_SET_TEXT_CHARACTER_EXTRA);
        pos = SetInt16(record, pos, extra);
        _records.Add(record);
    }

    public void SetTextColor(int color)
    {
        var record = new byte[10];
        var pos = 0;
        pos = SetUint32(record, pos, record.Length / 2);
        pos = SetUint16(record, pos, WmfConstants.RECORD_SET_TEXT_COLOR);
        pos = SetInt32(record, pos, color);
        _records.Add(record);
    }

    public void SetTextJustification(int breakExtra, int breakCount)
    {
        var record = new byte[10];
        var pos = 0;
        pos = SetUint32(record, pos, record.Length / 2);
        pos = SetUint16(record, pos, WmfConstants.RECORD_SET_TEXT_COLOR);
        pos = SetInt16(record, pos, breakCount);
        pos = SetInt16(record, pos, breakExtra);
        _records.Add(record);
    }

    public void SetViewportExtEx(int x, int y, Size? old)
    {
        var record = new byte[10];
        var pos = 0;
        pos = SetUint32(record, pos, record.Length / 2);
        pos = SetUint16(record, pos, WmfConstants.RECORD_SET_VIEWPORT_EXT_EX);
        pos = SetInt16(record, pos, y);
        pos = SetInt16(record, pos, x);
        _records.Add(record);
    }

    public void SetViewportOrgEx(int x, int y, Point? old)
    {
        var record = new byte[10];
        var pos = 0;
        pos = SetUint32(record, pos, record.Length / 2);
        pos = SetUint16(record, pos, WmfConstants.RECORD_SET_VIEWPORT_ORG_EX);
        pos = SetInt16(record, pos, y);
        pos = SetInt16(record, pos, x);
        _records.Add(record);
    }

    public void SetWindowExtEx(int width, int height, Size? old)
    {
        var record = new byte[10];
        var pos = 0;
        pos = SetUint32(record, pos, record.Length / 2);
        pos = SetUint16(record, pos, WmfConstants.RECORD_SET_WINDOW_EXT_EX);
        pos = SetInt16(record, pos, height);
        pos = SetInt16(record, pos, width);
        _records.Add(record);
    }

    public void SetWindowOrgEx(int x, int y, Point? old)
    {
        var record = new byte[10];
        var pos = 0;
        pos = SetUint32(record, pos, record.Length / 2);
        pos = SetUint16(record, pos, WmfConstants.RECORD_SET_WINDOW_ORG_EX);
        pos = SetInt16(record, pos, y);
        pos = SetInt16(record, pos, x);
        _records.Add(record);
    }

    public void StretchBlt(byte[] image, int dx, int dy, int dw, int dh,
        int sx, int sy, int sw, int sh, long rop)
    {
        var record = new byte[26 + (image.Length + image.Length % 2)];
        var pos = 0;
        pos = SetUint32(record, pos, record.Length / 2);
        pos = SetUint16(record, pos, WmfConstants.RECORD_STRETCH_BLT);
        pos = SetUint32(record, pos, rop);
        pos = SetInt16(record, pos, sh);
        pos = SetInt16(record, pos, sw);
        pos = SetInt16(record, pos, sy);
        pos = SetInt16(record, pos, sx);
        pos = SetInt16(record, pos, dw);
        pos = SetInt16(record, pos, dh);
        pos = SetInt16(record, pos, dy);
        pos = SetInt16(record, pos, dx);
        pos = SetBytes(record, pos, image);
        if (image.Length % 2 == 1)
        {
            pos = SetByte(record, pos, 0);
        }

        _records.Add(record);
    }

    public void StretchDIBits(int dx, int dy, int dw, int dh, int sx, int sy,
        int sw, int sh, byte[] image, int usage, long rop)
    {
        var record = new byte[26 + (image.Length + image.Length % 2)];
        var pos = 0;
        pos = SetUint32(record, pos, record.Length / 2);
        pos = SetUint16(record, pos, WmfConstants.RECORD_STRETCH_DIBITS);
        pos = SetUint32(record, pos, rop);
        pos = SetUint16(record, pos, usage);
        pos = SetInt16(record, pos, sh);
        pos = SetInt16(record, pos, sw);
        pos = SetInt16(record, pos, sy);
        pos = SetInt16(record, pos, sx);
        pos = SetInt16(record, pos, dw);
        pos = SetInt16(record, pos, dh);
        pos = SetInt16(record, pos, dy);
        pos = SetInt16(record, pos, dx);
        pos = SetBytes(record, pos, image);
        if (image.Length % 2 == 1)
        {
            pos = SetByte(record, pos, 0);
        }

        _records.Add(record);
    }

    public void TextOut(int x, int y, byte[] text)
    {
        var record = new byte[10 + text.Length + text.Length % 2];
        var pos = 0;
        pos = SetUint32(record, pos, record.Length / 2);
        pos = SetUint16(record, pos, WmfConstants.RECORD_TEXT_OUT);
        pos = SetInt16(record, pos, text.Length);
        pos = SetBytes(record, pos, text);
        if (text.Length % 2 == 1)
        {
            pos = SetByte(record, pos, 0);
        }

        pos = SetInt16(record, pos, y);
        pos = SetInt16(record, pos, x);
        _records.Add(record);
    }

    public void Footer()
    {
        var pos = 0;
        if (_header != null)
        {
            long size = _header.Length;
            long maxRecordSize = 0;
            foreach (var record in _records)
            {
                size += record.Length;
                if (record.Length > maxRecordSize)
                {
                    maxRecordSize = record.Length;
                }
            }

            pos = SetUint32(_header, 6, size / 2);
            pos = SetUint16(_header, pos, _objects.Count);
            pos = SetUint32(_header, pos, maxRecordSize / 2);
        }

        var footerRecord = new byte[6];
        pos = 0;
        pos = SetUint32(footerRecord, pos, footerRecord.Length / 2);
        pos = SetUint16(footerRecord, pos, 0x0000);
        _records.Add(footerRecord);
    }

    private static int SetByte(byte[] output, int pos, int value)
    {
        output[pos] = (byte)(0xFF & value);
        return pos + 1;
    }

    private static int SetBytes(byte[] output, int pos, byte[] data)
    {
        Array.Copy(data, 0, output, pos, data.Length);
        return pos + data.Length;
    }

    private static int SetInt16(byte[] output, int pos, int value)
    {
        output[pos] = (byte)(0xFF & value);
        output[pos + 1] = (byte)(0xFF & (value >> 8));
        return pos + 2;
    }

    private static int SetInt32(byte[] output, int pos, int value)
    {
        output[pos] = (byte)(0xFF & value);
        output[pos + 1] = (byte)(0xFF & (value >> 8));
        output[pos + 2] = (byte)(0xFF & (value >> 16));
        output[pos + 3] = (byte)(0xFF & (value >> 24));
        return pos + 4;
    }

    private static int SetUint16(byte[] output, int pos, int value)
    {
        output[pos] = (byte)(0xFF & value);
        output[pos + 1] = (byte)(0xFF & (value >> 8));
        return pos + 2;
    }

    private static int SetUint32(byte[] output, int pos, long value)
    {
        output[pos] = (byte)(0xFF & value);
        output[pos + 1] = (byte)(0xFF & (value >> 8));
        output[pos + 2] = (byte)(0xFF & (value >> 16));
        output[pos + 3] = (byte)(0xFF & (value >> 24));
        return pos + 4;
    }
}