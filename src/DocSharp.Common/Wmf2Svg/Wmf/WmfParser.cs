using System;
using System.IO;
using DocSharp;
using DocSharp.Wmf2Svg.Gdi;
using DocSharp.Wmf2Svg.Svg;

namespace DocSharp.Wmf2Svg.Wmf;

public sealed class WmfParser
{
    /// <summary>
    /// Creates SVG object from WMF bytes
    /// </summary>
    /// <param name="data">the source byte array</param>
    /// <param name="compatible">output IE9 compatible style. but it's dirty and approximative</param>
    /// <param name="replaceSymbolFont">replace SYMBOL Font to serif or sans-serif Unicode SYMBOL</param>
    /// <returns></returns>
    public SvgGdi Parse(byte[] data, bool compatible = false, bool replaceSymbolFont = false, IImageConverter? imageConverter = null)
    {
        using var stream = new MemoryStream(data, writable: false);

        return Parse(stream, compatible, replaceSymbolFont, imageConverter);
    }

    /// <summary>
    /// Creates SVG object from WMF bytes
    /// </summary>
    /// <param name="data">the source byte array</param>
    /// <param name="index">the index in byte array at which the stream begins</param>
    /// <param name="count">the length of the stream in bytes</param>
    /// <param name="compatible">output IE9 compatible style. but it's dirty and approximative</param>
    /// <param name="replaceSymbolFont">replace SYMBOL Font to serif or sans-serif Unicode SYMBOL</param>
    /// <returns></returns>
    public SvgGdi Parse(byte[] data, int index, int count, bool compatible = false, bool replaceSymbolFont = false, IImageConverter? imageConverter = null)
    {
        var stream = new MemoryStream(data, index: index, count: count, writable: false);

        return Parse(stream, compatible, replaceSymbolFont, imageConverter);
    }

    /// <summary>
    /// Creates SVG object from WMF stream
    /// </summary>
    /// <param name="stream">the source stream</param>
    /// <param name="compatible">output IE9 compatible style. but it's dirty and approximative</param>
    /// /// <param name="replaceSymbolFont">replace SYMBOL Font to serif or sans-serif Unicode SYMBOL</param>
    /// <returns></returns>
    /// <exception cref="WmfParseException"></exception>
    public SvgGdi Parse(Stream stream, bool compatible = false, bool replaceSymbolFont = false, IImageConverter? imageConverter = null)
    {
        var gdi = new SvgGdi(compatible, imageConverter);

        gdi.ReplaceSymbolFont = replaceSymbolFont;

        using var input = new DataInput(new BufferedStream(stream), isLittleEndian: true);

        var isEmpty = true;

        try
        {
            int mtType;
            int mtHeaderSize;

            var key = input.ReadUint32();
            isEmpty = false;

            if (key == 0x9AC6CDD7)
            {
                input.ReadInt16(); // hmf
                var vsx = input.ReadInt16();
                var vsy = input.ReadInt16();
                var vex = input.ReadInt16();
                var vey = input.ReadInt16();
                var dpi = input.ReadUint16();
                input.ReadUint32(); // reserved
                input.ReadUint16(); // checksum

                gdi.PlaceableHeader(vsx, vsy, vex, vey, dpi);

                mtType = input.ReadUint16();
                mtHeaderSize = input.ReadUint16();
            }
            else
            {
                mtType = (int)(key & 0x0000FFFF);
                mtHeaderSize = (int)((key & 0xFFFF0000) >> 16);
            }

            input.ReadUint16(); // mtVersion
            input.ReadUint32(); // mtSize
            var mtNoObjects = input.ReadUint16();
            input.ReadUint32(); // mtMaxRecord
            input.ReadUint16(); // mtNoParameters

            if (mtType != 1 || mtHeaderSize != 9)
            {
                throw new WmfParseException("Invalid format");
            }

            gdi.Header();

            var objs = new IGdiObject?[mtNoObjects];

            while (true)
            {
                var size = (int)input.ReadUint32() - 3;
                var id = input.ReadUint16();

                if (id == WmfConstants.RECORD_EOF)
                {
                    break;
                }

                input.Count = 0;

                switch (id)
                {
                    case WmfConstants.RECORD_REALIZE_PALETTE:
                        gdi.RealizePalette();
                        break;

                    case WmfConstants.RECORD_SET_PALETTE_ENTRIES:
                    {
                        var entries = new int[input.ReadUint16()];
                        var startIndex = input.ReadUint16();
                        var objId = input.ReadUint16();
                        for (var i = 0; i < entries.Length; i++)
                        {
                            entries[i] = input.ReadInt32();
                        }

                        gdi.SetPaletteEntries((IGdiPalette)objs[objId]!, startIndex, entries);
                        break;
                    }

                    case WmfConstants.RECORD_SET_BK_MODE:
                    {
                        var mode = input.ReadInt16();
                        gdi.SetBkMode(mode);
                        break;
                    }

                    case WmfConstants.RECORD_SET_MAP_MODE:
                    {
                        var mode = input.ReadInt16();
                        gdi.SetMapMode(mode);
                        break;
                    }

                    case WmfConstants.RECORD_SET_ROP2:
                    {
                        var mode = input.ReadInt16();
                        gdi.SetROP2(mode);
                        break;
                    }

                    case WmfConstants.RECORD_SET_REL_ABS:
                    {
                        var mode = input.ReadInt16();
                        gdi.SetRelAbs(mode);
                        break;
                    }

                    case WmfConstants.RECORD_SET_POLY_FILL_MODE:
                    {
                        var mode = input.ReadInt16();
                        gdi.SetPolyFillMode(mode);
                        break;
                    }

                    case WmfConstants.RECORD_SET_STRETCH_BLT_MODE:
                    {
                        var mode = input.ReadInt16();
                        gdi.SetStretchBltMode(mode);
                        break;
                    }

                    case WmfConstants.RECORD_SET_TEXT_CHARACTER_EXTRA:
                    {
                        var extra = input.ReadInt16();
                        gdi.SetTextCharacterExtra(extra);
                        break;
                    }

                    case WmfConstants.RECORD_RESTORE_DC:
                    {
                        var dc = input.ReadInt16();
                        gdi.RestoreDC(dc);
                        break;
                    }

                    case WmfConstants.RECORD_RESIZE_PALETTE:
                    {
                        var objId = input.ReadUint16();
                        gdi.ResizePalette((IGdiPalette)objs[objId]!);
                        break;
                    }

                    case WmfConstants.RECORD_DIB_CREATE_PATTERN_BRUSH:
                    {
                        var usage = input.ReadInt32();
                        var image = input.ReadBytes(size * 2 - input.Count);

                        for (var i = 0; i < objs.Length; i++)
                        {
                            if (objs[i] == null)
                            {
                                objs[i] = gdi.DibCreatePatternBrush(image, usage);
                                break;
                            }
                        }

                        break;
                    }

                    case WmfConstants.RECORD_SET_LAYOUT:
                    {
                        var layout = input.ReadUint32();
                        gdi.SetLayout(layout);
                        break;
                    }

                    case WmfConstants.RECORD_SET_BK_COLOR:
                    {
                        var color = input.ReadInt32();
                        gdi.SetBkColor(color);
                        break;
                    }

                    case WmfConstants.RECORD_SET_TEXT_COLOR:
                    {
                        var color = input.ReadInt32();
                        gdi.SetTextColor(color);
                        break;
                    }

                    case WmfConstants.RECORD_OFFSET_VIEWPORT_ORG_EX:
                    {
                        var y = input.ReadInt16();
                        var x = input.ReadInt16();
                        gdi.OffsetViewportOrgEx(x, y, null);
                        break;
                    }

                    case WmfConstants.RECORD_LINE_TO:
                    {
                        var ey = input.ReadInt16();
                        var ex = input.ReadInt16();
                        gdi.LineTo(ex, ey);
                        break;
                    }

                    case WmfConstants.RECORD_MOVE_TO_EX:
                    {
                        var y = input.ReadInt16();
                        var x = input.ReadInt16();
                        gdi.MoveToEx(x, y, null);
                        break;
                    }

                    case WmfConstants.RECORD_OFFSET_CLIP_RGN:
                    {
                        var y = input.ReadInt16();
                        var x = input.ReadInt16();
                        gdi.OffsetClipRgn(x, y);
                        break;
                    }

                    case WmfConstants.RECORD_FILL_RGN:
                    {
                        var brushId = input.ReadUint16();
                        var rgnId = input.ReadUint16();
                        gdi.FillRgn((IGdiRegion)objs[rgnId]!, (IGdiBrush)objs[brushId]!);
                        break;
                    }

                    case WmfConstants.RECORD_SET_MAPPER_FLAGS:
                    {
                        var flag = input.ReadUint32();
                        gdi.SetMapperFlags(flag);
                        break;
                    }

                    case WmfConstants.RECORD_SELECT_PALETTE:
                    {
                        var mode = input.ReadInt16() != 0;
                        if (size * 2 - input.Count > 0)
                        {
                            var objId = input.ReadUint16();
                            gdi.SelectPalette((IGdiPalette)objs[objId]!, mode);
                        }

                        break;
                    }

                    case WmfConstants.RECORD_POLYGON:
                    {
                        var points = new Point[input.ReadInt16()];
                        for (var i = 0; i < points.Length; i++)
                        {
                            points[i] = new Point(input.ReadInt16(), input.ReadInt16());
                        }

                        gdi.Polygon(points);
                        break;
                    }

                    case WmfConstants.RECORD_POLYLINE:
                    {
                        var points = new Point[input.ReadInt16()];
                        for (var i = 0; i < points.Length; i++)
                        {
                            points[i] = new Point(input.ReadInt16(), input.ReadInt16());
                        }

                        gdi.Polyline(points);
                        break;
                    }

                    case WmfConstants.RECORD_SET_TEXT_JUSTIFICATION:
                    {
                        var breakCount = input.ReadInt16();
                        var breakExtra = input.ReadInt16();
                        gdi.SetTextJustification(breakExtra, breakCount);
                        break;
                    }

                    case WmfConstants.RECORD_SET_WINDOW_ORG_EX:
                    {
                        var y = input.ReadInt16();
                        var x = input.ReadInt16();
                        gdi.SetWindowOrgEx(x, y, null);
                        break;
                    }

                    case WmfConstants.RECORD_SET_WINDOW_EXT_EX:
                    {
                        var height = input.ReadInt16();
                        var width = input.ReadInt16();
                        gdi.SetWindowExtEx(width, height, null);
                        break;
                    }

                    case WmfConstants.RECORD_SET_VIEWPORT_ORG_EX:
                    {
                        var y = input.ReadInt16();
                        var x = input.ReadInt16();
                        gdi.SetViewportOrgEx(x, y, null);
                        break;
                    }

                    case WmfConstants.RECORD_SET_VIEWPORT_EXT_EX:
                    {
                        var y = input.ReadInt16();
                        var x = input.ReadInt16();
                        gdi.SetViewportExtEx(x, y, null);
                        break;
                    }

                    case WmfConstants.RECORD_OFFSET_WINDOW_ORG_EX:
                    {
                        var y = input.ReadInt16();
                        var x = input.ReadInt16();
                        gdi.OffsetWindowOrgEx(x, y, null);
                        break;
                    }

                    case WmfConstants.RECORD_SCALE_WINDOW_EXT_EX:
                    {
                        var yd = input.ReadInt16();
                        var y = input.ReadInt16();
                        var xd = input.ReadInt16();
                        var x = input.ReadInt16();
                        gdi.ScaleWindowExtEx(x, xd, y, yd, null);
                        break;
                    }

                    case WmfConstants.RECORD_SCALE_VIEWPORT_EXT_EX:
                    {
                        var yd = input.ReadInt16();
                        var y = input.ReadInt16();
                        var xd = input.ReadInt16();
                        var x = input.ReadInt16();
                        gdi.ScaleViewportExtEx(x, xd, y, yd, null);
                        break;
                    }

                    case WmfConstants.RECORD_EXCLUDE_CLIP_RECT:
                    {
                        var ey = input.ReadInt16();
                        var ex = input.ReadInt16();
                        var sy = input.ReadInt16();
                        var sx = input.ReadInt16();
                        gdi.ExcludeClipRect(sx, sy, ex, ey);
                        break;
                    }

                    case WmfConstants.RECORD_INTERSECT_CLIP_RECT:
                    {
                        var ey = input.ReadInt16();
                        var ex = input.ReadInt16();
                        var sy = input.ReadInt16();
                        var sx = input.ReadInt16();
                        gdi.IntersectClipRect(sx, sy, ex, ey);
                        break;
                    }

                    case WmfConstants.RECORD_ELLIPSE:
                    {
                        var ey = input.ReadInt16();
                        var ex = input.ReadInt16();
                        var sy = input.ReadInt16();
                        var sx = input.ReadInt16();
                        gdi.Ellipse(sx, sy, ex, ey);
                        break;
                    }

                    case WmfConstants.RECORD_FLOOD_FILL:
                    {
                        var color = input.ReadInt32();
                        var y = input.ReadInt16();
                        var x = input.ReadInt16();
                        gdi.FloodFill(x, y, color);
                        break;
                    }

                    case WmfConstants.RECORD_FRAME_RGN:
                    {
                        var height = input.ReadInt16();
                        var width = input.ReadInt16();
                        var brushId = input.ReadUint16();
                        var rgnId = input.ReadUint16();
                        gdi.FrameRgn((IGdiRegion)objs[rgnId]!, (IGdiBrush)objs[brushId]!, width, height);
                        break;
                    }

                    case WmfConstants.RECORD_ANIMATE_PALETTE:
                    {
                        var entries = new int[input.ReadUint16()];
                        var startIndex = input.ReadUint16();
                        var objId = input.ReadUint16();
                        for (var i = 0; i < entries.Length; i++)
                        {
                            entries[i] = input.ReadInt32();
                        }

                        gdi.AnimatePalette((IGdiPalette)objs[objId]!, startIndex, entries);
                        break;
                    }

                    case WmfConstants.RECORD_TEXT_OUT:
                    {
                        var count = input.ReadInt16();
                        var text = input.ReadBytes(count);
                        if (count % 2 == 1)
                        {
                            input.ReadByte();
                        }

                        var y = input.ReadInt16();
                        var x = input.ReadInt16();
                        gdi.TextOut(x, y, text);
                        break;
                    }

                    case WmfConstants.RECORD_POLY_POLYGON:
                    {
                        var points = new Point[input.ReadInt16()][];
                        for (var i = 0; i < points.Length; i++)
                        {
                            points[i] = new Point[input.ReadInt16()];
                        }

                        for (var i = 0; i < points.Length; i++)
                        {
                            for (var j = 0; j < points[i].Length; j++)
                            {
                                points[i][j] = new Point(input.ReadInt16(), input.ReadInt16());
                            }
                        }

                        gdi.PolyPolygon(points);
                        break;
                    }

                    case WmfConstants.RECORD_EXT_FLOOD_FILL:
                    {
                        var type = input.ReadUint16();
                        var color = input.ReadInt32();
                        var y = input.ReadInt16();
                        var x = input.ReadInt16();
                        gdi.ExtFloodFill(x, y, color, type);
                        break;
                    }

                    case WmfConstants.RECORD_RECTANGLE:
                    {
                        var ey = input.ReadInt16();
                        var ex = input.ReadInt16();
                        var sy = input.ReadInt16();
                        var sx = input.ReadInt16();
                        gdi.Rectangle(sx, sy, ex, ey);
                        break;
                    }

                    case WmfConstants.RECORD_SET_PIXEL:
                    {
                        var color = input.ReadInt32();
                        var y = input.ReadInt16();
                        var x = input.ReadInt16();
                        gdi.SetPixel(x, y, color);
                        break;
                    }

                    case WmfConstants.RECORD_ROUND_RECT:
                    {
                        var rh = input.ReadInt16();
                        var rw = input.ReadInt16();
                        var ey = input.ReadInt16();
                        var ex = input.ReadInt16();
                        var sy = input.ReadInt16();
                        var sx = input.ReadInt16();
                        gdi.RoundRect(sx, sy, ex, ey, rw, rh);
                        break;
                    }

                    case WmfConstants.RECORD_PAT_BLT:
                    {
                        var rop = input.ReadUint32();
                        var height = input.ReadInt16();
                        var width = input.ReadInt16();
                        var y = input.ReadInt16();
                        var x = input.ReadInt16();
                        gdi.PatBlt(x, y, width, height, rop);
                        break;
                    }

                    case WmfConstants.RECORD_SAVE_DC:
                        gdi.SaveDC();
                        break;

                    case WmfConstants.RECORD_PIE:
                    {
                        var eyr = input.ReadInt16();
                        var exr = input.ReadInt16();
                        var syr = input.ReadInt16();
                        var sxr = input.ReadInt16();
                        var ey = input.ReadInt16();
                        var ex = input.ReadInt16();
                        var sy = input.ReadInt16();
                        var sx = input.ReadInt16();
                        gdi.Pie(sx, sy, ex, ey, sxr, syr, exr, eyr);
                        break;
                    }

                    case WmfConstants.RECORD_STRETCH_BLT:
                    {
                        var rop = input.ReadUint32();
                        var sh = input.ReadInt16();
                        var sw = input.ReadInt16();
                        var sy = input.ReadInt16();
                        var sx = input.ReadInt16();
                        var dh = input.ReadInt16();
                        var dw = input.ReadInt16();
                        var dy = input.ReadInt16();
                        var dx = input.ReadInt16();

                        var image = input.ReadBytes(size * 2 - input.Count);

                        gdi.StretchBlt(image, dx, dy, dw, dh, sx, sy, sw, sh, rop);
                        break;
                    }

                    case WmfConstants.RECORD_ESCAPE:
                    {
                        var data = input.ReadBytes(2 * size);
                        gdi.Escape(data);
                        break;
                    }

                    case WmfConstants.RECORD_INVERT_RGN:
                    {
                        var rgnId = input.ReadUint16();
                        gdi.InvertRgn((IGdiRegion)objs[rgnId]!);
                        break;
                    }

                    case WmfConstants.RECORD_PAINT_RGN:
                    {
                        var objId = input.ReadUint16();
                        gdi.PaintRgn((IGdiRegion)objs[objId]!);
                        break;
                    }

                    case WmfConstants.RECORD_SELECT_CLIP_RGN:
                    {
                        var objId = input.ReadUint16();
                        var rgn = objId > 0 ? (IGdiRegion)objs[objId]! : null;
                        gdi.SelectClipRgn(rgn);
                        break;
                    }

                    case WmfConstants.RECORD_SELECT_OBJECT:
                    {
                        var objId = input.ReadUint16();
                        gdi.SelectObject(objs[objId]!);
                        break;
                    }

                    case WmfConstants.RECORD_SET_TEXT_ALIGN:
                    {
                        var align = input.ReadInt16();
                        gdi.SetTextAlign(align);
                        break;
                    }

                    case WmfConstants.RECORD_ARC:
                    {
                        var eya = input.ReadInt16();
                        var exa = input.ReadInt16();
                        var sya = input.ReadInt16();
                        var sxa = input.ReadInt16();
                        var eyr = input.ReadInt16();
                        var exr = input.ReadInt16();
                        var syr = input.ReadInt16();
                        var sxr = input.ReadInt16();
                        gdi.Arc(sxr, syr, exr, eyr, sxa, sya, exa, eya);
                        break;
                    }

                    case WmfConstants.RECORD_CHORD:
                    {
                        var eya = input.ReadInt16();
                        var exa = input.ReadInt16();
                        var sya = input.ReadInt16();
                        var sxa = input.ReadInt16();
                        var eyr = input.ReadInt16();
                        var exr = input.ReadInt16();
                        var syr = input.ReadInt16();
                        var sxr = input.ReadInt16();
                        gdi.Chord(sxr, syr, exr, eyr, sxa, sya, exa, eya);
                        break;
                    }

                    case WmfConstants.RECORD_BIT_BLT:
                    {
                        var rop = input.ReadUint32();
                        var sy = input.ReadInt16();
                        var sx = input.ReadInt16();
                        var height = input.ReadInt16();
                        var width = input.ReadInt16();
                        var dy = input.ReadInt16();
                        var dx = input.ReadInt16();

                        var image = input.ReadBytes(size * 2 - input.Count);

                        gdi.BitBlt(image, dx, dy, width, height, sx, sy, rop);
                        break;
                    }

                    case WmfConstants.RECORD_EXT_TEXT_OUT:
                    {
                        var rsize = size;

                        var y = input.ReadInt16();
                        var x = input.ReadInt16();
                        var count = input.ReadInt16();
                        var options = input.ReadUint16();
                        rsize -= 4;

                        int[]? rect = null;
                        if ((options & 0x0006) > 0)
                        {
                            rect = [input.ReadInt16(), input.ReadInt16(), input.ReadInt16(), input.ReadInt16()];
                            rsize -= 4;
                        }

                        var text = input.ReadBytes(count);
                        if (count % 2 == 1)
                        {
                            input.ReadByte();
                        }

                        rsize -= (count + 1) / 2;

                        int[]? dx = null;
                        if (rsize > 0)
                        {
                            dx = new int[rsize];
                            for (var i = 0; i < dx.Length; i++)
                            {
                                dx[i] = input.ReadInt16();
                            }
                        }

                        gdi.ExtTextOut(x, y, options, rect, text, dx);
                        break;
                    }

                    case WmfConstants.RECORD_SET_DIBITS_TO_DEVICE:
                    {
                        var colorUse = input.ReadUint16();
                        var scanlines = input.ReadUint16();
                        var startscan = input.ReadUint16();
                        var sy = input.ReadInt16();
                        var sx = input.ReadInt16();
                        var dh = input.ReadInt16();
                        var dw = input.ReadInt16();
                        var dy = input.ReadInt16();
                        var dx = input.ReadInt16();

                        var image = input.ReadBytes(size * 2 - input.Count);

                        gdi.SetDIBitsToDevice(dx, dy, dw, dh, sx, sy, startscan, scanlines, image, colorUse);
                        break;
                    }

                    case WmfConstants.RECORD_DIB_BIT_BLT:
                    {
                        var isRop = false;

                        var rop = input.ReadUint32();
                        var sy = input.ReadInt16();
                        var sx = input.ReadInt16();
                        var height = input.ReadInt16();
                        if (height == 0)
                        {
                            height = input.ReadInt16();
                            isRop = true;
                        }

                        var width = input.ReadInt16();
                        var dy = input.ReadInt16();
                        var dx = input.ReadInt16();

                        if (isRop)
                        {
                            gdi.DibBitBlt(Array.Empty<byte>(), dx, dy, width, height, sx, sy, rop);
                        }
                        else
                        {
                            var image = input.ReadBytes(size * 2 - input.Count);
                            gdi.DibBitBlt(image, dx, dy, width, height, sx, sy, rop);
                        }

                        break;
                    }

                    case WmfConstants.RECORD_DIB_STRETCH_BLT:
                    {
                        var rop = input.ReadUint32();
                        var sh = input.ReadInt16();
                        var sw = input.ReadInt16();
                        var sx = input.ReadInt16();
                        var sy = input.ReadInt16();
                        var dh = input.ReadInt16();
                        var dw = input.ReadInt16();
                        var dy = input.ReadInt16();
                        var dx = input.ReadInt16();

                        var image = input.ReadBytes(size * 2 - input.Count);

                        gdi.DibStretchBlt(image, dx, dy, dw, dh, sx, sy, sw, sh, rop);
                        break;
                    }

                    case WmfConstants.RECORD_STRETCH_DIBITS:
                    {
                        var rop = input.ReadUint32();
                        var usage = input.ReadUint16();
                        var sh = input.ReadInt16();
                        var sw = input.ReadInt16();
                        var sy = input.ReadInt16();
                        var sx = input.ReadInt16();
                        var dh = input.ReadInt16();
                        var dw = input.ReadInt16();
                        var dy = input.ReadInt16();
                        var dx = input.ReadInt16();

                        var image = input.ReadBytes(size * 2 - input.Count);

                        gdi.StretchDIBits(dx, dy, dw, dh, sx, sy, sw, sh, image, usage, rop);
                        break;
                    }

                    case WmfConstants.RECORD_DELETE_OBJECT:
                    {
                        var objId = input.ReadUint16();
                        gdi.DeleteObject(objs[objId]!);
                        objs[objId] = null;
                        break;
                    }

                    case WmfConstants.RECORD_CREATE_PALETTE:
                    {
                        var version = input.ReadUint16();
                        var entries = new int[input.ReadUint16()];
                        for (var i = 0; i < entries.Length; i++)
                        {
                            entries[i] = input.ReadInt32();
                        }

                        for (var i = 0; i < objs.Length; i++)
                        {
                            if (objs[i] == null)
                            {
                                objs[i] = gdi.CreatePalette(version, entries);
                                break;
                            }
                        }

                        break;
                    }

                    case WmfConstants.RECORD_CREATE_PATTERN_BRUSH:
                    {
                        var image = input.ReadBytes(size * 2 - input.Count);

                        for (var i = 0; i < objs.Length; i++)
                        {
                            if (objs[i] == null)
                            {
                                objs[i] = gdi.CreatePatternBrush(image);
                                break;
                            }
                        }

                        break;
                    }

                    case WmfConstants.RECORD_CREATE_PEN_INDIRECT:
                    {
                        var style = input.ReadUint16();
                        var width = input.ReadInt16();
                        input.ReadInt16();
                        var color = input.ReadInt32();
                        for (var i = 0; i < objs.Length; i++)
                        {
                            if (objs[i] == null)
                            {
                                objs[i] = gdi.CreatePenIndirect(style, width, color);
                                break;
                            }
                        }

                        break;
                    }

                    case WmfConstants.RECORD_CREATE_FONT_INDIRECT:
                    {
                        var height = input.ReadInt16();
                        var width = input.ReadInt16();
                        var escapement = input.ReadInt16();
                        var orientation = input.ReadInt16();
                        var weight = input.ReadInt16();
                        var italic = input.ReadByte() == 1;
                        var underline = input.ReadByte() == 1;
                        var strikeout = input.ReadByte() == 1;
                        var charset = input.ReadByte();
                        var outPrecision = input.ReadByte();
                        var clipPrecision = input.ReadByte();
                        var quality = input.ReadByte();
                        var pitchAndFamily = input.ReadByte();
                        var faceName = input.ReadBytes(size * 2 - input.Count);

                        IGdiObject obj = gdi.CreateFontIndirect(height, width, escapement, orientation, weight, italic,
                            underline, strikeout, charset, outPrecision, clipPrecision, quality, pitchAndFamily,
                            faceName);

                        for (var i = 0; i < objs.Length; i++)
                        {
                            if (objs[i] == null)
                            {
                                objs[i] = obj;
                                break;
                            }
                        }

                        break;
                    }

                    case WmfConstants.RECORD_CREATE_BRUSH_INDIRECT:
                    {
                        var style = input.ReadUint16();
                        var color = input.ReadInt32();
                        var hatch = input.ReadUint16();
                        for (var i = 0; i < objs.Length; i++)
                        {
                            if (objs[i] == null)
                            {
                                objs[i] = gdi.CreateBrushIndirect(style, color, hatch);
                                break;
                            }
                        }

                        break;
                    }

                    case WmfConstants.RECORD_CREATE_RECT_RGN:
                    {
                        var ey = input.ReadInt16();
                        var ex = input.ReadInt16();
                        var sy = input.ReadInt16();
                        var sx = input.ReadInt16();
                        for (var i = 0; i < objs.Length; i++)
                        {
                            if (objs[i] == null)
                            {
                                objs[i] = gdi.CreateRectRgn(sx, sy, ex, ey);
                                break;
                            }
                        }

                        break;
                    }

                    default:
                        throw new WmfParseException($"Unsupported id: {id} (size={size})");
                }

                var rest = size * 2 - input.Count;
                for (var i = 0; i < rest; i++)
                {
                    input.ReadByte();
                }
            }

            gdi.Footer();
        }
        catch (EndOfStreamException e)
        {
            if (isEmpty)
            {
                throw new WmfParseException("Failed to parse WMF, the source is empty", e);
            }

            throw new WmfParseException("Failed to parse WMF", e);
        }

        return gdi;
    }
}