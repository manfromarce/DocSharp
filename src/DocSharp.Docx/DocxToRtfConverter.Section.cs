using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using DocumentFormat.OpenXml.Wordprocessing;

namespace DocSharp.Docx;

public partial class DocxToRtfConverter
{
    internal override void ProcessSectionProperties(SectionProperties sectionProperties, StringBuilder sb)
    {
        if (sectionProperties.GetFirstChild<PageSize>() is PageSize size)
        {
            if (size.Width != null)
            {
                sb.Append($"\\paperw{size.Width.Value}");
            }
            if (size.Height != null)
            {
                sb.Append($"\\paperh{size.Height.Value}");
            }
            if (size.Orient != null && size.Orient.Value == PageOrientationValues.Landscape)
            {
                sb.Append($"\\landscape");
            }
            if (size.Code != null)
            {
                sb.Append($"\\psz{size.Code.Value}");
            }
        }
        if (sectionProperties.GetFirstChild<PageMargin>() is PageMargin margins)
        {
            if (margins.Top != null)
            {
                sb.Append($"\\margt{margins.Top.Value}");
            }
            if (margins.Bottom != null)
            {
                sb.Append($"\\margb{margins.Bottom.Value}");
            }
            if (margins.Left != null)
            {
                sb.Append($"\\margl{margins.Left.Value}");
            }
            if (margins.Right != null)
            {
                sb.Append($"\\margr{margins.Right.Value}");
            }
            if (margins.Gutter != null)
            {
                sb.Append($"\\gutter{margins.Gutter.Value}");
            }
        }
        if (sectionProperties.GetFirstChild<PageBorders>() is PageBorders borders)
        {
            int pageBorderOptions = 0;
            if (borders?.Display != null)
            {
                //PageBorderDisplayValues.AllPages --> 0
                if (borders.Display.Value == PageBorderDisplayValues.FirstPage)
                {
                    pageBorderOptions |= 1;
                }
                else if (borders.Display.Value == PageBorderDisplayValues.NotFirstPage)
                {
                    pageBorderOptions |= 2;
                }
            }
            if (borders?.ZOrder != null && borders.ZOrder == PageBorderZOrderValues.Back)
            {
                pageBorderOptions |= 1 << 3;
            }
            else
            {
                pageBorderOptions |= 0 << 3; // Front (default)
            }
            if (borders?.OffsetFrom != null && borders.OffsetFrom.Value == PageBorderOffsetValues.Page)
            {
                pageBorderOptions |= 1 << 5;
            }
            else
            {
                pageBorderOptions |= 0 << 5; // Offset from text
            }
            sb.Append(@"\pgbrdropt" + pageBorderOptions);
            if (borders?.TopBorder != null)
            {
                sb.Append(@"\pgbrdrt");
                ProcessBorder(borders.TopBorder, sb);
            }
            if (borders?.LeftBorder != null)
            {
                sb.Append(@"\pgbrdrl");
                ProcessBorder(borders.LeftBorder, sb);
            }
            if (borders?.BottomBorder != null)
            {
                sb.Append(@"\pgbrdrb");
                ProcessBorder(borders.BottomBorder, sb);
            }
            if (borders?.RightBorder != null)
            {
                sb.Append(@"\pgbrdrr");
                ProcessBorder(borders.RightBorder, sb);
            }
        }
        if (sectionProperties.GetFirstChild<Columns>() is Columns cols)
        {
            if (cols.ColumnCount != null)
            {
                sb.Append($"\\cols{cols.ColumnCount.Value}");
            }
            if (cols.Space != null)
            {
                sb.Append($"\\colsx{cols.Space.Value}");
            }
            if (cols.Separator != null && cols.Separator.HasValue && cols.Separator.Value)
            {
                sb.Append($"\\linebetcol");
            }
        }
        sb.AppendLine(" ");
    }
}
