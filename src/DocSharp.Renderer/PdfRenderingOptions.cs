using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using PeachPDF.PdfSharpCore.Drawing;

namespace DocSharp.Renderer;

public class PdfRenderingOptions
{
    public static readonly PdfRenderingOptions Default = new PdfRenderingOptions();

    public PdfRenderingOptions(
        bool hiddenChars = false,
        XPen? sectionBorders = null,
        XPen? headerBorders = null,
        XPen? footerBorders = null,
        XPen? paragraphBorders = null,
        XPen? lineBorders = null,
        XPen? wordBorders = null)
    {
        this.HiddenChars = hiddenChars;
        this.SectionBorders = sectionBorders;
        this.HeaderBorders = headerBorders;
        this.FooterBorders = footerBorders;
        this.ParagraphBorders = paragraphBorders;
        this.LineBorders = lineBorders;
        this.WordBorders = wordBorders;
    }

    /// <summary>
    /// e.g. Paragraph, PageBreak, SectionBreak
    /// </summary>
    public bool HiddenChars { get; }

    public XPen? SectionBorders { get; }
    public XPen? HeaderBorders { get; }
    public XPen? FooterBorders { get; }
    public XPen? ParagraphBorders { get; }
    public XPen? LineBorders { get; }
    public XPen? WordBorders { get; }

    public static PdfRenderingOptions WithDefaults(
        bool hiddenChars = true,
        bool section = false,
        bool header = false,
        bool footer = false,
        bool paragraph = false,
        bool line = false,
        bool word = false)
    {
        return new PdfRenderingOptions(
            hiddenChars,
            sectionBorders: section ? SectionDefault : null,
            headerBorders: header ? HeaderDefault : null,
            footerBorders: footer ? FooterDefault : null,
            paragraphBorders: paragraph ? ParagraphDefault : null,
            lineBorders: line ? LineDefault : null,
            wordBorders: word ? WordDefault : null
            );
    }

    public static readonly XPen ParagraphDefault = new XPen(XColors.Orange, 0.5f);
    public static readonly XPen LineDefault = new XPen(XColors.Red, 0.5f);
    public static readonly XPen WordDefault = new XPen(XColors.Green, 0.5f);
    public static readonly XPen HeaderDefault = new XPen(XColors.LightBlue, 0.5f);
    public static readonly XPen FooterDefault = new XPen(XColors.DarkBlue, 0.5f);
    public static readonly XPen SectionDefault = new XPen(XColors.Olive, 0.5f);
}
