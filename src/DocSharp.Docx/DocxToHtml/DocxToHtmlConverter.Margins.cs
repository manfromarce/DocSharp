using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DocSharp.Helpers;
using DocSharp.Writers;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;

namespace DocSharp.Docx;

public partial class DocxToHtmlConverter : DocxToXmlWriterBase<HtmlTextWriter>
{
    internal string? MapMarginAttribute(OpenXmlElement margin, bool isVertical)
    {
        if (margin is LeftMargin || margin is TableCellLeftMargin)
            return "padding-left";
        else if (margin is RightMargin || margin is TableCellRightMargin)
            return "padding-right";
        else if (margin is TopMargin)
            return "padding-top";
        else if (margin is BottomMargin)
            return "padding-bottom";
        else if (margin is StartMargin)
            return isVertical ? "padding-left" : "padding-inline-start";
            // If the cell has vertical orientation, inline-start is considered the top padding (incorrect)
        else if (margin is EndMargin)
            return isVertical ? "padding-right" : "padding-inline-end";
            // If the cell has vertical orientation, inline-end is considered the bottom padding (incorrect)
        else 
            return null;
    }
}
