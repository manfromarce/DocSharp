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

public partial class DocxToHtmlConverter : DocxToTextWriterBase<HtmlTextWriter>
{
    internal void ProcessBorder(BorderType? border, ref List<string> styles, bool isTableCell, bool isLastRow = false, bool isLastColumn = false, bool isVertical = false)
    {
        if (border == null)
        {
            return;
        }
        string cssAttribute = "border-left";
        if (border is RightBorder)
        {
            cssAttribute = "border-right";
        }
        // TODO: top and bottom borders should create a box around paragraphs with the same borders;
        // currently they are treated like the "Between" border.
        else if (border is TopBorder)
        {
            cssAttribute = "border-top";
        }
        else if (border is BottomBorder)
        {
            cssAttribute = "border-bottom";
        }
        else if (border is BetweenBorder) // horizontal border between identical paragraphs
        {
            cssAttribute = "border-bottom";
        }
        else if (border is BarBorder) // paragraph border between facing pages
        {
            cssAttribute = isVertical ? "border-right" : "border-inline-end";
            // If the cell has vertical orientation, inline-end is considered the bottom border (incorrect)
        }
        else if (border is StartBorder)
        {
            cssAttribute = isVertical ? "border-left" : "border-inline-start";
            // If the cell has vertical orientation, inline-start is considered the top border (incorrect)
        }
        else if (border is EndBorder)
        {
            cssAttribute = isVertical ? "border-right" : "border-inline-end";
            // If the cell has vertical orientation, inline-end is considered the bottom border (incorrect)
        }
        else if (border is InsideHorizontalBorder) // for tables
        {
            if (isLastRow)
                return;
            cssAttribute = "border-bottom";
        }
        else if (border is InsideVerticalBorder) // for tables
        {
            if (isLastColumn)
                return;
            cssAttribute = isVertical ? "border-right" : "border-inline-end"; 
            // If the cell has vertical orientation, inline-end is considered the bottom border (incorrect)
        }
        else if (border is Border) // Used for characters borders (same for top, left, bottom and right)
        {
            cssAttribute = "border";
        }

        string borderStyle = "solid";
        if (border.Val != null)
        {
            if (border.Val.Value == BorderValues.Dashed || border.Val.Value == BorderValues.DashSmallGap ||
                border.Val.Value == BorderValues.DotDash || border.Val.Value == BorderValues.DotDotDash ||
                border.Val.Value == BorderValues.DashDotStroked)
            {
                borderStyle = "dashed";
            }
            else if (border.Val.Value == BorderValues.Dotted)
            {
                borderStyle = "dotted";
            }
            else if (border.Val.Value == BorderValues.Double || border.Val.Value == BorderValues.DoubleWave)
            {
                borderStyle = "double";
            }
            else if (border.Val.Value == BorderValues.Triple)
            {
                borderStyle = "double"; // Triple is not supported in CSS
            }
            else if (border.Val.Value == BorderValues.Outset)
            {
                borderStyle = "outset";
            }
            else if (border.Val.Value == BorderValues.Inset)
            {
                borderStyle = "inset";
            }
            else if (border.Val.Value == BorderValues.ThreeDEmboss)
            {
                borderStyle = "ridge";
            }
            else if (border.Val.Value == BorderValues.ThreeDEngrave)
            {
                borderStyle = "groove";
            }
            else if (border.Val.Value == BorderValues.None || border.Val.Value == BorderValues.Nil)
            {
                borderStyle = "none";
            }
            else
            {
                borderStyle = "solid"; // Default value
            }
        }

        string borderWidth = "1px";
        if (border.Size != null)
        {
            // Open XML uses 1/8 points for border width
            double sizeInPoints = border.Size.Value / 8.0;
            borderWidth = $"{sizeInPoints.ToStringInvariant(2)}pt";
        }

        string borderColor = "#000000";
        if (border.Color?.Value != null && ColorHelpers.IsValidHexColor(border.Color.Value))
        {
            borderColor = border.Color.Value;
        }
        if (!borderColor.StartsWith('#'))
        {
            borderColor = "#" + borderColor;
        }

        // TODO: box-shadow
        //if (border.Shadow != null && ((!border.Shadow.HasValue) || border.Shadow.Value))
        //{
        //    if (border is RightBorder || border is EndBorder)
        //    {
        //    }
        //    if (border is BottomBorder)
        //    {
        //    }
        //}
        //if (border.Frame != null && ((!border.Frame.HasValue) || border.Frame.Value))
        //{
        //}

        //if (isTableCell && border.Space != null && border.Space.Value > 0)

        styles.Add($"{cssAttribute}: {borderWidth} {borderStyle} {borderColor};");
    }
}
