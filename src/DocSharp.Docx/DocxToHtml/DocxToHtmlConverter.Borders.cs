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
    internal void ProcessBorder(BorderType? border, string? cssAttribute, ref List<string> styles)
    {
        if (border == null || cssAttribute == null)
        {
            return;
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

        string borderColor;
        if (border.Color?.Value != null)
            borderColor = ColorHelpers.EnsureHexColor(border.Color.Value) ?? "000000";
        else
            borderColor = "000000";

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

        if (border.Space != null && border.Space.Value > 0) // for paragraphs only
        {
            string? paddingAttribute = MapParagraphBorderToPadding(border);
            if (paddingAttribute != null)
            {
                // Border spacing is expressed in points (not twips) in DOCX.
                styles.Add($"{paddingAttribute}: {border.Space.Value.ToStringInvariant()}pt;");
            }
        }

        styles.Add($"{cssAttribute}: {borderWidth} {borderStyle} #{borderColor};");
    }

    internal string? MapParagraphBorderAttribute(BorderType border, bool isVertical = false)
    {
        // Only border types that can be found in ParagraphBorders are processed here.
        if (border is LeftBorder)
            return "border-left";
        else if (border is RightBorder)
            return "border-right";
        // TODO: top and bottom borders should create a box around paragraphs with the same borders;
        // currently they are treated like the "Between" border.
        else if (border is TopBorder)
            return "border-top";
        else if (border is BottomBorder)
            return "border-bottom";
        else if (border is BetweenBorder) // horizontal border between identical paragraphs
            return "border-bottom";
        else if (border is BarBorder) // paragraph border between facing pages
            return isVertical ? "border-right" : "border-inline-end";
        // If the paragraph has vertical orientation, inline-end is considered the bottom border (incorrect)
        else
            return null;
    }

    internal string? MapParagraphBorderToPadding(BorderType border, bool isVertical = false)
    {
        // Only border types that can be found in ParagraphBorders are processed here.
        if (border is LeftBorder)
            return "padding-left";
        else if (border is RightBorder)
            return "padding-right";
        // TODO: top and bottom borders should create a box around paragraphs with the same borders;
        // currently they are treated like the "Between" border.
        else if (border is TopBorder)
            return "padding-top";
        else if (border is BottomBorder)
            return "padding-bottom";
        else if (border is BetweenBorder) // horizontal border between identical paragraphs
            return "padding-bottom";
        else if (border is BarBorder) // paragraph border between facing pages
            return isVertical ? "padding-right" : "padding-inline-end";
        // If the paragraph has vertical orientation, inline-end is considered the bottom border (incorrect)
        else
            return null;
    }

    internal string? MapTableBorderAttribute(BorderType border)
    {
        // Only border types that can be found in TableBorders are processed here.
        if (border is LeftBorder)
            return "border-left";
        else if (border is RightBorder)
            return "border-right";
        else if (border is TopBorder)
            return "border-top";
        else if (border is BottomBorder)
            return "border-bottom";
        else if (border is StartBorder)
            return "border-inline-start";
        else if (border is EndBorder)
            return "border-inline-end";
        else
            return null;
        // Note: InsideHorizontalBorder and InsideVerticalBorder don't have a CSS equivalent,
        // so they are detected when processing table cells instead.
    }

    internal string? MapTableCellBorderAttribute(BorderType border, Primitives.BorderValue effectiveBorderType, bool isVertical, bool isFirstRow, bool isFirstColumn, bool isLastRow, bool isLastColumn)
    {
        // Only border types that are relevant to table cells are processed here.
        if (border is LeftBorder)
            return "border-left";
        else if (border is RightBorder)
            return "border-right";
        else if (border is TopBorder)
            return "border-top";
        else if (border is BottomBorder)
            return "border-bottom";
        else if (border is StartBorder)
            return isVertical ? "border-left" : "border-inline-start";
            // If the cell has vertical orientation, inline-start is considered the top border (incorrect)
        else if (border is EndBorder)
            return isVertical ? "border-right" : "border-inline-end";
            // If the cell has vertical orientation, inline-end is considered the bottom border (incorrect)
        else if (border is InsideHorizontalBorder)
        {
            if (effectiveBorderType == Primitives.BorderValue.Top)
                return "border-top";
            else if (effectiveBorderType == Primitives.BorderValue.Bottom)
                return "border-bottom";
            else
                return null;
        }
        else if (border is InsideVerticalBorder)
        {
            if (effectiveBorderType == Primitives.BorderValue.Start)
                return isVertical ? "border-left" : "border-inline-start";
            else if (effectiveBorderType == Primitives.BorderValue.End)
                return isVertical ? "border-right" : "border-inline-end";
            // If the cell has vertical orientation, inline-start and inline-end are considered the top/bottom borders (incorrect)
            else if (effectiveBorderType == Primitives.BorderValue.Left)
                return "border-left";
            else if (effectiveBorderType == Primitives.BorderValue.End)
                return "border-right";
            else
                return null;
        }
        else 
            return null;
    }
}
