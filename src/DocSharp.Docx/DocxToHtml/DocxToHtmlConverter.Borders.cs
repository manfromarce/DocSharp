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
    internal void ProcessBorder(BorderType? border, ref List<string> styles, bool isTableCell, bool isLastRow = false, bool isLastColumn = false)
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
        else if (border is TopBorder)
        {
            cssAttribute = "border-top";
        }
        else if (border is BottomBorder)
        {
            cssAttribute = "border-bottom";
        }
        else if (border is BarBorder)
        {
            cssAttribute = "border-right";
        }
        else if (border is BetweenBorder)
        {
            cssAttribute = "border-bottom";
        }
        else if (border is StartBorder)
        {
            cssAttribute = "border-inline-start";
        }
        else if (border is EndBorder)
        {
            cssAttribute = "border-inline-end";
        }
        else if (border is InsideHorizontalBorder)
        {
            if (isLastRow)
                return;
            cssAttribute = "border-bottom";
        }
        else if (border is InsideVerticalBorder)
        {
            if (isLastColumn)
                return;
            cssAttribute = "border-inline-end";
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
            borderWidth = $"{sizeInPoints.ToStringInvariant()}pt";
        }

        string borderColor = "000000";
        if (border.Color?.Value != null && !string.IsNullOrWhiteSpace(border.Color?.Value) && border.Color!.Value.Length == 6)
        {
            borderColor = border.Color.Value;
        }
        if (border.Shadow != null && ((!border.Shadow.HasValue) || border.Shadow.Value))
        {
            //borderStyle = "inset";
            borderStyle = "outset";
        }
        if (border.Frame != null && ((!border.Frame.HasValue) || border.Frame.Value))
        {
            borderStyle = "ridge";
        }

        //if (isTableCell && border.Space != null && border.Space.Value > 0)

        styles.Add($"{cssAttribute}: {borderWidth} {borderStyle} #{borderColor};");
    }

}
