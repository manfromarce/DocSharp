using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DocSharp.Writers;
using DocumentFormat.OpenXml.Wordprocessing;

namespace DocSharp.Docx;

public partial class DocxToRtfConverter : DocxToTextConverterBase<RtfStringWriter>
{
    // This function is used for page, paragraph and table borders.
    internal void ProcessBorder(BorderType border, RtfStringWriter sb)
    {
        if (border.Val != null)
        {
            sb.Write(RtfBorderMapper.GetBorderType(border.Val.Value));
        }
        if (border.Size != null)
        {
            // Open XML uses 1/8 points for border width, while RTF uses twips
            sb.WriteWordWithValue("brdrw", Math.Round(border.Size.Value * 2.5m));
        }
        if (border.Space != null)
        {
            // Open XML uses points for border spacing, while RTF uses twips
            sb.WriteWordWithValue("brsp", Math.Round(border.Space.Value * 20.0m));
        }
        if (ColorHelpers.EnsureHexColor(border.Color?.Value) is string borderColor)
        {
            colors.TryAddAndGetIndex(borderColor, out int colorIndex);
            sb.Write($"\\brdrcf{colorIndex}");            
        }
        else
        {
            //if (border.Color?.Value != null && border.Color.Value.Equals("auto", StringComparison.OrdinalIgnoreCase))
            //{
            sb.Write(@"\brdrcf0");
            //}
        }
        if (border.Shadow != null && ((!border.Shadow.HasValue) || border.Shadow.Value))
        {
            sb.Write(@"\brdrsh");
        }
        if (border.Frame != null && ((!border.Frame.HasValue) || border.Frame.Value))
        {
            sb.Write(@"\brdrframe");
        }
    }
}
