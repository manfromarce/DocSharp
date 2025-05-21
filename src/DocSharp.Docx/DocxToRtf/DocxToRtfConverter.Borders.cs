using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DocumentFormat.OpenXml.Wordprocessing;

namespace DocSharp.Docx;

public partial class DocxToRtfConverter : DocxToTextConverterBase
{
    // This function is used for page, paragraph and table borders.
    internal void ProcessBorder(BorderType border, StringBuilder sb)
    {
        if (border.Val != null)
        {
            sb.Append(RtfBorderMapper.GetBorderType(border.Val.Value));
        }
        if (border.Size != null)
        {
            // Open XML uses 1/8 points for border width, while RTF uses twips
            double twipsSize = Math.Round(border.Size.Value * 2.5);
            sb.Append($"\\brdrw{twipsSize}");
        }
        if (border.Space != null)
        {
            // Open XML uses points for border spacing, while RTF uses twips
            uint twipsSize = border.Space.Value * 20;
            sb.Append($"\\brsp{twipsSize}");
        }
        if (border.Color != null && !string.IsNullOrEmpty(border.Color?.Value))
        {
            if (border.Color.Value.Equals("auto", StringComparison.OrdinalIgnoreCase))
            {
                sb.Append(@"\brdrcf0");
            }
            else
            {
                colors.TryAddAndGetIndex(border.Color.Value, out int colorIndex);
                sb.Append($"\\brdrcf{colorIndex}");
            }
        }
        if (border.Shadow != null && ((!border.Shadow.HasValue) || border.Shadow.Value))
        {
            sb.Append(@"\brdrsh");
        }
        if (border.Frame != null && ((!border.Frame.HasValue) || border.Frame.Value))
        {
            sb.Append(@"\brdrframe");
        }
    }
}
