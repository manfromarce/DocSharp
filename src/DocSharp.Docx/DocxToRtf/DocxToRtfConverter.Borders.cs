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
            double twipsSize = Math.Round(border.Size.Value * 2.5);
            sb.Write($"\\brdrw{twipsSize}");
        }
        if (border.Space != null)
        {
            // Open XML uses points for border spacing, while RTF uses twips
            uint twipsSize = border.Space.Value * 20;
            sb.Write($"\\brsp{twipsSize}");
        }
        if (border.Color != null && !string.IsNullOrEmpty(border.Color?.Value))
        {
            if (border.Color.Value.Equals("auto", StringComparison.OrdinalIgnoreCase))
            {
                sb.Write(@"\brdrcf0");
            }
            else
            {
                colors.TryAddAndGetIndex(border.Color.Value, out int colorIndex);
                sb.Write($"\\brdrcf{colorIndex}");
            }
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
