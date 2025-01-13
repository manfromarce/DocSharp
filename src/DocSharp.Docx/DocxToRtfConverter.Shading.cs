using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DocumentFormat.OpenXml.Wordprocessing;

namespace DocSharp.Docx;

public partial class DocxToRtfConverter
{
    internal void ProcessShading(Shading shading, StringBuilder sb)
    {
        if (shading.Val != null && shading.Val != ShadingPatternValues.Nil)
        {
            if (shading.Val == ShadingPatternValues.Clear)
            {
                // Just add colors (see below)
            }
            else if (shading.Val == ShadingPatternValues.HorizontalCross)
            {
                sb.Append("\\bgdkcross");
            }
            else if (shading.Val == ShadingPatternValues.ThinHorizontalCross)
            {
                sb.Append("\\bgcross");
            }
            else if (shading.Val == ShadingPatternValues.HorizontalStripe)
            {
                sb.Append("\\bgdkhoriz");
            }
            else if (shading.Val == ShadingPatternValues.ThinHorizontalStripe)
            {
                sb.Append("\\bghoriz");
            }
            else if (shading.Val == ShadingPatternValues.VerticalStripe)
            {
                sb.Append("\\bgdkvert");
            }
            else if (shading.Val == ShadingPatternValues.ThinVerticalStripe)
            {
                sb.Append("\\bgvert");
            }
            else if (shading.Val == ShadingPatternValues.DiagonalCross)
            {
                sb.Append("\\bgdkdcross");
            }
            else if (shading.Val == ShadingPatternValues.ThinDiagonalCross)
            {
                sb.Append("\\bgdcross");
            }
            else if (shading.Val == ShadingPatternValues.DiagonalStripe)
            {
                sb.Append("\\bgdkbdiag");
            }
            else if (shading.Val == ShadingPatternValues.ThinDiagonalStripe)
            {
                sb.Append("\\bgbdiag");
            }            
            else if (shading.Val == ShadingPatternValues.ReverseDiagonalStripe)
            {
                sb.Append("\\bgdkfdiag");
            }
            else if (shading.Val == ShadingPatternValues.ThinReverseDiagonalStripe)
            {
                sb.Append("\\bgfdiag");
            }
            else if (shading.Val == ShadingPatternValues.Percent5)
            {
                sb.Append("\\shading500");
            }
            else if (shading.Val == ShadingPatternValues.Percent10)
            {
                sb.Append("\\shading1000");
            }
            else if (shading.Val == ShadingPatternValues.Percent12)
            {
                sb.Append("\\shading1250");
            }
            else if (shading.Val == ShadingPatternValues.Percent15)
            {
                sb.Append("\\shading2000");
            }
            else if (shading.Val == ShadingPatternValues.Percent20)
            {
                sb.Append("\\shading2000");
            }
            else if (shading.Val == ShadingPatternValues.Percent25)
            {
                sb.Append("\\shading2500");
            }
            else if (shading.Val == ShadingPatternValues.Percent30)
            {
                sb.Append("\\shading3000");
            }
            else if (shading.Val == ShadingPatternValues.Percent35)
            {
                sb.Append("\\shading3500");
            }
            else if (shading.Val == ShadingPatternValues.Percent37)
            {
                sb.Append("\\shading3750");
            }
            else if (shading.Val == ShadingPatternValues.Percent40)
            {
                sb.Append("\\shading4000");
            }
            else if (shading.Val == ShadingPatternValues.Percent45)
            {
                sb.Append("\\shading4500");
            }
            else if (shading.Val == ShadingPatternValues.Percent50)
            {
                sb.Append("\\shading5000");
            }
            else if (shading.Val == ShadingPatternValues.Percent55)
            {
                sb.Append("\\shading5500");
            }
            else if (shading.Val == ShadingPatternValues.Percent60)
            {
                sb.Append("\\shading6000");
            }
            else if (shading.Val == ShadingPatternValues.Percent62)
            {
                sb.Append("\\shading6250");
            }
            else if (shading.Val == ShadingPatternValues.Percent65)
            {
                sb.Append("\\shading6500");
            }
            else if (shading.Val == ShadingPatternValues.Percent70)
            {
                sb.Append("\\shading7000");
            }
            else if (shading.Val == ShadingPatternValues.Percent75)
            {
                sb.Append("\\shading7500");
            }
            else if (shading.Val == ShadingPatternValues.Percent80)
            {
                sb.Append("\\shading8000");
            }
            else if (shading.Val == ShadingPatternValues.Percent85)
            {
                sb.Append("\\shading8500");
            }
            else if (shading.Val == ShadingPatternValues.Percent87)
            {
                sb.Append("\\shading8750");
            }
            else if (shading.Val == ShadingPatternValues.Percent90)
            {
                sb.Append("\\shading9000");
            }
            else if (shading.Val == ShadingPatternValues.Percent95)
            {
                sb.Append("\\shading9500");
            }
            else if (shading.Val == ShadingPatternValues.Solid)
            {
                sb.Append("\\shading10000");
            }

            // If shading is not a solid color, check the second/foreground solor
            if (shading.Val != ShadingPatternValues.Clear && shading.Color != null &&
                !string.IsNullOrEmpty(shading.Color.Value))
            {
                if (shading.Color.Value.Equals("auto", StringComparison.OrdinalIgnoreCase))
                {
                    // Interpreted as black or default document background color
                }
                else
                {
                    colors.TryAddAndGetIndex(shading.Color.Value, out int colorIndex);
                    sb.Append($"\\cfpat{colorIndex}");
                }
            }

            // Check the main background color
            if (shading.Fill != null && !string.IsNullOrEmpty(shading.Fill.Value))
            {
                if (shading.Fill.Value.Equals("auto", StringComparison.OrdinalIgnoreCase))
                {
                    // Interpreted as transparent (no background)
                }
                else
                {
                    colors.TryAddAndGetIndex(shading.Fill.Value, out int colorIndex);
                    sb.Append($"\\cbpat{colorIndex}");
                }
            }
        }
    }
}
