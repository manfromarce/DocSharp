using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DocumentFormat.OpenXml.Wordprocessing;

namespace DocSharp.Docx;

public partial class DocxToRtfConverter
{
    internal void ProcessShading(Shading shading, StringBuilder sb, ShadingType shadingType)
    {
        if (shading.Val != null && shading.Val != ShadingPatternValues.Nil)
        {
            if (shading.Val == ShadingPatternValues.Clear)
            {
                // Just add colors (see below)
            }
            else if (shading.Val == ShadingPatternValues.HorizontalCross)
            {
                sb.Append(shadingType == ShadingType.Paragraph ? "\\bgdkcross" : "\\clbgdkcross");
            }
            else if (shading.Val == ShadingPatternValues.ThinHorizontalCross)
            {
                sb.Append(shadingType == ShadingType.Paragraph ? "\\bgcross" : "\\clbgcross");
            }
            else if (shading.Val == ShadingPatternValues.HorizontalStripe)
            {
                sb.Append(shadingType == ShadingType.Paragraph ? "\\bgdkhoriz" : "\\clbgdkhor");
            }
            else if (shading.Val == ShadingPatternValues.ThinHorizontalStripe)
            {
                sb.Append(shadingType == ShadingType.Paragraph ? "\\bghoriz" : "\\clbghoriz");
            }
            else if (shading.Val == ShadingPatternValues.VerticalStripe)
            {
                sb.Append(shadingType == ShadingType.Paragraph ? "\\bgdkvert" : "\\clbgdkvert");
            }
            else if (shading.Val == ShadingPatternValues.ThinVerticalStripe)
            {
                sb.Append(shadingType == ShadingType.Paragraph ? "\\bgvert" : "\\clbgvert");
            }
            else if (shading.Val == ShadingPatternValues.DiagonalCross)
            {
                sb.Append(shadingType == ShadingType.Paragraph ? "\\bgdkdcross" : "\\clbgdkdcross");
            }
            else if (shading.Val == ShadingPatternValues.ThinDiagonalCross)
            {
                sb.Append(shadingType == ShadingType.Paragraph ? "\\bgdcross" : "\\clbgdcross");
            }
            else if (shading.Val == ShadingPatternValues.DiagonalStripe)
            {
                sb.Append(shadingType == ShadingType.Paragraph ? "\\bgdkbdiag" : "\\clbgdkbdiag");
            }
            else if (shading.Val == ShadingPatternValues.ThinDiagonalStripe)
            {
                sb.Append(shadingType == ShadingType.Paragraph ? "\\bgbdiag" : "\\clbgbdiag");
            }            
            else if (shading.Val == ShadingPatternValues.ReverseDiagonalStripe)
            {
                sb.Append(shadingType == ShadingType.Paragraph ? "\\bgdkfdiag" : "\\clbgdkfdiag");
            }
            else if (shading.Val == ShadingPatternValues.ThinReverseDiagonalStripe)
            {
                sb.Append(shadingType == ShadingType.Paragraph ? "\\bgfdiag" : "\\clbgfdiag");
            }
            else if (shading.Val == ShadingPatternValues.Percent5)
            {
                sb.Append(shadingType == ShadingType.Paragraph ? "\\shading500" : "\\clshdng500");
            }
            else if (shading.Val == ShadingPatternValues.Percent10)
            {
                sb.Append(shadingType == ShadingType.Paragraph ? "\\shading1000" : "\\clshdng1000");
            }
            else if (shading.Val == ShadingPatternValues.Percent12)
            {
                sb.Append(shadingType == ShadingType.Paragraph ? "\\shading1250" : "\\clshdng1250");
            }
            else if (shading.Val == ShadingPatternValues.Percent15)
            {
                sb.Append(shadingType == ShadingType.Paragraph ? "\\shading1500" : "\\clshdng1500");
            }
            else if (shading.Val == ShadingPatternValues.Percent20)
            {
                sb.Append(shadingType == ShadingType.Paragraph ? "\\shading2000" : "\\clshdng2000");
            }
            else if (shading.Val == ShadingPatternValues.Percent25)
            {
                sb.Append(shadingType == ShadingType.Paragraph ? "\\shading2500" : "\\clshdng2500");
            }
            else if (shading.Val == ShadingPatternValues.Percent30)
            {
                sb.Append(shadingType == ShadingType.Paragraph ? "\\shading3000" : "\\clshdng3000");
            }
            else if (shading.Val == ShadingPatternValues.Percent35)
            {
                sb.Append(shadingType == ShadingType.Paragraph ? "\\shading3500" : "\\clshdng3500");
            }
            else if (shading.Val == ShadingPatternValues.Percent37)
            {
                sb.Append(shadingType == ShadingType.Paragraph ? "\\shading3750" : "\\clshdng3750");
            }
            else if (shading.Val == ShadingPatternValues.Percent40)
            {
                sb.Append(shadingType == ShadingType.Paragraph ? "\\shading4000" : "\\clshdng4000");
            }
            else if (shading.Val == ShadingPatternValues.Percent45)
            {
                sb.Append(shadingType == ShadingType.Paragraph ? "\\shading4500" : "\\clshdng4500");
            }
            else if (shading.Val == ShadingPatternValues.Percent50)
            {
                sb.Append(shadingType == ShadingType.Paragraph ? "\\shading5000" : "\\clshdng5000");
            }
            else if (shading.Val == ShadingPatternValues.Percent55)
            {
                sb.Append(shadingType == ShadingType.Paragraph ? "\\shading5500" : "\\clshdng5500");
            }
            else if (shading.Val == ShadingPatternValues.Percent60)
            {
                sb.Append(shadingType == ShadingType.Paragraph ? "\\shading6000" : "\\clshdng6000");
            }
            else if (shading.Val == ShadingPatternValues.Percent62)
            {
                sb.Append(shadingType == ShadingType.Paragraph ? "\\shading6250" : "\\clshdng6250");
            }
            else if (shading.Val == ShadingPatternValues.Percent65)
            {
                sb.Append(shadingType == ShadingType.Paragraph ? "\\shading6500" : "\\clshdng6500");
            }
            else if (shading.Val == ShadingPatternValues.Percent70)
            {
                sb.Append(shadingType == ShadingType.Paragraph ? "\\shading7000" : "\\clshdng7000");
            }
            else if (shading.Val == ShadingPatternValues.Percent75)
            {
                sb.Append(shadingType == ShadingType.Paragraph ? "\\shading7500" : "\\clshdng7500");
            }
            else if (shading.Val == ShadingPatternValues.Percent80)
            {
                sb.Append(shadingType == ShadingType.Paragraph ? "\\shading8000" : "\\clshdng8000");
            }
            else if (shading.Val == ShadingPatternValues.Percent85)
            {
                sb.Append(shadingType == ShadingType.Paragraph ? "\\shading8500" : "\\clshdng8500");
            }
            else if (shading.Val == ShadingPatternValues.Percent87)
            {
                sb.Append(shadingType == ShadingType.Paragraph ? "\\shading8750" : "\\clshdng8750");
            }
            else if (shading.Val == ShadingPatternValues.Percent90)
            {
                sb.Append(shadingType == ShadingType.Paragraph ? "\\shading9000" : "\\clshdng9000");
            }
            else if (shading.Val == ShadingPatternValues.Percent95)
            {
                sb.Append(shadingType == ShadingType.Paragraph ? "\\shading9500" : "\\clshdng9500");
            }
            else if (shading.Val == ShadingPatternValues.Solid)
            {
                sb.Append(shadingType == ShadingType.Paragraph ? "\\shading10000" : "\\clshdng10000");
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
                    sb.Append(shadingType == ShadingType.Paragraph ? $"\\cfpat{colorIndex}" :
                              $"\\clcfpat{colorIndex}");
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
                    sb.Append(shadingType == ShadingType.Paragraph ? $"\\cbpat{colorIndex}" :
                              $"\\clcbpat{colorIndex}");
                }
            }
        }
    }
}
