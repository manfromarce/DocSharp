using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;

namespace DocSharp.Rtf;

internal static class RtfShadingMapper
{
    internal static EnumValue<ShadingPatternValues>? GetShadingType(string word, int? value)
    {
        switch (word)
        {
            case "chbgdkcross":
            case "bgdkcross":
            case "trbgdkcross":
            case "clbgdkcross":
                return ShadingPatternValues.HorizontalCross;
            case "chbgcross":
            case "bgcross":
            case "trbgcross":
            case "clbgcross":
                return ShadingPatternValues.ThinHorizontalCross;
            case "chbgdkhoriz":
            case "trbgdkhoriz":
            case "clbgdkhoriz":
                return ShadingPatternValues.HorizontalStripe;
            case "chbghoriz":
            case "bghoriz":
            case "trbghoriz":
            case "clbghoriz":
                return ShadingPatternValues.ThinHorizontalStripe;
            case "chbgdkvert":
            case "bgdkvert":
            case "trbgdkvert":
            case "clbgdkvert":
                return ShadingPatternValues.VerticalStripe;
            case "chbgvert":
            case "bgvert":
            case "trbgvert":
            case "clbgvert":
                return ShadingPatternValues.ThinVerticalStripe;
            case "chbgdkdcross":
            case "bgdkdcross":
            case "trbgdkdcross":
            case "clbgdkdcross":
                return ShadingPatternValues.DiagonalCross;
            case "chbgdcross":
            case "bgdcross":
            case "trbgdcross":
            case "clbgdcross":
                return ShadingPatternValues.ThinDiagonalCross;
            case "chbgdkbdiag":
            case "bgdkbdiag":
            case "trbgdkbdiag":
            case "clbgdkbdiag":
                return ShadingPatternValues.DiagonalStripe;
            case "chbgbdiag":
            case "bgbdiag":
            case "trbgbdiag":
            case "clbgbdiag":
                return ShadingPatternValues.ThinDiagonalStripe;
            case "chbgdkfdiag":
            case "bgdkfdiag":
            case "trbgdkfdiag":
            case "clbgdkfdiag":
                return ShadingPatternValues.ReverseDiagonalStripe;
            case "chbgfdiag":
            case "bgfdiag":
            case "trbgfdiag":
            case "clbgfdiag":
                return ShadingPatternValues.ThinReverseDiagonalStripe;
            case "chshdng":
            case "shdng":
            case "trshdng":
            case "clshdng":
                if (value.HasValue)
                {
                    if (value.Value == 0)
                        return ShadingPatternValues.Clear;
                        // return ShadingPatternValues.Nil;
                    else if (value.Value == 10000)
                        return ShadingPatternValues.Solid;
                    else if (value.Value >= 9500)
                        return ShadingPatternValues.Percent95;
                    else if (value.Value >= 9000)
                        return ShadingPatternValues.Percent90;
                    else if (value.Value >= 8750)
                        return ShadingPatternValues.Percent87;
                    else if (value.Value >= 8500)
                        return ShadingPatternValues.Percent85;
                    else if (value.Value >= 8000)
                        return ShadingPatternValues.Percent80;
                     else if (value.Value >= 7500)
                        return ShadingPatternValues.Percent75;
                    else if (value.Value >= 7000)
                        return ShadingPatternValues.Percent70;
                    else if (value.Value >= 6500)
                        return ShadingPatternValues.Percent65;
                    else if (value.Value >= 6250)
                        return ShadingPatternValues.Percent62;
                    else if (value.Value >= 6000)
                        return ShadingPatternValues.Percent60;
                    else if (value.Value >= 5500)
                        return ShadingPatternValues.Percent55;
                    else if (value.Value >= 5000)
                        return ShadingPatternValues.Percent50;
                    else if (value.Value >= 4500)
                        return ShadingPatternValues.Percent45;
                    else if (value.Value >= 4000)
                        return ShadingPatternValues.Percent40;
                    else if (value.Value >= 3750)
                        return ShadingPatternValues.Percent37;
                    else if (value.Value >= 3500)
                        return ShadingPatternValues.Percent35;
                    else if (value.Value >= 3000)
                        return ShadingPatternValues.Percent30;
                    else if (value.Value >= 2500)
                        return ShadingPatternValues.Percent25;
                    else if (value.Value >= 2000)
                        return ShadingPatternValues.Percent20;
                    else if (value.Value >= 1500)
                        return ShadingPatternValues.Percent15;
                    else if (value.Value >= 1250)
                        return ShadingPatternValues.Percent12;
                    else if (value.Value >= 1000)
                        return ShadingPatternValues.Percent10;
                    else 
                        return ShadingPatternValues.Percent5;
                }
                break;
        }
        return null;
    }
}