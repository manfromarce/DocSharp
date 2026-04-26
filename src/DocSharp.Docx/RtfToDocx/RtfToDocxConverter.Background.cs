using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DocumentFormat.OpenXml.Wordprocessing;
using DocSharp.Writers;
using System.IO;
using DocumentFormat.OpenXml.Packaging;
using System.Globalization;
using DocSharp.Helpers;
using System.Xml;
using System.Diagnostics;
using DocumentFormat.OpenXml;
using DocSharp.Rtf;

namespace DocSharp.Docx;

public partial class RtfToDocxConverter : ITextToDocxConverter
{
    private void ProcessBackgroundDestination(RtfDestination background)
    {
        if (background.Tokens.FirstOrDefault() is RtfDestination dest && dest.Name.Equals("shp", StringComparison.OrdinalIgnoreCase))
        {
            ProcessBackgroundShape(dest);
        }
    }

    private void ProcessBackgroundShape(RtfDestination shp)
    {
        if (shp.Tokens.FirstOrDefault() is RtfDestination dest && dest.Name.Equals("shpinst", StringComparison.OrdinalIgnoreCase))
        {
            if (ReadShapePropertyAsBool(dest, "fBackground") == true && ReadShapePropertyAsBool(dest, "fFilled") == true)
            {
                if (ReadShapePropertyAsLong(dest, "fillColor") is long fillColorBgr)
                {
                    string fillColorHex = ColorHelpers.BgrToHex(fillColorBgr).TrimStart('#'); // in this context, the color is not recognized if it has the leading #
                    if (ColorHelpers.IsValidHexString(fillColorHex)) // if the BGR value is in unexpected range, it produces an invalid hex string because of bitwise operations in the conversion method
                    {
                        mainPart.Document ??= new Document();
                        mainPart.Document.DocumentBackground = new DocumentBackground() { Color = fillColorHex };
                    }
                }
            }
        }
    }
}