using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DocSharp.Writers;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;
using V = DocumentFormat.OpenXml.Vml;
using O = DocumentFormat.OpenXml.Vml.Office;
using DocSharp.IO;
using DocumentFormat.OpenXml.Packaging;
using System.IO;
using System.Globalization;

namespace DocSharp.Docx;

public partial class DocxToRtfConverter : DocxToStringWriterBase<RtfStringWriter>
{    
    internal override void ProcessDocumentBackground(DocumentBackground documentBackground, RtfStringWriter sb)
    {
        if (documentBackground.Background != null)
        {
            ProcessBackground(documentBackground.Background, sb);
        }
        else if (documentBackground.Color?.Value != null)
        {
            int? bgr = ColorHelpers.HexToBgr(documentBackground.Color);
            if (bgr != null)
            {
                sb.WriteLine(@"{\*\background {\shp{\*\shpinst\shpleft0\shptop0\shpright0\shpbottom0\shpfhdr0\shpbxmargin\shpbxignore\shpbymargin\shpbyignore\shpwr0\shpwrk0\shpfblwtxt1\shpz0");
                sb.WriteShapeProperty("shapeType", "1");
                sb.WriteShapeProperty("fFlipH", "0");
                sb.WriteShapeProperty("fFlipV", "0");
                sb.WriteShapeProperty("fillColor", bgr.Value);
                sb.WriteShapeProperty("fFilled", "1");
                sb.WriteShapeProperty("lineWidth", "0");
                sb.WriteShapeProperty("fLine", "0");
                sb.WriteShapeProperty("bWMode", "9");
                sb.WriteShapeProperty("fBackground", "1");
                sb.WriteShapeProperty("fLayoutInCell", "1");
                sb.WriteLine("}}}");

                sb.WriteLine(@"\viewbksp1");
            }
        }        
    }

    internal void ProcessBackground(V.Background background, RtfStringWriter sb)
    {
        if (background.Fill is V.Fill fill && fill.Type != null && fill.Type.HasValue)
        {
            sb.WriteLine(@"{\*\background {\shp{\*\shpinst\shpleft0\shptop0\shpright0\shpbottom0\shpfhdr0\shpbxmargin\shpbxignore\shpbymargin\shpbyignore\shpwr0\shpwrk0\shpfblwtxt1\shpz0");

            sb.WriteShapeProperty("shapeType", "1");
            sb.WriteShapeProperty("fFlipH", "0");
            sb.WriteShapeProperty("fFlipV", "0");
            sb.WriteShapeProperty("lineWidth", "0");
            sb.WriteShapeProperty("fLine", "0");
            sb.WriteShapeProperty("fBackground", "1");
            sb.WriteShapeProperty("fLayoutInCell", "1");

            ProcessCommonVmlProperties(background, sb, fillByDefault: true);

            sb.WriteLine("}}}");

            sb.WriteLine(@"\viewbksp1");
        }
    }
}
