using DocSharp.Binary.StructuredStorage.Reader;

namespace DocSharp.Binary.Spreadsheet.XlsFileFormat.Structures
{
    public class XmlTkTickLabelPositionFrt
    {
        public XmlTkToken xmltkHigh;

        public XmlTkTickLabelPositionFrt(IStreamReader reader)
        {
            this.xmltkHigh = new XmlTkToken(reader);   
        }
    }
}
