using DocSharp.Binary.StructuredStorage.Reader;

namespace DocSharp.Binary.Spreadsheet.XlsFileFormat.Structures
{
    public class XmlTkHeightPercent
    {
        public XmlTkDouble heightPercent;

        public XmlTkHeightPercent(IStreamReader reader)
        {
            this.heightPercent = new XmlTkDouble(reader);   
        }
    }
}
