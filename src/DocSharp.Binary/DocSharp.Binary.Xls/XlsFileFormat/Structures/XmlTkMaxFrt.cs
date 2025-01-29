using DocSharp.Binary.StructuredStorage.Reader;

namespace DocSharp.Binary.Spreadsheet.XlsFileFormat.Structures
{
    public class XmlTkMaxFrt
    {
        public XmlTkDouble maxScale;

        public XmlTkMaxFrt(IStreamReader reader)
        {
            this.maxScale = new XmlTkDouble(reader);   
        }
    }
}
