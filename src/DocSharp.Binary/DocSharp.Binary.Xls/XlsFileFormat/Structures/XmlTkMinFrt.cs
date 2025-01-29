using DocSharp.Binary.StructuredStorage.Reader;

namespace DocSharp.Binary.Spreadsheet.XlsFileFormat.Structures
{
    public class XmlTkMinFrt
    {
        public XmlTkDouble minScale;

        public XmlTkMinFrt(IStreamReader reader)
        {
            this.minScale = new XmlTkDouble(reader);   
        }
    }
}
