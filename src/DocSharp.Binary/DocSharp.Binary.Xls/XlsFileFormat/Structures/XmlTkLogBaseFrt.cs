using DocSharp.Binary.StructuredStorage.Reader;

namespace DocSharp.Binary.Spreadsheet.XlsFileFormat.Structures
{
    public class XmlTkLogBaseFrt
    {
        public XmlTkDouble logScale;

        public XmlTkLogBaseFrt(IStreamReader reader)
        {
            this.logScale = new XmlTkDouble(reader);   
        }
    }
}
