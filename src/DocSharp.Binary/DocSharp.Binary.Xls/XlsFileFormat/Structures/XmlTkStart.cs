using DocSharp.Binary.StructuredStorage.Reader;

namespace DocSharp.Binary.Spreadsheet.XlsFileFormat.Structures
{
    public class XmlTkStart
    {
        public XmlTkHeader xtHeader;

        public XmlTkStart(IStreamReader reader)
        {
            this.xtHeader = new XmlTkHeader(reader);   
        }
    }
}
