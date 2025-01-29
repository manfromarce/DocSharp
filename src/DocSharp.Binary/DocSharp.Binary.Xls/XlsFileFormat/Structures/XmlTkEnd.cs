using DocSharp.Binary.StructuredStorage.Reader;

namespace DocSharp.Binary.Spreadsheet.XlsFileFormat.Structures
{
    public class XmlTkEnd
    {
        public XmlTkHeader xtHeader;

        public XmlTkEnd(IStreamReader reader)
        {
            this.xtHeader = new XmlTkHeader(reader);   
        }
    }
}
