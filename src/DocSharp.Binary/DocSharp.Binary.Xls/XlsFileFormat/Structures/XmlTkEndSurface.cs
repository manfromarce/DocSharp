using DocSharp.Binary.StructuredStorage.Reader;

namespace DocSharp.Binary.Spreadsheet.XlsFileFormat.Structures
{
    public class XmlTkEndSurface
    {
        public XmlTkEnd endSurface;

        public XmlTkEndSurface(IStreamReader reader)
        {
            this.endSurface = new XmlTkEnd(reader);   
        }
    }
}
