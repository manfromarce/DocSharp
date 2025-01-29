using DocSharp.Binary.StructuredStorage.Reader;

namespace DocSharp.Binary.Spreadsheet.XlsFileFormat.Structures
{
    public class XmlTkStartSurface
    {
        public XmlTkStart startSurface;

        public XmlTkStartSurface(IStreamReader reader)
        {
            this.startSurface = new XmlTkStart(reader);   
        }
    }
}
