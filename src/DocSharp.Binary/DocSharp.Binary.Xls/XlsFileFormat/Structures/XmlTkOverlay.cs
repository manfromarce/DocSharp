using DocSharp.Binary.StructuredStorage.Reader;

namespace DocSharp.Binary.Spreadsheet.XlsFileFormat.Structures
{
    public class XmlTkOverlay
    {
        public XmlTkBool fOverlay;

        public XmlTkOverlay(IStreamReader reader)
        {
            this.fOverlay = new XmlTkBool(reader);    
        }
    }
}
