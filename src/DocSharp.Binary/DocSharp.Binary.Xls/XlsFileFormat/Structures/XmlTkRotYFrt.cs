using DocSharp.Binary.StructuredStorage.Reader;

namespace DocSharp.Binary.Spreadsheet.XlsFileFormat.Structures
{
    public class XmlTkRotYFrt
    {
        public XmlTkDWord rotationY;

        public XmlTkRotYFrt(IStreamReader reader)
        {
            this.rotationY = new XmlTkDWord(reader);   
        }
    }
}
