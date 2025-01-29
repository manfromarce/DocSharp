using DocSharp.Binary.StructuredStorage.Reader;

namespace DocSharp.Binary.Spreadsheet.XlsFileFormat.Structures
{
    public class XmlTkRotXFrt
    {
        public XmlTkDWord rotationX;

        public XmlTkRotXFrt(IStreamReader reader)
        {
            this.rotationX = new XmlTkDWord(reader);   
        }
    }
}
