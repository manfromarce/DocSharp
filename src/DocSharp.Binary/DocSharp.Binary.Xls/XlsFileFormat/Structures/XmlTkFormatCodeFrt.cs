using DocSharp.Binary.StructuredStorage.Reader;

namespace DocSharp.Binary.Spreadsheet.XlsFileFormat.Structures
{
    public class XmlTkFormatCodeFrt
    {
        public XmlTkString stFormat;

        public XmlTkFormatCodeFrt(IStreamReader reader)
        {
            this.stFormat = new XmlTkString(reader);   
        }
    }
}
