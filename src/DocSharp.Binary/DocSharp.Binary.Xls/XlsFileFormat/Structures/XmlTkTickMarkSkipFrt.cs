using DocSharp.Binary.StructuredStorage.Reader;

namespace DocSharp.Binary.Spreadsheet.XlsFileFormat.Structures
{
    public class XmlTkTickMarkSkipFrt
    {
        public XmlTkDWord nInternal;

        public XmlTkTickMarkSkipFrt(IStreamReader reader)
        {
            this.nInternal = new XmlTkDWord(reader);   
        }
    }
}
