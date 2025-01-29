using DocSharp.Binary.StructuredStorage.Reader;

namespace DocSharp.Binary.Spreadsheet.XlsFileFormat.Structures
{
    public class XmlTkTickLabelSkipFrt
    {
        public XmlTkDWord nInternal;

        public XmlTkTickLabelSkipFrt(IStreamReader reader)
        {
            this.nInternal = new XmlTkDWord(reader);   
        }
    }
}
