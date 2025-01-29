using DocSharp.Binary.StructuredStorage.Reader;

namespace DocSharp.Binary.Spreadsheet.XlsFileFormat.Structures
{
    public class XmlTkStyle
    {
        public XmlTkDWord chartStyle;

        public XmlTkStyle(IStreamReader reader)
        {
            this.chartStyle = new XmlTkDWord(reader);  
        }
    }
}
