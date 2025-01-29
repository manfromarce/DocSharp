using DocSharp.Binary.StructuredStorage.Reader;

namespace DocSharp.Binary.Spreadsheet.XlsFileFormat.Structures
{
    public class XmlTkDWord
    {
        public XmlTkHeader xtHeader;

        public int dValue;

        public XmlTkDWord(IStreamReader reader)
        {
            this.xtHeader = new XmlTkHeader(reader);

            this.dValue = reader.ReadInt32();
        }
    }
}
