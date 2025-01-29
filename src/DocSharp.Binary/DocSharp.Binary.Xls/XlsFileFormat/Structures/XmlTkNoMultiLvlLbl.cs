using DocSharp.Binary.StructuredStorage.Reader;

namespace DocSharp.Binary.Spreadsheet.XlsFileFormat.Structures
{
    public class XmlTkNoMultiLvlLbl
    {
        public XmlTkBool fNoMultiLvlLbl;

        public XmlTkNoMultiLvlLbl(IStreamReader reader)
        {
            this.fNoMultiLvlLbl = new XmlTkBool(reader);   
        }
    }
}
