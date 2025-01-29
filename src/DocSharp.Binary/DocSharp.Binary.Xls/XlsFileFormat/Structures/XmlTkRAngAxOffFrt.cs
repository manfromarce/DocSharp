using DocSharp.Binary.StructuredStorage.Reader;

namespace DocSharp.Binary.Spreadsheet.XlsFileFormat.Structures
{
    public class XmlTkRAngAxOffFrt
    {
        public XmlTkBool fRightAngAxOff;

        public XmlTkRAngAxOffFrt(IStreamReader reader)
        {
            this.fRightAngAxOff = new XmlTkBool(reader);   
        }
    }
}
