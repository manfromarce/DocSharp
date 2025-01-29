using DocSharp.Binary.StructuredStorage.Reader;

namespace DocSharp.Binary.Spreadsheet.XlsFileFormat.Structures
{
    public class XmlTkMinorUnitFrt
    {
        public XmlTkDouble minorUnit;

        public XmlTkMinorUnitFrt(IStreamReader reader)
        {
            this.minorUnit = new XmlTkDouble(reader);   
        }
    }
}
