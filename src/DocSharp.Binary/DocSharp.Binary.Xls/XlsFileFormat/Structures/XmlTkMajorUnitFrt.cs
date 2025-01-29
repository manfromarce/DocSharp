using DocSharp.Binary.StructuredStorage.Reader;

namespace DocSharp.Binary.Spreadsheet.XlsFileFormat.Structures
{
    public class XmlTkMajorUnitFrt
    {
        public XmlTkDouble majorUnit;

        public XmlTkMajorUnitFrt(IStreamReader reader)
        {
            this.majorUnit = new XmlTkDouble(reader);   
        }
    }
}
