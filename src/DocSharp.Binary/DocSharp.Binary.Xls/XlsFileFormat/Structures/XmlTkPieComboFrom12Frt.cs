using DocSharp.Binary.StructuredStorage.Reader;

namespace DocSharp.Binary.Spreadsheet.XlsFileFormat.Structures
{
    public class XmlTkPieComboFrom12Frt
    {
        public XmlTkBool fPieCombo;

        public XmlTkPieComboFrom12Frt(IStreamReader reader)
        {
            this.fPieCombo = new XmlTkBool(reader);   
        }
    }
}
