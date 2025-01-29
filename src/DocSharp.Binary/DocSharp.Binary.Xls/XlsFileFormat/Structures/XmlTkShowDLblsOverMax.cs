using DocSharp.Binary.StructuredStorage.Reader;

namespace DocSharp.Binary.Spreadsheet.XlsFileFormat.Structures
{
    public class XmlTkShowDLblsOverMax
    {
        public XmlTkBool fVDLOverMax;

        public XmlTkShowDLblsOverMax(IStreamReader reader)
        {
            this.fVDLOverMax = new XmlTkBool(reader);     
        }
    }
}
