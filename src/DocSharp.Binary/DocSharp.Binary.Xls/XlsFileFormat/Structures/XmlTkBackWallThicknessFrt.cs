using DocSharp.Binary.StructuredStorage.Reader;

namespace DocSharp.Binary.Spreadsheet.XlsFileFormat.Structures
{
    public class XmlTkBackWallThicknessFrt
    {
        public XmlTkDWord wallThickness;

        public XmlTkBackWallThicknessFrt(IStreamReader reader)
        {
            this.wallThickness = new XmlTkDWord(reader);   
        }
    }
}
