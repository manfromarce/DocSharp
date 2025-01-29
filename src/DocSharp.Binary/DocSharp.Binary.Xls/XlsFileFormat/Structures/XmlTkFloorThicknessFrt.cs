using DocSharp.Binary.StructuredStorage.Reader;

namespace DocSharp.Binary.Spreadsheet.XlsFileFormat.Structures
{
    public class XmlTkFloorThicknessFrt
    {
        public XmlTkDWord floorThickness;

        public XmlTkFloorThicknessFrt(IStreamReader reader)
        {
            this.floorThickness = new XmlTkDWord(reader);   
        }
    }
}
