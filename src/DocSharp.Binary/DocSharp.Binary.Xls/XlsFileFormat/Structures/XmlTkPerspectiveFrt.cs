using DocSharp.Binary.StructuredStorage.Reader;

namespace DocSharp.Binary.Spreadsheet.XlsFileFormat.Structures
{
    public class XmlTkPerspectiveFrt
    {
       public XmlTkDWord perspectiveAngle;

        public XmlTkPerspectiveFrt(IStreamReader reader)
        {
            this.perspectiveAngle = new XmlTkDWord(reader);   
        }
    }
}
