using DocSharp.Binary.StructuredStorage.Reader;

namespace DocSharp.Binary.Spreadsheet.XlsFileFormat.Structures
{
    public class XmlTkSpb
    {
        public XmlTkBlob shapePropsStream;

        public XmlTkSpb(IStreamReader reader)
        {
            this.shapePropsStream = new XmlTkBlob(reader);   
        }
    }
}
