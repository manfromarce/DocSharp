using DocSharp.Binary.StructuredStorage.Reader;

namespace DocSharp.Binary.Spreadsheet.XlsFileFormat.Structures
{
    public class XmlTkTpb
    {
        public XmlTkBlob textPropsStream;

        public XmlTkTpb(IStreamReader reader)
        {
            this.textPropsStream = new XmlTkBlob(reader);   
        }
    }
}
