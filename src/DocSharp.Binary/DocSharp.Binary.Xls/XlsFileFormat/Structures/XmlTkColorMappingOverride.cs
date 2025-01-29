using DocSharp.Binary.StructuredStorage.Reader;

namespace DocSharp.Binary.Spreadsheet.XlsFileFormat.Structures
{
    public class XmlTkColorMappingOverride
    {
        public XmlTkBlob rgThemeOverride;

        public XmlTkColorMappingOverride(IStreamReader reader)
        {
            this.rgThemeOverride = new XmlTkBlob(reader);   
        }
    }
}
