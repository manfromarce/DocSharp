using DocSharp.Binary.StructuredStorage.Reader;

namespace DocSharp.Binary.Spreadsheet.XlsFileFormat.Structures
{
    public class XmlTkThemeOverride
    {
        public XmlTkBlob rgThemeOverride;

        public XmlTkThemeOverride(IStreamReader reader)
        {
            this.rgThemeOverride = new XmlTkBlob(reader);   
        }
    }
}
