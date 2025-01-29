using System.Xml;
using DocSharp.Binary.CommonTranslatorLib;
using DocSharp.Binary.OpenXmlLib;
using DocSharp.Binary.Spreadsheet.XlsFileFormat;


namespace DocSharp.Binary.SpreadsheetMLMapping
{
    public abstract class ExcelMapping :
        AbstractOpenXmlMapping,
        IMapping<XlsDocument>
    {
        protected XlsDocument xls;
        protected ExcelContext xlscon;

        public ExcelMapping(ExcelContext xlscon, OpenXmlPart targetPart)
            : base(XmlWriter.Create(targetPart.GetStream(), xlscon.WriterSettings))
        {
            this.xlscon = xlscon; 
        }

        public abstract void Apply(XlsDocument xls); 
        }

    
}
