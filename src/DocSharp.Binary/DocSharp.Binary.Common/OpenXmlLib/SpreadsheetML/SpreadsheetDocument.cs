namespace DocSharp.Binary.OpenXmlLib.SpreadsheetML
{
    /// <summary>
    /// Includes some information about the spreadsheetdocument 
    /// </summary>
    public class SpreadsheetDocument : OpenXmlPackage
    {
        protected WorkbookPart workBookPart;
        protected OpenXmlDocumentType _documentType;

        /// <summary>
        /// Ctor 
        /// </summary>
        /// <param name="fileName">Filename of the file which should be written</param>
        protected SpreadsheetDocument(string fileName, OpenXmlDocumentType type)
            : base(fileName)
        {
            switch (type)
            {
                case OpenXmlDocumentType.MacroEnabledDocument:
                    this.workBookPart = new WorkbookPart(this, SpreadsheetMLContentTypes.WorkbookMacro);
                    break;
                case OpenXmlDocumentType.Template:
                    this.workBookPart = new WorkbookPart(this, WordprocessingMLContentTypes.MainDocumentTemplate);
                    break;
                case OpenXmlDocumentType.MacroEnabledTemplate:
                    this.workBookPart = new WorkbookPart(this, WordprocessingMLContentTypes.MainDocumentMacroTemplate);
                    break;
                default:
                    this.workBookPart = new WorkbookPart(this, SpreadsheetMLContentTypes.Workbook);
                    break;
            }
            this._documentType = type;
            this.AddPart(this.workBookPart);
        }

        /// <summary>
        /// creates a new excel document with the choosen filename 
        /// </summary>
        /// <param name="fileName">The name of the file which should be written</param>
        /// <returns>The object itself</returns>
        public static SpreadsheetDocument Create(string fileName, OpenXmlDocumentType type)
        {
            var spreadsheet = new SpreadsheetDocument(fileName, type);
            return spreadsheet;
        }

        public OpenXmlDocumentType DocumentType
        {
            get { return this._documentType; }
            set { this._documentType = value; }
        }

        /// <summary>
        /// returns the workbookPart from the new excel document 
        /// </summary>
        public WorkbookPart WorkbookPart
        {
            get { return this.workBookPart; }
        }
    }
}
