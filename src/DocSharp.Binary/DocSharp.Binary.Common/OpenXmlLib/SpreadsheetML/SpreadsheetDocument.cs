using System.IO;

namespace DocSharp.Binary.OpenXmlLib.SpreadsheetML
{
    /// <summary>
    /// Includes some information about the spreadsheetdocument 
    /// </summary>
    public class SpreadsheetDocument : OpenXmlPackage
    {
        protected WorkbookPart workBookPart;
        protected SpreadsheetDocumentType _documentType;

        protected SpreadsheetDocument(string fileName, SpreadsheetDocumentType type)
            : base(fileName)
        {
            Initialize(type);
        }
        
        protected SpreadsheetDocument(Stream stream, SpreadsheetDocumentType type)
            : base(stream)
        {
            Initialize(type);
        }

        private void Initialize(SpreadsheetDocumentType type)
        {
            switch (type)
            {
                case SpreadsheetDocumentType.MacroEnabledWorkbook:
                    this.workBookPart = new WorkbookPart(this, SpreadsheetMLContentTypes.WorkbookMacro);
                    break;
                case SpreadsheetDocumentType.Template:
                    this.workBookPart = new WorkbookPart(this, SpreadsheetMLContentTypes.WorkbookTemplate);
                    break;
                case SpreadsheetDocumentType.MacroEnabledTemplate:
                    this.workBookPart = new WorkbookPart(this, SpreadsheetMLContentTypes.WorkbookMacroTemplate);
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
        public static SpreadsheetDocument Create(string fileName, SpreadsheetDocumentType type)
        {
            var spreadsheet = new SpreadsheetDocument(fileName, type);
            return spreadsheet;
        }
        
        /// <summary>
        /// creates a new excel document
        /// </summary>
        /// <param name="stream">Stream which should be written</param>
        /// <returns>The object itself</returns>
        public static SpreadsheetDocument Create(Stream stream, SpreadsheetDocumentType type)
        {
            var doc = new SpreadsheetDocument(stream, type);
            
            return doc;
        }

        public SpreadsheetDocumentType DocumentType
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
