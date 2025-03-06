namespace DocSharp.Binary.OpenXmlLib.WordprocessingML
{
    public class WordprocessingDocument : OpenXmlPackage
    {
        protected WordprocessingDocumentType _documentType;
        protected CustomXmlPropertiesPart _customFilePropertiesPart;
        protected MainDocumentPart _mainDocumentPart;

        protected WordprocessingDocument(string fileName, WordprocessingDocumentType type)
            : base(fileName)
        {
            switch (type)
            {
                case WordprocessingDocumentType.MacroEnabledDocument:
                    this._mainDocumentPart = new MainDocumentPart(this, WordprocessingMLContentTypes.MainDocumentMacro);
                    break;
                case WordprocessingDocumentType.Template:
                    this._mainDocumentPart = new MainDocumentPart(this, WordprocessingMLContentTypes.MainDocumentTemplate);
                    break;
                case WordprocessingDocumentType.MacroEnabledTemplate:
                    this._mainDocumentPart = new MainDocumentPart(this, WordprocessingMLContentTypes.MainDocumentMacroTemplate);
                    break;
                default:
                    this._mainDocumentPart = new MainDocumentPart(this, WordprocessingMLContentTypes.MainDocument);
                    break;
            }

            this._documentType = type;
            this.AddPart(this._mainDocumentPart);
        }

        public static WordprocessingDocument Create(string fileName, WordprocessingDocumentType type)
        {
            var doc = new WordprocessingDocument(fileName, type);
            
            return doc;
        }

        public WordprocessingDocumentType DocumentType
        {
            get { return this._documentType; }
            set { this._documentType = value; }
        }

        public CustomXmlPropertiesPart CustomFilePropertiesPart
        {
            get { return this._customFilePropertiesPart; }
        }
        
        public MainDocumentPart MainDocumentPart
        {
            get { return this._mainDocumentPart; }
        }
    }
}
