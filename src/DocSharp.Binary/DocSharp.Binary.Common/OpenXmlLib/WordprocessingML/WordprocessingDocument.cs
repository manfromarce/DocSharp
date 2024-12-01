namespace b2xtranslator.OpenXmlLib.WordprocessingML
{
    public class WordprocessingDocument : OpenXmlPackage
    {
        protected OpenXmlDocumentType _documentType;
        protected CustomXmlPropertiesPart? _customFilePropertiesPart;
        protected MainDocumentPart _mainDocumentPart;

        protected WordprocessingDocument(string fileName, OpenXmlDocumentType type)
            : base(fileName)
        {
            switch (type)
            {
                case OpenXmlDocumentType.MacroEnabledDocument:
                    this._mainDocumentPart = new MainDocumentPart(this, WordprocessingMLContentTypes.MainDocumentMacro);
                    break;
                case OpenXmlDocumentType.Template:
                    this._mainDocumentPart = new MainDocumentPart(this, WordprocessingMLContentTypes.MainDocumentTemplate);
                    break;
                case OpenXmlDocumentType.MacroEnabledTemplate:
                    this._mainDocumentPart = new MainDocumentPart(this, WordprocessingMLContentTypes.MainDocumentMacroTemplate);
                    break;
                default:
                    this._mainDocumentPart = new MainDocumentPart(this, WordprocessingMLContentTypes.MainDocument);
                    break;
            }

            this._documentType = type;
            this.AddPart(this._mainDocumentPart);
        }

        public static WordprocessingDocument Create(string fileName, OpenXmlDocumentType type)
        {
            var doc = new WordprocessingDocument(fileName, type);
            
            return doc;
        }

        public OpenXmlDocumentType DocumentType
        {
            get { return this._documentType; }
            set { this._documentType = value; }
        }

        public CustomXmlPropertiesPart? CustomFilePropertiesPart
        {
            get { return this._customFilePropertiesPart; }
        }

        
        public MainDocumentPart MainDocumentPart
        {
            get { return this._mainDocumentPart; }
        }
    }
}
