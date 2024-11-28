namespace b2xtranslator.OpenXmlLib.PresentationML
{
    public class PresentationDocument : OpenXmlPackage
    {
        protected PresentationPart _presentationPart;
        protected OpenXmlDocumentType _documentType;

        protected PresentationDocument(string fileName, OpenXmlDocumentType type)
            : base(fileName)
        {
            switch (type)
            {
                case OpenXmlDocumentType.Document:
                    this._presentationPart = new PresentationPart(this, PresentationMLContentTypes.Presentation);
                    break;
                case OpenXmlDocumentType.MacroEnabledDocument:
                    this._presentationPart = new PresentationPart(this, PresentationMLContentTypes.PresentationMacro);
                    break;
                case OpenXmlDocumentType.Template:
                    break;
                case OpenXmlDocumentType.MacroEnabledTemplate:
                    break;
            }

            this.AddPart(this._presentationPart);
        }

        public static PresentationDocument Create(string fileName, OpenXmlDocumentType type)
        {
            var presentation = new PresentationDocument(fileName, type);

            return presentation;
        }

        public PresentationPart PresentationPart
        {
            get { return this._presentationPart; }
        }
    }
}
