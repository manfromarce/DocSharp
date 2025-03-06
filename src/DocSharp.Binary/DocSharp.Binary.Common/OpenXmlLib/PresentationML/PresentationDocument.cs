namespace DocSharp.Binary.OpenXmlLib.PresentationML
{
    public class PresentationDocument : OpenXmlPackage
    {
        protected PresentationPart _presentationPart;
        protected PresentationDocumentType _documentType;

        protected PresentationDocument(string fileName, PresentationDocumentType type)
            : base(fileName)
        {
            switch (type)
            {
                case PresentationDocumentType.MacroEnabledPresentation:
                    this._presentationPart = new PresentationPart(this, PresentationMLContentTypes.PresentationMacro);
                    break;
                case PresentationDocumentType.Template:
                    this._presentationPart = new PresentationPart(this, PresentationMLContentTypes.PresentationTemplate);
                    break;
                case PresentationDocumentType.MacroEnabledTemplate:
                    this._presentationPart = new PresentationPart(this, PresentationMLContentTypes.PresentationMacroTemplate);
                    break;
                case PresentationDocumentType.Slideshow:
                    this._presentationPart = new PresentationPart(this, PresentationMLContentTypes.Slideshow);
                    break;
                case PresentationDocumentType.MacroEnabledSlideshow:
                    this._presentationPart = new PresentationPart(this, PresentationMLContentTypes.SlideshowMacro);
                    break;
                default:
                    this._presentationPart = new PresentationPart(this, PresentationMLContentTypes.Presentation);
                    break;
            }

            this.AddPart(this._presentationPart);
        }

        public static PresentationDocument Create(string fileName, PresentationDocumentType type)
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
