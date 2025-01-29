using DocSharp.Binary.CommonTranslatorLib;
using DocSharp.Binary.OpenXmlLib;
using System.Xml;

namespace DocSharp.Binary.PresentationMLMapping
{
    public abstract class PresentationMapping<T> :
        AbstractOpenXmlMapping,
        IMapping<T>
        where T : IVisitable
    {
        protected ConversionContext _ctx;
        public ContentPart targetPart;
        
        public PresentationMapping(ConversionContext ctx, ContentPart targetPart)
            : base(XmlWriter.Create(targetPart.GetStream(), ctx.WriterSettings))
        {
            this._ctx = ctx;
            this.targetPart = targetPart;
        }

        public abstract void Apply(T mapElement);
    }
}
