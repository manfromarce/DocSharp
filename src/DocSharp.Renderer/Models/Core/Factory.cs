using System.Collections.Generic;
using System.Linq;
using DocumentFormat.OpenXml;
using DocSharp.Renderer.Core;
using DocSharp.Renderer.Models.Styles;
using Word = DocumentFormat.OpenXml.Wordprocessing;

namespace DocSharp.Renderer.Models
{
    internal static class Factory
    {
        public static IEnumerable<PageContextElement> CreatePageElements(
            this IEnumerable<OpenXmlCompositeElement> openXmlComposites,
            IImageAccessor imageAccessor,
            IStyleFactory styleFactory)
        {
            return openXmlComposites
                .Select(xml =>
                {
                    var e = xml.CreateElement(imageAccessor, styleFactory);
                    return e;
                });
        }

        public static PageContextElement CreateElement(
            this OpenXmlCompositeElement openXmlComposite,
            IImageAccessor imageAccessor,
            IStyleFactory styleFactory)
        {
            return openXmlComposite switch
            {
                Word.Paragraph p => Paragraphs.Builders.ParagraphFactory.Create(p, imageAccessor, styleFactory),
                Word.Table t => Tables.Builders.TableFactory.Create(t, imageAccessor, styleFactory),
                _ => throw new RendererException("Unhandled element")
            };
        }
    }
}
