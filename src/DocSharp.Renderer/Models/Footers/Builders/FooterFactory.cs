using DocSharp.Renderer.Core;
using DocSharp.Renderer.Models.Common;
using DocSharp.Renderer.Models.Styles;
using Word = DocumentFormat.OpenXml.Wordprocessing;

namespace DocSharp.Renderer.Models.Footers.Builders
{
    internal static class FooterFactory
    {
        public static FooterBase CreateInheritedFooter(PageMargin previousSectionMargin)
        {
            return new NoFooter(previousSectionMargin);
        }

        public static FooterBase CreateFooter(
            this Word.Footer wordFooter,
            PageMargin pageMargin,
            IImageAccessor imageAccessor,
            IStyleFactory styleFactory)
        {
            if(wordFooter == null)
            {
                return new NoFooter(pageMargin);
            }

            var childElements = wordFooter.RenderableChildren().CreatePageElements(imageAccessor, styleFactory);
            return new Footer(childElements, pageMargin);
        }
    }
}
