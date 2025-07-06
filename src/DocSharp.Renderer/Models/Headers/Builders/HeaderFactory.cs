using DocSharp.Renderer.Core;
using DocSharp.Renderer.Models.Common;
using DocSharp.Renderer.Models.Styles;
using Word = DocumentFormat.OpenXml.Wordprocessing;

namespace DocSharp.Renderer.Models.Headers.Builders
{
    internal static class HeaderFactory
    {
        public static HeaderBase CreateInheritedHeader(PageMargin pageMargin)
        {
            return new NoHeader(pageMargin);
        }

        public static HeaderBase CreateHeader(
            this Word.Header wordHeader,
            PageMargin pageMargin,
            IImageAccessor imageAccessor,
            IStyleFactory styleFactory)
        {
            if(wordHeader == null)
            {
                return new NoHeader(pageMargin);
            }

            var childElements = wordHeader.RenderableChildren().CreatePageElements(imageAccessor, styleFactory);
            return new Header(childElements, pageMargin);
        }
    }
}
