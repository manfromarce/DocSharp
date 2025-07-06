using System.Collections.Generic;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;

namespace DocSharp.Renderer
{
    internal static class ChildElementsExtensions
    {
        public static IEnumerable<OpenXmlCompositeElement> RenderableChildren(this OpenXmlElement xmlElement)
        {
            return xmlElement
                .ChildElements
                .Where(c => c is Paragraph || c is Table || c is SdtBlock)
                .SelectMany(c =>
                {
                    return c switch
                    {
                        SdtBlock block => block.SdtContentBlock.ChildElements.OfType<OpenXmlCompositeElement>().ToArray(),
                        _ => new[] { c }
                    };
                })
                .Cast<OpenXmlCompositeElement>();
        }

        public static IEnumerable<T> ChildsOfType<T>(this OpenXmlElement xmlElement)
            where T : OpenXmlElement
        {
            return xmlElement.ChildElements.OfType<T>();
        }
    }
}
