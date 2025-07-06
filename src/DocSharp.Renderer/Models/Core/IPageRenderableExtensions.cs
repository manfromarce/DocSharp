using System.Collections.Generic;
using DocSharp.Renderer.Core;

namespace DocSharp.Renderer.Models
{
    internal static class IPageRenderableExtensions
    {
        public static void Render(this IEnumerable<IPageRenderable> elements, IRendererPage renderer)
        {
            foreach (var e in elements)
            {
                e.Render(renderer);
            }
        }
    }
}
