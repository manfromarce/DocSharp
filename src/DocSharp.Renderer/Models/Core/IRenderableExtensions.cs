using System.Collections.Generic;
using DocSharp.Renderer.Core;

namespace DocSharp.Renderer.Models
{
    internal static class IRenderableExtensions
    {
        public static void Render(this IEnumerable<IRenderable> elements, IRenderer renderer)
        {
            foreach (var e in elements)
            {
                e.Render(renderer);
            }
        }
    }
}
