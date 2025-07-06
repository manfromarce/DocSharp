using DocSharp.Renderer.Core;

namespace DocSharp.Renderer.Models
{
    internal interface IPageRenderable
    {
        void Render(IRendererPage page);
    }
}
