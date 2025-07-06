using DocSharp.Renderer.Core;

namespace DocSharp.Renderer.Models
{
    internal interface IRenderable
    {
        void Render(IRenderer renderer);
    }
}
