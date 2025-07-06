using DocSharp.Renderer.Core;
using DocSharp.Renderer.Models.Common;

namespace DocSharp.Renderer.Models
{
    internal class PageRegion
    {
        public static readonly PageRegion None = new PageRegion(PagePosition.None, Rectangle.Empty);

        public PageRegion(PagePosition pagePosition, Rectangle region)
        {
            this.PagePosition = pagePosition;
            this.Region = region;
        }

        public PagePosition PagePosition { get; }
        public Rectangle Region { get; }
    }
}
