using DocSharp.Renderer.Core;
using DocSharp.Renderer.Models.Common;

namespace DocSharp.Renderer.Models
{
    internal class PageContext
    {
        public PageContext(
            PagePosition pagePosition,
            Rectangle region,
            DocumentVariables variables)
        {
            this.Region = region;
            this.PagePosition = pagePosition;
            this.PageVariables = new PageVariables(pagePosition.PageNumber, variables.TotalPages);
            this.TopLeft = new DocumentPosition(pagePosition, this.Region.TopLeft);
        }

        public Rectangle Region { get; }
        public PageVariables PageVariables { get; }

        public DocumentPosition TopLeft { get; }
        public PagePosition PagePosition { get; }

        public PageContext Crop(Margin margin)
            => this.Crop(margin.Top, margin.Right, margin.Bottom, margin.Left);

        public PageContext Crop(HorizontalSpace space)
            => this.Crop(0, this.Region.Width - space.X - space.Width, 0, space.X);

        public PageContext CropFromTop(double top)
            => this.Crop(top, 0, 0, 0);

        public PageContext Crop(double top, double right, double bottom, double left)
        {
            var region = this.Region.Crop(top, right, bottom, left);
            return this.WithRegion(region);
        }

        private PageContext WithRegion(Rectangle region)
        {
            return new PageContext(this.PagePosition, region, this.PageVariables);
        }
    }
}
