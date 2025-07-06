using DocSharp.Renderer.Core;

namespace DocSharp.Renderer.Models.Common
{
    internal class VerticalSpace
    {
        public VerticalSpace(PageNumber pageNumber, double y, double height)
        {
            this.PageNumber = pageNumber;
            this.Y = y;
            this.Height = height;
            this.BottomY = y + height;
        }

        public PageNumber PageNumber { get; }
        public double Y { get; }
        public double Height { get; }
        public double BottomY { get; }
    }
}
