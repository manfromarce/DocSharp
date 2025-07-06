namespace DocSharp.Renderer.Core
{
    internal class PageConfiguration
    {
        public static readonly PageConfiguration Empty = new PageConfiguration(Size.Zero, PageOrientation.Portrait);

        public PageConfiguration(
            Size size,
            PageOrientation pageOrientation)
        {
            this.Size = size;
            this.PageOrientation = pageOrientation;
        }

        public Size Size { get; }

        public double Width => this.Size.Width;
        public double Height => this.Size.Height;

        public PageOrientation PageOrientation { get; }
    }
}
