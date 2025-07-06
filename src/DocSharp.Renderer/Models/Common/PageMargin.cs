namespace DocSharp.Renderer.Models.Common
{
    internal class PageMargin : Margin
    {
        public static readonly PageMargin PageNone = new PageMargin(0, 0, 0, 0, 0, 0);

        public PageMargin(double top, double right, double bottom, double left, double header, double footer) : base(top, right, bottom, left)
        {
            this.Header = header;
            this.Footer = footer;
        }

        public double Header { get; }
        public double Footer { get; }

        public double MinimalHeaderHeight => this.Top - this.Header;
        public double FooterHeight => this.Bottom - this.Footer;

        public PageMargin WithHorizontal(double left, double right)
            => new PageMargin(this.Top, right, this.Bottom, left, this.Header, this.Footer);

        public PageMargin WithTop(double header, double top)
            => new PageMargin(top, this.Right, this.Bottom, this.Left, header, this.Footer);

        public PageMargin WithBottom(double footer, double bottom)
            => new PageMargin(this.Top, this.Right, bottom, this.Left, this.Header, footer);
    }
}
