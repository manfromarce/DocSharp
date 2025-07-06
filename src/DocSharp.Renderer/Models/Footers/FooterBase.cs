using DocSharp.Renderer.Models.Common;

namespace DocSharp.Renderer.Models.Footers
{
    internal abstract class FooterBase : PageElement
    {
        protected FooterBase(PageMargin pageMargin)
        {
            this.PageMargin = pageMargin;
        }

        public double Height => this.LastPageRegion.Region.Height;
        public double FooterMargin => this.PageMargin.Footer;
        public double HeightWithFooterMargin => this.Height + this.PageMargin.Footer;

        protected PageMargin PageMargin { get; }

        public abstract void Prepare(IPage page);
    }
}
