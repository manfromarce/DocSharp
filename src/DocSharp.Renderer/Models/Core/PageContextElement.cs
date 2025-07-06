using System;
using System.Collections.Generic;
using System.Linq;
using DocSharp.Renderer.Core;
using DocSharp.Renderer.Models.Common;

namespace DocSharp.Renderer.Models
{
    internal abstract class PageContextElement : PageElement
    {
        public abstract void SetPageOffset(Point pageOffset);

        public abstract void Prepare(
            PageContext pageContext,
            Func<PagePosition, PageContextElement, PageContext> nextPageContextFactory);
    }
}
