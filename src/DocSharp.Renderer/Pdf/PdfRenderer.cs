using System.Collections.Generic;
using PeachPDF.PdfSharpCore.Drawing;
using PeachPDF.PdfSharpCore.Pdf;
using DocSharp.Renderer.Core;

namespace DocSharp.Renderer.Pdf
{
    internal class PdfRenderer : IRenderer
    {
        private readonly PdfDocument _pdfDocument;

        private Dictionary<PageNumber, PdfRendererPage> _pages = new Dictionary<PageNumber, PdfRendererPage>();

        public PdfRenderer(PdfDocument pdfDocument, PdfRenderingOptions renderingOptions)
        {
            _pdfDocument = pdfDocument;
            this.Options = renderingOptions;
        }

        public PdfRenderingOptions Options { get; }

        public void CreatePage(PageNumber pageNumber, PageConfiguration configuration)
        {
            if (_pages.ContainsKey(pageNumber))
            {
                return;
            }

            var pdfPage = new PdfPage
            {
                Orientation = (PeachPDF.PdfSharpCore.PageOrientation)configuration.PageOrientation
            };

            pdfPage.Width = configuration.Size.Width;
            pdfPage.Height = configuration.Size.Height;

            _pdfDocument.AddPage(pdfPage);
            _pages.Add(pageNumber, new PdfRendererPage(pageNumber, XGraphics.FromPdfPage(pdfPage), this.Options));
        }

        public IRendererPage GetPage(PageNumber pageNumber)
            => this.GetPage(pageNumber, Point.Zero);
        
        public IRendererPage GetPage(PageNumber pageNumber, Point offsetRendering)
        {
            var page = _pages[pageNumber];
            return offsetRendering == Point.Zero
                ? page
                : page.Offset(offsetRendering);
        }
    }
}
