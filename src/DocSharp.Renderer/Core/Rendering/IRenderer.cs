namespace DocSharp.Renderer.Core
{
    internal interface IRenderer
    {
        PdfRenderingOptions Options { get; }

        void CreatePage(PageNumber pageNumber, PageConfiguration configuration);

        IRendererPage GetPage(PageNumber pageNumber);

        IRendererPage GetPage(PageNumber pageNumber, Point offsetRendering);
    }
}
