using System.Linq;
using DocumentFormat.OpenXml.Packaging;
using DocSharp.Renderer.Core;
using DocSharp.Renderer.Models.Common;
using DocSharp.Renderer.Models.Sections;
using DocSharp.Renderer.Models.Sections.Builders;
using DocSharp.Renderer.Models.Styles;

namespace DocSharp.Renderer.Models
{
    internal class Document
    {
        private Section[] _sections = new Section[0];
        private readonly WordprocessingDocument _docx;
        private readonly IStyleFactory _styleAccessor;

        public Document(WordprocessingDocument docx)
        {
            _docx = docx;
            _styleAccessor = StyleFactory.Default(docx.MainDocumentPart);
        }

        public void Render(IRenderer renderer)
        {
            this.InitializeSections();

            this.PrepareSections();

            this.RenderSections(renderer);
        }

        private void InitializeSections()
        {
            _sections = _docx.MainDocumentPart
                .SplitToSections(_styleAccessor)
                .ToArray();
        }

        private void PrepareSections()
        {
            bool isFinished;
            var lastPageNumber = PageNumber.None;

            do
            {
                var previousSection = PageRegion.None;
                var previousSectionMargin = PageMargin.PageNone;

                foreach (var section in _sections)
                {
                    section.Prepare(previousSection, previousSectionMargin, new DocumentVariables(lastPageNumber));
                    previousSection = section.PageRegions.Last();
                    previousSectionMargin = section.Pages.Last().Margin;
                }

                var secionLastPage = _sections.Last()
                    .Pages
                    .Last();

                isFinished = lastPageNumber == secionLastPage.PageNumber;
                lastPageNumber = secionLastPage.PageNumber;
            } while (!isFinished);
        }

        private void RenderSections(IRenderer renderer)
        {
            foreach(var section in _sections)
            {
                foreach(var page in section.Pages)
                {
                    renderer.CreatePage(page.PageNumber, page.Configuration);
                }

                section.Render(renderer);
            }
        }
    }
}
