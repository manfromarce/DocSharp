using System.Collections.Generic;
using System.Linq;
using DocSharp.Renderer.Core;
using DocSharp.Renderer.Models.Common;
using DocSharp.Renderer.Models.Sections.Columns;
using DocSharp.Renderer.Models.Styles;

using OpenXml = DocumentFormat.OpenXml;
using Pack = DocumentFormat.OpenXml.Packaging;
using Word = DocumentFormat.OpenXml.Wordprocessing;

namespace DocSharp.Renderer.Models.Sections.Builders
{
    internal static class SectionBuilder
    {
        public static IEnumerable<Section> SplitToSections(
            this Pack.MainDocumentPart mainDocument,
            IStyleFactory styleFactory)
        {
            var sections = mainDocument.Document.Body
                .SplitToSectionsCore(mainDocument, styleFactory)
                .ToArray();

            return sections;
        }

        private static IEnumerable<Section> SplitToSectionsCore(
            this Word.Body body,
            Pack.MainDocumentPart mainDocumentPart,
            IStyleFactory styleFactory)
        {
            var sections = new List<Section>();

            var sectionElements = new List<OpenXml.OpenXmlCompositeElement>();
            var headerFooterConfiguration = HeaderFooterConfiguration.Empty;
            Word.SectionProperties wordSectionProperties;

            foreach (var e in body.RenderableChildren())
            {
                sectionElements.Add(e);
                if (!(e is Word.Paragraph paragraph))
                {
                    continue;
                }

                wordSectionProperties = paragraph.GetSectionProperties();
                if (wordSectionProperties == null)
                {
                    continue;
                }

                var section = sectionElements.CreateSection(wordSectionProperties, mainDocumentPart, headerFooterConfiguration, sections.Count == 0, styleFactory);
                headerFooterConfiguration = section.HeaderFooterConfiguration;
                sections.Add(section);
                sectionElements.Clear();
            }

            wordSectionProperties = body
               .ChildsOfType<Word.SectionProperties>()
               .Single();

            var lastSection = sectionElements.CreateSection(wordSectionProperties, mainDocumentPart, headerFooterConfiguration, sections.Count == 0, styleFactory);
            sections.Add(lastSection);
            return sections;
        }

        private static Section CreateSection(
            this IReadOnlyCollection<OpenXml.OpenXmlCompositeElement> xmlElements,
            Word.SectionProperties wordSectionProperties,
            Pack.MainDocumentPart mainDocumentPart,
            HeaderFooterConfiguration headerFooterConfiguration,
            bool isFirst,
            IStyleFactory styleFactory)
        {
            var imageAccessor = new ImageAccessor(mainDocumentPart);

            var sectionProperties = wordSectionProperties.CreateSectionProperties(mainDocumentPart, isFirst, headerFooterConfiguration);
            var columnsConfiguration = wordSectionProperties.CreateColumnsConfiguration(sectionProperties.PageConfiguration, sectionProperties.Margin);
            var sectionContents = xmlElements.SplitToSectionContents(columnsConfiguration, imageAccessor, styleFactory);
            var sd = new Section(sectionProperties, sectionContents, imageAccessor, styleFactory);

            return sd;
        }

        private static IEnumerable<SectionContent> SplitToSectionContents(
            this IEnumerable<OpenXml.OpenXmlCompositeElement> xmlElements,
            ColumnsConfiguration columnsConfiguration,
            IImageAccessor imageAccessor,
            IStyleFactory styleFactory)
        {
            var sectionContents = new List<SectionContent>();

            var stack = xmlElements.ToStack();
            var contentElements = new List<OpenXml.OpenXmlCompositeElement>();

            while (stack.Count > 0)
            {
                var e = stack.Pop();
                switch (e)
                {
                    case Word.Paragraph p:
                        {
                            var (begin, @break, end) = p.SplitByNextBreak();
                            if (@break == SectionContentBreak.None)
                            {
                                contentElements.Add(p);
                            }
                            else
                            {
                                if (end != null)
                                {
                                    stack.Push(end);
                                }

                                contentElements.Add(begin);
                                var childElements = contentElements.CreatePageElements(imageAccessor, styleFactory);
                                sectionContents.Add(new SectionContent(childElements, columnsConfiguration, @break));
                                contentElements.Clear();
                            }
                        }
                        break;
                    default:
                        contentElements.Add(e);
                        break;
                }
            }

            if (contentElements.Count > 0)
            {
                var childElements = contentElements.CreatePageElements(imageAccessor, styleFactory);
                sectionContents.Add(new SectionContent(childElements, columnsConfiguration, SectionContentBreak.None));
            }

            return sectionContents;
        }

        private static Word.SectionProperties GetSectionProperties(this Word.Paragraph paragraph)
        {
            return paragraph.ParagraphProperties?.SectionProperties;
        }

        private static SectionProperties CreateSectionProperties(
            this Word.SectionProperties wordSectionProperties,
            Pack.MainDocumentPart mainDocument,
            bool isFirstSection,
            HeaderFooterConfiguration inheritHeaderFooterConfiguration)
        {
            var pageCongifuration = wordSectionProperties.GetPageConfiguration();
            var pageMargin = wordSectionProperties.GetPageMargin();

            var sectionMark = wordSectionProperties.ChildsOfType<Word.SectionType>().SingleOrDefault()?.Val ?? Word.SectionMarkValues.NextPage;

            var requiresNewPage = isFirstSection || sectionMark == Word.SectionMarkValues.NextPage;

            var headerFooterConfiguration = wordSectionProperties
                .GetHeaderFooterConfiguration(mainDocument, inheritHeaderFooterConfiguration);

            return new SectionProperties(
                pageCongifuration,
                headerFooterConfiguration,
                pageMargin,
                requiresNewPage);
        }

        private static PageConfiguration GetPageConfiguration(
            this Word.SectionProperties sectionProperties)
        {
            var pageSize = sectionProperties.ChildsOfType<Word.PageSize>().Single();
            var w = pageSize.Width.DxaToPoint();
            var h = pageSize.Height.DxaToPoint();


            var orientation = (pageSize.Orient?.Value ?? Word.PageOrientationValues.Portrait) == Word.PageOrientationValues.Portrait
                ? PageOrientation.Portrait
                : PageOrientation.Landscape;

            return new PageConfiguration(new Size(w, h), orientation);
        }

        private static PageMargin GetPageMargin(this Word.SectionProperties sectionProperties)
        {
            var pageMargin = sectionProperties.ChildsOfType<Word.PageMargin>().Single();
            return new PageMargin(
                pageMargin.Top.DxaToPoint(),
                pageMargin.Right.DxaToPoint(),
                pageMargin.Bottom.DxaToPoint(),
                pageMargin.Left.DxaToPoint(),
                pageMargin.Header.DxaToPoint(),
                pageMargin.Footer.DxaToPoint());
        }

        private static HeaderFooterConfiguration GetHeaderFooterConfiguration(
            this Word.SectionProperties wordSectionProperties,
            Pack.MainDocumentPart mainDocument,
            HeaderFooterConfiguration previousHeaderFooterConfiguration)
        {
            var hasTitlePage = wordSectionProperties.ChildsOfType<Word.TitlePage>().SingleOrDefault()
                  .IsOn(ifOnOffTypeNull: false, ifOnOffValueNull: true);

            var headerRefs = wordSectionProperties
                .ChildsOfType<Word.HeaderReference>()
                .Select(fr => new HeaderFooterRef(fr.Id, fr.Type));

            var footerRefs = wordSectionProperties
                .ChildsOfType<Word.FooterReference>()
                .Select(fr => new HeaderFooterRef(fr.Id, fr.Type));

            return previousHeaderFooterConfiguration.Inherited(mainDocument, hasTitlePage, headerRefs, footerRefs);
        }

        private static (Word.Paragraph, SectionContentBreak, Word.Paragraph) SplitByNextBreak(this Word.Paragraph paragraph)
        {
            var index = paragraph
                .ChildElements
                .IndexOf(e => e is Word.Run r && r.ChildsOfType<Word.Break>().Any(b => b.Type != null && (b.Type == Word.BreakValues.Page || b.Type == Word.BreakValues.Column)));

            if (index == -1)
            {
                return (paragraph, SectionContentBreak.None, null);
            }

            var beginElements = paragraph
                .ChildElements
                .Take(index + 1)
                .Select(r => r.CloneNode(true))
                .ToArray();

            var begin = new Word.Paragraph(beginElements)
            {
                ParagraphProperties = (Word.ParagraphProperties)paragraph.ParagraphProperties.CloneNode(true)
            };

            var endElements = paragraph
                .ChildElements
                .Skip(index + 1)
                .Select(r => r.CloneNode(true))
                .ToArray();

            var end = endElements.Length == 0
                ? null
                : new Word.Paragraph(endElements)
                {
                    ParagraphProperties = (Word.ParagraphProperties)paragraph.ParagraphProperties.CloneNode(true)
                };

            var breakRun = (Word.Run)paragraph
                .ChildElements
                .ElementAt(index);

            var breakType = breakRun
                .ChildElements
                .OfType<Word.Break>()
                .Single()
                .Type;


            var @break = SectionContentBreak.None;
            if (breakType != null && breakType.HasValue && breakType.Value == Word.BreakValues.Column)
            {
                @break = SectionContentBreak.Column;
            }
            else if (breakType != null && breakType.HasValue && breakType.Value == Word.BreakValues.Page)
            {
                @break = SectionContentBreak.Page;
            }

            return (begin, @break, end);
        }
    }
}
