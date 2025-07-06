using System.Collections.Generic;
using System.Linq;
using DocumentFormat.OpenXml.Packaging;
using DocSharp.Renderer.Core;

using Word = DocumentFormat.OpenXml.Wordprocessing;

namespace DocSharp.Renderer.Models.Sections
{
    internal class HeaderFooterConfiguration
    {
        public static readonly HeaderFooterConfiguration Empty = new HeaderFooterConfiguration(null, false, new HeaderFooterRef[0], new HeaderFooterRef[0]);

        private readonly HeaderFooterRef[] _headers;
        private readonly HeaderFooterRef[] _footers;
        private readonly MainDocumentPart _mainDocument;
        private readonly bool _hasTitlePage;

        private HeaderFooterConfiguration(
            MainDocumentPart mainDocument,
            bool hasTitlePage,
            IEnumerable<HeaderFooterRef> headers,
            IEnumerable<HeaderFooterRef> footers /*, int pageNumberStart*/)
        {
            _mainDocument = mainDocument;
            _hasTitlePage = hasTitlePage;
            _headers = headers.ToArray();
            _footers = footers.ToArray();
        }

        private bool UseEvenOddHeadersAndFooters => _mainDocument?.DocumentSettingsPart.EvenOddHeadersAndFooters() ?? false;

        public Word.Header FindHeader(PageNumber pageNumber)
        {
            var referenceId = this.GetHeaderReferenceId(pageNumber);
            var header = _mainDocument?.FindHeader(referenceId);
            return header;
        }

        public Word.Footer FindFooter(PageNumber pageNumber)
        {
            var referenceId = this.GetFooterReferenceId(pageNumber);
            var footer = _mainDocument?.FindFooter(referenceId);
            return footer;
        }

        public HeaderFooterConfiguration Inherited(
            MainDocumentPart mainDocument,
            bool hasTitlePage,
            IEnumerable<HeaderFooterRef> headers,
            IEnumerable<HeaderFooterRef> footers)
        {
            var h = headers.ToArray();
            if (h.Length == 0)
            {
                h = _headers.ToArray();
            }

            var f = footers.ToArray();
            if (f.Length == 0)
            {
                f = _footers.ToArray();
            }

            return new HeaderFooterConfiguration(mainDocument, hasTitlePage, h, f);
        }

        private string GetHeaderReferenceId(PageNumber pageNumber)
        {
            return this.GetReferenceId(_headers, pageNumber);
        }

        private string GetFooterReferenceId(PageNumber pageNumber)
        {
            return this.GetReferenceId(_footers, pageNumber);
        }

        private string GetReferenceId(IEnumerable<HeaderFooterRef> references, int pageNumber)
        {
            if (_hasTitlePage && pageNumber == 1)
            {
                return references.FirstOrDefault(r => r.Type == Word.HeaderFooterValues.First)?.Id
                    ?? references.FirstOrDefault(r => r.Type == Word.HeaderFooterValues.Default)?.Id;
            }

            if (!this.UseEvenOddHeadersAndFooters || pageNumber % 2 == 1)
            {
                return references.FirstOrDefault(r => r.Type == Word.HeaderFooterValues.Default)?.Id;
            }

            return references.FirstOrDefault(r => r.Type == Word.HeaderFooterValues.Even)?.Id
                ?? references.FirstOrDefault(r => r.Type == Word.HeaderFooterValues.Default)?.Id;
        }
    }
}
