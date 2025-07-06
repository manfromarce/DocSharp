using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace DocSharp.Renderer
{
    internal static class HeaderFooterXmlExtensions
    {
        public static Header FindHeader(
            this MainDocumentPart mainDocumentPart,
            string headerReferenceId)
        {
            if (headerReferenceId == null)
            {
                return null;
            }

            var headerPart = (HeaderPart)mainDocumentPart.GetPartById(headerReferenceId);
            return headerPart.Header;
        }

        public static Footer FindFooter(
            this MainDocumentPart mainDocumentPart,
            string footerReferenceId)
        {
            if (footerReferenceId == null)
            {
                return null;
            }

            var footerPart = (FooterPart)mainDocumentPart.GetPartById(footerReferenceId);
            return footerPart.Footer;
        }
    }
}
