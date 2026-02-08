using System.Linq;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace DocSharp.Docx;

public static class HeaderFooterHelpers
{
    public static Header? GetHeaderFromReference(HeaderReference? headerReference, MainDocumentPart? mainPart)
    {
        if (headerReference != null && mainPart != null && headerReference?.Id?.Value is string headerId &&
            !string.IsNullOrWhiteSpace(headerId) &&
            mainPart.TryGetPartById(headerId, out OpenXmlPart? part) && 
            part is HeaderPart headerPart)
            return headerPart.Header;
        else 
            return null;
    }

    public static Footer? GetFooterFromReference(FooterReference? footerReference, MainDocumentPart? mainPart)
    {
        if (footerReference != null && mainPart != null && footerReference?.Id?.Value is string footerId && 
            !string.IsNullOrWhiteSpace(footerId) && 
            mainPart.TryGetPartById(footerId, out OpenXmlPart? part) && 
            part is FooterPart footerPart)
            return footerPart.Footer;
        else 
            return null;
    }
}
