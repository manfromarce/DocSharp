using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace DocSharp.Docx;

public static class HeaderFooterHelpers
{
    public static Header? GetHeaderFromReference(HeaderReference? headerReference, MainDocumentPart? mainPart)
    {
        if (headerReference != null && mainPart != null && headerReference?.Id?.Value is string headerId)
            return (mainPart.GetPartById(headerId) as HeaderPart)?.Header;
        else 
            return null;
    }

    public static Footer? GetFooterFromReference(FooterReference? footerReference, MainDocumentPart? mainPart)
    {
        if (footerReference != null && mainPart != null && footerReference?.Id?.Value is string footerId)
            return (mainPart.GetPartById(footerId) as FooterPart)?.Footer;
        else 
            return null;
    }
}
