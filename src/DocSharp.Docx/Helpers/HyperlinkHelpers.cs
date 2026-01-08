using System.Linq;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace DocSharp.Docx;

public static class HyperlinkHelpers
{
    public static string? GetUrl(this Hyperlink hyperlink)
    {
        if (hyperlink.Id?.Value is string rId)
        {
            var maindDocumentPart = OpenXmlHelpers.GetMainDocumentPart(hyperlink);
            if (maindDocumentPart?.HyperlinkRelationships.FirstOrDefault(x => x.Id == rId) is HyperlinkRelationship relationship)
            {
                return relationship.Uri.OriginalString;
            }
        }
        return null;
    }

    public static string? GetAnchor(this Hyperlink hyperlink)
    {
        if (hyperlink.Anchor?.Value is string anchor)
        {
            return anchor;
        }
        return null;
    }
}
