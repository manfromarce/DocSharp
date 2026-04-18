using System;
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
            if (hyperlink.GetRootPart()?.HyperlinkRelationships.FirstOrDefault(x => x.Id == rId) is HyperlinkRelationship relationship)
            {
                return relationship.Uri.OriginalString;
            }
        }
        return null;
    }

    public static void SetUrl(this Hyperlink hyperlink, string url)
    {
        if (hyperlink.GetRootPart() is OpenXmlPart part)
        {
            if (hyperlink.Id?.Value is string rId && part.HyperlinkRelationships.FirstOrDefault(x => x.Id == rId) is HyperlinkRelationship relationship)
                part.DeleteReferenceRelationship(relationship);

            var rel = part.AddHyperlinkRelationship(new Uri(url), true);
            hyperlink.Anchor = null;
            hyperlink.Id = rel.Id;
        }
    }

    public static string? GetAnchor(this Hyperlink hyperlink)
    {
        if (hyperlink.Anchor?.Value is string anchor)
        {
            return anchor;
        }
        return null;
    }

    public static void SetAnchor(this Hyperlink hyperlink, string anchor)
    {
        if (hyperlink.GetRootPart() is OpenXmlPart part)
        {
            if (hyperlink.Id?.Value is string rId && part.HyperlinkRelationships.FirstOrDefault(x => x.Id == rId) is HyperlinkRelationship relationship)
                part.DeleteReferenceRelationship(relationship);

            hyperlink.Anchor = anchor;
            hyperlink.Id = null;
        }
    }
}
