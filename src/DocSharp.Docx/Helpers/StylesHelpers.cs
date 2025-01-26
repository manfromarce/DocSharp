using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace DocSharp.Docx;

public static class StylesHelpers
{
    public static Styles GetOrCreateStylesPart(this MainDocumentPart mainDocumentPart)
    {
        var part = mainDocumentPart.StyleDefinitionsPart;
        if (part == null)
        {
            part = mainDocumentPart.AddNewPart<StyleDefinitionsPart>();
        }
        var styles = part.Styles;
        if (styles == null)
        {
            styles = new Styles();
        }
        return styles;
    } 

    public static Style? GetStyleFromId(this Styles? stylesPart, string? id, StyleValues styleType)
    {
        if (string.IsNullOrEmpty(id))
            return null;

        return stylesPart?.Elements<Style>().FirstOrDefault(s => s.StyleId != null && 
                                                                 s.StyleId == id &&
                                                                 s.Type != null &&
                                                                 s.Type == styleType);
    }

    public static Style? GetStyleFromName(this Styles? stylesPart, string? name, StyleValues styleType)
    {
        if (string.IsNullOrEmpty(name))
            return null;

        return stylesPart?.Elements<Style>().FirstOrDefault(s => s.StyleName != null && 
                                                                 s.StyleName.Val == name && 
                                                                 s.Type != null && 
                                                                 s.Type == styleType);
    }

    public static Style? GetBaseStyle(this Styles? stylesPart, Style? childStyle)
    {
        if (stylesPart is null || childStyle is null || childStyle.BasedOn is null || childStyle.Type is null)
            return null;

        return stylesPart?.GetStyleFromId(childStyle.BasedOn.Val, childStyle.Type);
    }

    public static RunPropertiesBaseStyle? GetDefaultRunStyle(this Styles? stylesPart)
    {
        return stylesPart?.DocDefaults?.RunPropertiesDefault?.RunPropertiesBaseStyle;
    }

    public static ParagraphPropertiesBaseStyle? GetDefaultParagraphStyle(this Styles? stylesPart)
    {
        return stylesPart?.DocDefaults?.ParagraphPropertiesDefault?.ParagraphPropertiesBaseStyle;
    }

    // Return true if the style id is in the document, false otherwise.
    public static bool ContainsStyle(this Styles styles, string styleid, StyleValues styleType)
    {        
        return styles.GetStyleFromId(styleid, styleType) != null;
    }

    // Return styleid that matches the styleName, or null when there's no match.
    public static string? GetStyleIdFromStyleName(Document doc, string styleName, StyleValues styleType)
    {
        var stylePart = doc.MainDocumentPart?.StyleDefinitionsPart;
        var styleId = stylePart?.Styles?.Descendants<StyleName>()
            .Where(s => styleName.Equals(s.Val?.Value)
                        && s.Parent is Style parent
                        && parent.Type != null
                        && parent.Type == styleType)
            .Select(n => n.Parent as Style)
            .Where(n => n != null)
            .Select(n => n!.StyleId).FirstOrDefault();

        return styleId;
    }
}
