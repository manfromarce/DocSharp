using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace DocSharp.Docx;

public static class StylesHelpers
{
    public static bool IsSameStyle(Paragraph par1, Paragraph par2)
    {
        var styleId1 = par1.ParagraphProperties?.ParagraphStyleId?.Val?.Value;
        var styleId2 = par1.ParagraphProperties?.ParagraphStyleId?.Val?.Value;
        if (string.IsNullOrEmpty(styleId1) && string.IsNullOrEmpty(styleId2))
            return true;
        
        if (string.IsNullOrEmpty(styleId1) || string.IsNullOrEmpty(styleId2))
            return false;
        else 
            return styleId1!.Equals(styleId2, StringComparison.OrdinalIgnoreCase);
    }

    public static Styles GetOrCreateStylesPart(this WordprocessingDocument document)
    {
        var mainPart = document.MainDocumentPart ?? document.AddMainDocumentPart();
        return mainPart.GetOrCreateStylesPart();
    }

    public static Styles GetOrCreateStylesPart(this MainDocumentPart mainPart)
    {
        var stylesPart = mainPart.StyleDefinitionsPart ?? mainPart.AddNewPart<StyleDefinitionsPart>();
        stylesPart.Styles ??= new Styles();
        return stylesPart.Styles;
    }

    /// <summary>
    /// Helper function to retrieve styles part from an Open XML element.
    /// </summary>
    /// <returns></returns>
    public static Styles? GetStylesPart(this OpenXmlElement element)
    {
        return element.GetMainDocumentPart()?.StyleDefinitionsPart?.Styles;
    }

    public static Style? GetStyleFromId(this Styles? stylesPart, string? id, StyleValues styleType)
    {
        if (string.IsNullOrEmpty(id))
            return null;

        return stylesPart?.FirstOrDefault<Style>(s => s.StyleId != null && 
                                                      s.StyleId == id &&
                                                      s.Type != null &&
                                                      s.Type == styleType);
    }

    public static Style? GetStyleFromName(this Styles? stylesPart, string? name, StyleValues styleType)
    {
        if (string.IsNullOrEmpty(name))
            return null;

        return stylesPart?.FirstOrDefault<Style>(s => s.StyleName != null && 
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
    public static bool ContainsStyleId(this Styles styles, string styleId, StyleValues styleType)
    {
        return styles.GetStyleFromId(styleId, styleType) != null;
    }

    // Return true if the style name is in the document, false otherwise.
    public static bool ContainsStyleName(this Styles styles, string styleName, StyleValues styleType)
    {
        return styles.GetStyleFromName(styleName, styleType) != null;
    }

    // Return styleId name that matches the specified style name, or null if there is no match.
    public static string? GetStyleIdFromStyleName(Document doc, string styleName, StyleValues styleType)
    {
        var styles = doc.MainDocumentPart?.StyleDefinitionsPart?.Styles;
        if (styles == null) return null;
        return GetStyleFromName(styles, styleName, styleType)?.StyleId?.Value;
    }

    // Return style name that matches the specified style id, or null if there is no match.
    public static string? GetStyleNameFromStyleId(Document doc, string styleId, StyleValues styleType)
    {
        var styles = doc.MainDocumentPart?.StyleDefinitionsPart?.Styles;
        if (styles == null) return null;
        return GetStyleFromId(styles, styleId, styleType)?.StyleName?.Val?.Value;
    }

        // Return styleId name that matches the specified style name, or null if there is no match.
    public static string? GetStyleIdFromStyleName(Styles styles, string styleName, StyleValues styleType)
    {
        return GetStyleFromName(styles, styleName, styleType)?.StyleId?.Value;
    }

    // Return style name that matches the specified style id, or null if there is no match.
    public static string? GetStyleNameFromStyleId(Styles styles, string styleId, StyleValues styleType)
    {
        return GetStyleFromId(styles, styleId, styleType)?.StyleName?.Val?.Value;
    }

    public static string? GetStyleId(this Run run)
    {
        return run?.RunProperties?.RunStyle?.Val?.Value;
    }

    public static string? GetStyleName(this Run run)
    {
        var styles = run.GetStylesPart();
        var styleId = run.GetStyleId();
        if (styles == null || styleId == null) return null;
        return GetStyleNameFromStyleId(styles, styleId, StyleValues.Character);
    }

    public static string? GetStyleId(this Paragraph paragraph)
    {
        return paragraph?.ParagraphProperties?.ParagraphStyleId?.Val?.Value;
    }

    public static string? GetStyleName(this Paragraph paragraph)
    {
        var styles = paragraph.GetStylesPart();
        var styleId = paragraph.GetStyleId();
        if (styles == null || styleId == null) return null;
        return GetStyleNameFromStyleId(styles, styleId, StyleValues.Paragraph);
    }
}
