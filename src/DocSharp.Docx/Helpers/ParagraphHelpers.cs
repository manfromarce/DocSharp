using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using DocSharp.Helpers;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace DocSharp.Docx;

public static class ParagraphHelpers
{
    public static ParagraphProperties GetOrCreateProperties(this Paragraph p)
    {
        if (p.ParagraphProperties == null)
        {
            p.ParagraphProperties = new ParagraphProperties();
        }

        return p.ParagraphProperties;
    }

    public static void SetStyle(this Paragraph? para, string? styleId)
    {
        if (para == null || styleId == null) return;

        var pPr = para.ParagraphProperties;
        if (pPr == null)
        {
            pPr = new ParagraphProperties();
            para.ParagraphProperties = pPr;
        }

        var style = new ParagraphStyleId() { Val = styleId };
        pPr.ParagraphStyleId = style;
    }

    public static Paragraph AddParagraph(this MainDocumentPart mainDocumentPart, string text, string? styleId)
    {
        var p = CreateParagraph(text);

        mainDocumentPart.Document.Body ??= new Body();
        mainDocumentPart.Document.Body.AppendChild(p);

        if (!string.IsNullOrEmpty(styleId))
        {
            mainDocumentPart.Document.ApplyStyleToParagraph(styleId, p);
        }
        return p;
    }

    public static Paragraph CreateParagraph(string? text)
    {
        var para = new Paragraph();

        if (text == null) return para;

        var splits = text.NormalizeNewLines().Split("\n");

        var afterNewline = false;
        var run = new Run();
        foreach (var s in splits)
        {
            if (afterNewline)
            {
                var br = new Break();
                run.AppendChild(br);
            }

            Text t = new Text(s);

            if (s.StartsWith(" ") || s.EndsWith(" "))
            {
                t.Space = SpaceProcessingModeValues.Preserve;
            }

            run.AppendChild(t);				
            afterNewline = true;
        }

        para.AppendChild(run);
        return para;
    }

    public static void ApplyStyleToParagraph(this Document doc, string styleid, Paragraph p)
    {
        if (doc.MainDocumentPart != null)
        {
            // If the paragraph has no ParagraphProperties object, create one.
            var pPr = p.GetOrCreateProperties();

            // Get the Styles part for this document.
            var styles = doc.MainDocumentPart.GetOrCreateStylesPart();
            var style = styles.GetStyleFromId(styleid, StyleValues.Paragraph);
            if (style != null)
            {
                pPr.ParagraphStyleId = new ParagraphStyleId() { Val = styleid };
            }
        }        
    }

    public static Paragraph? FindParagraphContainingText(WordprocessingDocument document, string text)
    {
        if (document.MainDocumentPart == null || document.MainDocumentPart.Document.Body == null) return null;

        var textElement = document.MainDocumentPart.Document.Body
            .Descendants<Text>().FirstOrDefault(t => t.Text.Contains(text));

        if (textElement == null) return null;

        var p = textElement.Ancestors<Paragraph>().FirstOrDefault();
        return p;
    }
}
