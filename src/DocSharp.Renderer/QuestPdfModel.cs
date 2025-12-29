using System;
using System.Collections.Generic;
using System.Linq;
using QuestPDF.Fluent;
using QuestPDF.Helpers;
using QuestPDF.Infrastructure;

namespace DocSharp.Renderer;

internal class QuestPdfModel
{
    internal List<QuestPdfPageSet> PageSets = new();
    
    internal Document ToQuestPdfDocument()
    {
        return QuestPDF.Fluent.Document.Create(container => {
            foreach (var pageSet in PageSets)
            {
                // QuestPDF automatically creates a new page when content exceeds page size
                container.Page(page => {
                    page.Size(pageSet.PageWidth, pageSet.PageHeight, pageSet.Unit);   
                    page.MarginLeft(pageSet.MarginLeft, pageSet.Unit);            
                    page.MarginTop(pageSet.MarginTop, pageSet.Unit);            
                    page.MarginRight(pageSet.MarginRight, pageSet.Unit);            
                    page.MarginBottom(pageSet.MarginBottom, pageSet.Unit);
                    page.PageColor(pageSet.Color);

                    if (pageSet.Header != null && pageSet.Header.Content != null)
                    {
                        page.Header().Column(header =>
                        {
                            
                        });              
                    }
                    if (pageSet.Footer != null && pageSet.Footer.Content != null)
                    {
                        page.Footer().Column(footer =>
                        {
                            
                        });
                    }
                    if (pageSet.Content != null && pageSet.NumberOfColumns > 0 && pageSet.Content.Content != null)
                    {
                        page.Content().MultiColumn(content =>
                        {
                            content.Columns(pageSet.NumberOfColumns);
                            if (pageSet.NumberOfColumns > 1 && pageSet.SpaceBetweenColumns.HasValue)
                                content.Spacing(pageSet.SpaceBetweenColumns.Value, Unit.Point);
                            content.Content().Column(column =>
                            {
                                foreach (var element in pageSet.Content.Content)
                                {
                                    if (element is QuestPdfParagraph paragraph)
                                    {
                                        var paragraphItem = column.Item().PaddingLeft(paragraph.LeftIndent, Unit.Point)
                                                     .PaddingRight(paragraph.RightIndent, Unit.Point)
                                                     .PaddingTop(paragraph.SpaceBefore, Unit.Point)
                                                     .PaddingBottom(paragraph.SpaceAfter, Unit.Point);

                                        if (paragraph.KeepTogether)
                                        {
                                            paragraphItem = paragraphItem.PreventPageBreak();
                                        }
                                        if (paragraph.BackgroundColor.HasValue)
                                        {
                                            paragraphItem = paragraphItem.Background(paragraph.BackgroundColor.Value);
                                        }
                                        // TODO: paragraph borders
                                        
                                        paragraphItem.Text(text =>
                                        {
                                            // text.ParagraphSpacing(paragraph.SpaceAfter, Unit.Point);
                                            text.ParagraphFirstLineIndentation(paragraph.FirstLineIndent, Unit.Point);
                                            switch (paragraph.Alignment)
                                            {
                                                case ParagraphAlignment.Left: text.AlignLeft(); break;
                                                case ParagraphAlignment.Center: text.AlignCenter(); break;
                                                case ParagraphAlignment.Right: text.AlignRight(); break;
                                                case ParagraphAlignment.Start: text.AlignStart(); break;
                                                case ParagraphAlignment.End: text.AlignEnd(); break;
                                                case ParagraphAlignment.Justify: text.Justify(); break;
                                            }

                                            foreach (var span in paragraph.Spans)
                                            {
                                                // Why is LineHeight at the span level in QuestPDF rather than at the same level as ParagraphFirstLineIndentation?
                                                if (span.HyperlinkUrl != null)
                                                    text.Hyperlink(span.Text, span.HyperlinkUrl).Style(span.Style).LineHeight(paragraph.LineHeight);
                                                    // TODO: use SectionLink rather than Hyperlink if the hyperlink points to an internal bookmark
                                                else 
                                                    text.Span(span.Text).Style(span.Style).LineHeight(paragraph.LineHeight);                                                
                                            }
                                        });
                                    }
                                    else if (element is QuestPdfTable table)
                                    {
                                        column.Item().Table(t =>
                                        {
                                            
                                        });
                                    }
                                }
                            });
                        });                        
                    }
                });
            }
        });
    }
}

internal class QuestPdfPageSet(float pageWidth, float pageHeight, 
                               float marginLeft, float marginTop, float marginRight, float marginBottom,
                               Unit unit)
{
    internal float PageWidth{ get; set; } = pageWidth;
    internal float PageHeight { get; set; } = pageHeight;
    internal float MarginLeft { get; set; } = marginLeft;
    internal float MarginTop { get; set; } = marginTop;
    internal float MarginRight { get; set; } = marginRight;
    internal float MarginBottom { get; set; } = marginBottom;
    internal Unit Unit { get; set; } = unit;
    internal Color Color { get; set; } = Colors.White;

    internal int NumberOfColumns { get; set; } = 1;
    internal float? SpaceBetweenColumns { get; set; } // in points (if set)

    internal QuestPdfContainer? Header;
    internal QuestPdfContainer? Footer;
    internal QuestPdfContainer? Content;
}

internal class QuestPdfContainer
{
    internal List<QuestPdfBlock>? Content;
}

internal abstract class QuestPdfBlock
{
}

internal class QuestPdfParagraph : QuestPdfBlock
{
    internal List<QuestPdfSpan> Spans = new();

    internal List<QuestPdfSpan> Hyperlinks => Spans.Where(s => s.HyperlinkUrl != null).ToList();

    internal Color? BackgroundColor = null;

    internal ParagraphAlignment Alignment = ParagraphAlignment.Left;
    internal float LineHeight { get; set; } = 0; // relative factor (1.0 = standard)
    internal float SpaceBefore { get; set; } = 0; // points
    internal float SpaceAfter { get; set; } = 0; // points
    internal float LeftIndent { get; set; } = 0; // points
    internal float RightIndent { get; set; } = 0; // points
    internal float FirstLineIndent { get; set; } = 0; // points
    public bool KeepTogether { get; internal set; }
}

internal class QuestPdfTable : QuestPdfBlock
{
    internal List<QuestPdfTableRow> Rows = new();
}

internal class QuestPdfTableRow : QuestPdfBlock
{
    internal List<QuestPdfTableCell> Cells = new();
}

internal class QuestPdfTableCell : QuestPdfBlock
{
    internal List<QuestPdfBlock>? Content;
}

internal class QuestPdfSpan
{
    internal string Text { get; set; } = string.Empty;
    internal TextStyle Style { get; set; } = TextStyle.Default;

    internal string? HyperlinkUrl = null;

    internal QuestPdfSpan(string text, bool bold, bool italic, UnderlineStyle underline, StrikethroughStyle strikethrough, bool subscript, bool superscript, bool allCaps, bool smallCaps, string fontFamily, int? fontSize, Color? fontColor, Color? backgroundColor, Color? underlineColor, float? letterSpacing)
    {
        // TODO: span borders (not supported by QuestPDF)
        
        Text = text;
        if (bold)
            Style = Style.Bold();
        if (italic)
            Style = Style.Italic();
        
        if (subscript)
            Style = Style.Subscript();
        else if (superscript) // subscript and superscript are mutually exclusive
            Style = Style.Superscript();
       
        if (smallCaps)
            Style = Style.EnableFontFeature(FontFeatures.SmallCapitals); // Unclear if this works
        else if (smallCaps)
            Text = Text.ToUpper();

        // QuestPDF does not support independent styles for underline and strikethrough, decorations styles are applied to both. 
        // In addition, in Microsoft Word documents only solid single/double strikethrough with standard thickness and color are available.
        // For now, just ignore decoration styles if both underline and strikethrough are enabled.
        if (underline != UnderlineStyle.None)
        {
            Style = Style.Underline();
            if (strikethrough == StrikethroughStyle.None)
            {
                switch (underline)
                {
                    // Note that DOCX supports more underline styles, they need to be mapped to these in DocxRenderer 
                    // (for example LongDash and DashDot --> Dash; DoubleWavy --> Wavy).
                    case UnderlineStyle.Dashed: Style = Style.DecorationDashed(); break;
                    case UnderlineStyle.Dotted: Style = Style.DecorationDotted(); break;
                    case UnderlineStyle.Wavy: Style = Style.DecorationWavy(); break;
                    case UnderlineStyle.Double: Style = Style.DecorationDouble(); break;
                    // Otherwise stick to the default underline style (solid, single)
                }
                if (underlineColor.HasValue)
                    Style = Style.DecorationColor(underlineColor.Value);
            }   
        }
        if (strikethrough != StrikethroughStyle.None)
        {
            Style = Style.Strikethrough();
            if (underline == UnderlineStyle.None && strikethrough == StrikethroughStyle.Double)
            {
                Style = Style.DecorationDouble();
            }   
        }
        
        if (fontFamily != null)
            Style = Style.FontFamily([fontFamily]); 
            // TODO: add a fallback if font is not installed in the runtime environment; 
            // ship some royalty-free fonts with the library and register them using QuestPDF.Drawing.FontManager
        if (fontSize.HasValue)
            Style = Style.FontSize(fontSize.Value); // value in points
        if (letterSpacing.HasValue)
            Style = Style.LetterSpacing(letterSpacing.Value); // relative factor. 
            // The default value is 0, a negative value shrinks distance between letters, a positive value increases it.

        if (fontColor.HasValue)
            Style = Style.FontColor(fontColor.Value);
        if (backgroundColor.HasValue)
            Style = Style.BackgroundColor(backgroundColor.Value);
    }
}

internal enum ParagraphAlignment
{
    Left,
    Center,
    Right,
    Justify,
    Start,
    End
}

internal enum UnderlineStyle
{
    None,
    Solid,
    Dashed,
    Dotted,
    Double,
    Wavy
}

internal enum StrikethroughStyle
{
    None,
    Single,
    Double
}