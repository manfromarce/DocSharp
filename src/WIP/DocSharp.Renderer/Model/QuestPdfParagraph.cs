using System;
using System.Collections.Generic;
using System.Linq;
using QuestPDF.Infrastructure;
using QuestPDF.Helpers;

namespace DocSharp.Renderer;
internal class QuestPdfParagraph : QuestPdfBlock, IQuestPdfRunContainer
{
    internal List<QuestPdfInlineElement> Elements = new();

    internal IEnumerable<QuestPdfSpan> Spans => Elements.OfType<QuestPdfSpan>();
    internal IEnumerable<QuestPdfHyperlink> Hyperlinks => Elements.OfType<QuestPdfHyperlink>();

    internal ParagraphAlignment Alignment = ParagraphAlignment.Left;
    internal float LineHeight { get; set; } = 0; // relative factor (1.0 = standard)
    internal float SpaceBefore { get; set; } = 0; // points
    internal float SpaceAfter { get; set; } = 0; // points
    internal float LeftIndent { get; set; } = 0; // points
    internal float RightIndent { get; set; } = 0; // points
    internal float StartIndent { get; set; } = 0; // points
    internal float EndIndent { get; set; } = 0; // points
    internal float FirstLineIndent { get; set; } = 0; // points
    internal Color? BackgroundColor = null;
    public bool KeepTogether { get; internal set; }
    public float LeftBorderThickness = 0.0f;
    public float RightBorderThickness = 0.0f;
    public float TopBorderThickness = 0.0f;
    public float BottomBorderThickness = 0.0f;
    public Color BordersColor = Colors.Black;

    internal  bool IsEmpty => Elements.Count(e => e is not QuestPdfBookmark) == 0;
    // include text, pictures, hyperlinks, footnotes references and page numbers in the count

    internal QuestPdfParagraph()
    {
        
    }

    internal QuestPdfParagraph(QuestPdfSpan span)
    {
        AddSpan(span);
    }
    
    internal QuestPdfParagraph(ParagraphAlignment alignment, float lineHeight, float spaceBefore, float spaceAfter, float leftIndent, float rightIndent, float startIndent, float endIndent, float firstLineIndent, Color? backgroundColor, bool keepTogether)
    {
        Alignment = alignment;
        LineHeight = lineHeight;
        SpaceBefore = spaceBefore;
        SpaceAfter = spaceAfter;
        LeftIndent = leftIndent;
        RightIndent = rightIndent;
        StartIndent = startIndent;
        EndIndent = endIndent;
        FirstLineIndent = firstLineIndent;
        BackgroundColor = backgroundColor;
        KeepTogether = keepTogether;
    }

    public void PrependSpan(QuestPdfSpan span)
    {
        Elements.Insert(0, span);
    }

    public void AddSpan(QuestPdfSpan span)
    {
        Elements.Add(span);
    }

    public void AddPageNumber()
    {
        Elements.Add(new QuestPdfPageNumber());        
    }

    public void AddFootnoteReference(long id)
    {
        Elements.Add(new QuestPdfFootnoteReference(id));        
    }

    public void AddBookmark(string name)
    {
        Elements.Add(new QuestPdfBookmark(name));
    }
    
    public IQuestPdfRunContainer CloneEmpty()
    {
        return new QuestPdfParagraph(Alignment, LineHeight, SpaceBefore, SpaceAfter, LeftIndent, RightIndent, StartIndent, EndIndent, FirstLineIndent, BackgroundColor, KeepTogether);
    }
}
