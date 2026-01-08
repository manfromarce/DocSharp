using System.Collections.Generic;
using System.Linq;
using QuestPDF.Infrastructure;

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

    internal QuestPdfParagraph()
    {
        
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

    public IQuestPdfRunContainer CloneEmpty()
    {
        return new QuestPdfParagraph(Alignment, LineHeight, SpaceBefore, SpaceAfter, LeftIndent, RightIndent, StartIndent, EndIndent, FirstLineIndent, BackgroundColor, KeepTogether);
    }
}
