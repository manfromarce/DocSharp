using System.Collections.Generic;
using System.Linq;
using QuestPDF.Infrastructure;

namespace DocSharp.Renderer;
internal class QuestPdfParagraph : QuestPdfBlock, IQuestPdfRunContainer
{
    internal List<QuestPdfInlineElement> Elements = new();

    internal IEnumerable<QuestPdfSpan> Spans => Elements.OfType<QuestPdfSpan>();
    internal IEnumerable<QuestPdfHyperlink> Hyperlinks => Elements.OfType<QuestPdfHyperlink>();

    internal Color? BackgroundColor = null;

    internal ParagraphAlignment Alignment = ParagraphAlignment.Left;
    internal float LineHeight { get; set; } = 0; // relative factor (1.0 = standard)
    internal float SpaceBefore { get; set; } = 0; // points
    internal float SpaceAfter { get; set; } = 0; // points
    internal float LeftIndent { get; set; } = 0; // points
    internal float RightIndent { get; set; } = 0; // points
    internal float FirstLineIndent { get; set; } = 0; // points
    public bool KeepTogether { get; internal set; }

    public void AddSpan(QuestPdfSpan span)
    {
        Elements.Add(span);
    }
}
