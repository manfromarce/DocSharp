using System.Collections.Generic;

namespace DocSharp.Renderer;

internal class QuestPdfHyperlink : QuestPdfInlineElement, IQuestPdfRunContainer
{
    internal List<QuestPdfSpan> Spans = new();

    internal string? Url { get; set; }
    internal string? Anchor { get; set; }

    public void AddSpan(QuestPdfSpan span)
    {
        Spans.Add(span);
    }
}
