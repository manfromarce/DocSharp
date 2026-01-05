using System.Collections.Generic;
using System.Linq;

namespace DocSharp.Renderer;

internal class QuestPdfHyperlink : QuestPdfInlineElement, IQuestPdfRunContainer
{
    internal List<QuestPdfInlineElement> Elements = new();

    internal QuestPdfHyperlink()
    {
        
    }
    
    internal QuestPdfHyperlink(string? url, string? anchor)
    {
        Url = url;
        Anchor = anchor;
    }

    internal IEnumerable<QuestPdfSpan> Spans => Elements.OfType<QuestPdfSpan>();

    internal string? Url { get; set; }
    internal string? Anchor { get; set; }

    public void AddSpan(QuestPdfSpan span)
    {
        Elements.Add(span);
    }

    public IQuestPdfRunContainer CloneEmpty()
    {
        return new QuestPdfHyperlink(Url, Anchor);
    }
}
