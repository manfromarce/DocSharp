using QuestPDF.Elements;
using QuestPDF.Fluent;
using QuestPDF.Infrastructure;

namespace DocSharp.Renderer;

internal class QuestPdfFootnote : QuestPdfContainer
{
    public int PageNumber { get; set; }
    public long Id { get; set; }

}
