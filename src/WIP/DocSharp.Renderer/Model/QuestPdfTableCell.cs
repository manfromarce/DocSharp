using QuestPDF.Infrastructure;

namespace DocSharp.Renderer;

internal class QuestPdfTableCell : QuestPdfContainer
{
    internal Color? BackgroundColor = null;

    internal uint ColumnSpan { get; set; } = 1;
    internal uint RowSpan { get; set; } = 1;
}
