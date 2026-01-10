using System.Collections.Generic;

namespace DocSharp.Renderer;

internal class QuestPdfTable : QuestPdfBlock
{
    public List<QuestPdfTableRow> Rows = new();

    public List<float> ColumnsWidth { get; set; } = new();

    public HorizontalAlignment Alignment { get; set; } = HorizontalAlignment.Left;
}
