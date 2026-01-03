using System.Collections.Generic;

namespace DocSharp.Renderer;

internal class QuestPdfTable : QuestPdfBlock
{
    internal List<QuestPdfTableRow> Rows = new();

    internal int ColumnsCount { get; set; } = 0;
}
