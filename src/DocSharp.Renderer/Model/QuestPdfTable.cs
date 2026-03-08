using System;
using System.Collections.Generic;
using System.Linq;

namespace DocSharp.Renderer;

internal class QuestPdfTable : QuestPdfBlock
{
    public List<QuestPdfTableRow> Rows = new();

    public List<float> ColumnsWidth { get; set; } = new();

    public HorizontalAlignment Alignment { get; set; } = HorizontalAlignment.Left;

    public bool ScaleToFit { get; set; } = false;

    public QuestPdfTable()
    {
    }

    public QuestPdfTable(List<float> columnsWidth)
    {
        ColumnsWidth = columnsWidth;
    }

    public QuestPdfTable(uint numberOfColumns)
    {
        // Create relative equally-sized columns
        for (uint i = 0; i < numberOfColumns; i++)
        {
            ColumnsWidth.Add(-1f);
        }
    }
}
