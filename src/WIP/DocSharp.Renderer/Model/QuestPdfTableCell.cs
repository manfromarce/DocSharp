using QuestPDF.Helpers;
using QuestPDF.Infrastructure;

namespace DocSharp.Renderer;

internal class QuestPdfTableCell : QuestPdfContainer
{
    internal Color? BackgroundColor = null;
    internal float LeftBorderThickness = 1.0f;
    internal float RightBorderThickness = 1.0f;
    internal float TopBorderThickness = 1.0f;
    internal float BottomBorderThickness = 1.0f;
    internal Color BordersColor = Colors.Black;
    internal float PaddingLeft = 0f;
    internal float PaddingRight = 0f;
    internal float PaddingTop = 0f;
    internal float PaddingBottom = 0f;
    internal VerticalAlignment VertAlignment = VerticalAlignment.Top;

    internal uint ColumnSpan { get; set; } = 1;
    internal uint RowSpan { get; set; } = 1;
    internal uint ColumnNumber { get; set; } = 1;
    internal uint RowNumber { get; set; } = 1;
}
