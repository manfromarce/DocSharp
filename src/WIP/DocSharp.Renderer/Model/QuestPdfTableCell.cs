using QuestPDF.Helpers;
using QuestPDF.Infrastructure;

namespace DocSharp.Renderer;

internal class QuestPdfTableCell : QuestPdfContainer
{
    public Color? BackgroundColor = null;
    public float LeftBorderThickness = 1.0f;
    public float RightBorderThickness = 1.0f;
    public float TopBorderThickness = 1.0f;
    public float BottomBorderThickness = 1.0f;
    public Color BordersColor = Colors.Black;
    public float PaddingLeft = 0f;
    public float PaddingRight = 0f;
    public float PaddingTop = 0f;
    public float PaddingBottom = 0f;
    public float Height = 0f;
    public float MinHeight = 0f;
    public VerticalAlignment VertAlignment = VerticalAlignment.Top;
    public uint ColumnSpan { get; set; } = 1;
    public uint RowSpan { get; set; } = 1;
    public uint ColumnNumber { get; set; } = 1;
    public uint RowNumber { get; set; } = 1;
}
