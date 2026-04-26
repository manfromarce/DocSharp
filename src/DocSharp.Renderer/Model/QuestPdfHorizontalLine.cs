using QuestPDF.Helpers;
using QuestPDF.Infrastructure;

namespace DocSharp.Renderer;

internal class QuestPdfHorizontalLine : QuestPdfBlock
{
    public float Thickness { get; set; } = 1;
    public Color? Color { get; set; }
}
