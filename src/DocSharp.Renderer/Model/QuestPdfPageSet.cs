using QuestPDF.Helpers;
using QuestPDF.Infrastructure;

namespace DocSharp.Renderer;

internal class QuestPdfPageSet(float pageWidth, float pageHeight, 
                               float marginLeft, float marginTop, float marginRight, float marginBottom,
                               Unit unit)
{
    internal PageSize PagesSize { get; set; } = new PageSize(pageWidth, pageHeight, unit);

    internal float MarginLeft { get; set; } = marginLeft;
    internal float MarginTop { get; set; } = marginTop;
    internal float MarginRight { get; set; } = marginRight;
    internal float MarginBottom { get; set; } = marginBottom;

    internal Unit Unit { get; set; } = unit;
    internal Color BackgroundColor { get; set; } = Colors.White;

    // TODO: page borders (not directly supported by QuestPDF, we would need to draw them manually)

    internal int NumberOfColumns { get; set; } = 1;
    internal float? SpaceBetweenColumns { get; set; } // in points (if set)

    internal QuestPdfContainer Header = new();
    internal QuestPdfContainer Footer = new();
    internal QuestPdfContainer Content = new();
}
