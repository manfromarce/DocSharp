using System.Collections.Generic;
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

    internal QuestPdfContainer? HeaderFirst;
    internal QuestPdfContainer? HeaderEven;
    internal QuestPdfContainer HeaderOddOrDefault = new();

    internal QuestPdfContainer? FooterFirst;
    internal QuestPdfContainer? FooterEven;
    internal QuestPdfContainer FooterOddOrDefault = new();
    
    internal List<QuestPdfFootnote> Footnotes = new();
    internal List<QuestPdfEndnote> Endnotes = new();

    internal QuestPdfContainer Content = new();

    internal bool DifferentHeaderFooterForOddAndEvenPages => HeaderEven != null;
    // Note: HeaderEven and FooterEven are both null or both not null considering how the model is built
    // (based on the Open XML structure) 

    internal bool DifferentHeaderFooterForFirstPage => HeaderFirst != null;
    // Note: HeaderFirst and FooterFirst are both null or both not null considering how the model is built
    // (based on the Open XML structure) 
}
