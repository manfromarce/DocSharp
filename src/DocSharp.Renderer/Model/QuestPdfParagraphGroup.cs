using System.Collections.Generic;
using QuestPDF.Helpers;
using QuestPDF.Infrastructure;

namespace DocSharp.Renderer;

internal class QuestPdfParagraphGroup : QuestPdfBlock
{
    internal List<QuestPdfParagraph> Paragraphs = new();

    public float LeftBorderThickness = 0.0f;
    public float RightBorderThickness = 0.0f;
    public float TopBorderThickness = 0.0f;
    public float BottomBorderThickness = 0.0f;
    public Color BordersColor = Colors.Black;

    // Indentation applied to the whole group (container padding)
    public float LeftIndent = 0f;
    public float RightIndent = 0f;

    // Spacing outside the group (applied to the container)
    public float SpaceBefore = 0f;
    public float SpaceAfter = 0f;
}
