using System.Collections.Generic;

namespace Markdig.Renderers.Docx;

public class DocumentStyles
{
    public Dictionary<int, string?> Headings { get; } = new()
    {
        [0] = "Title",
        [1] = "MDHeading1",
        [2] = "MDHeading2",
        [3] = "MDHeading3",
        [4] = "MDHeading4",
        [5] = "MDHeading5",
        [6] = "MDHeading6",
    };

    public string UndefinedHeading { get; set; } = "MDHeading5";

    public string Paragraph { get; set; } = "MDParagraphTextBody";

    public string CodeBlock { get; set; } = "MDPreformattedText";

    public string UnknownFormatting { get; set; } = "MDNormal";
    public string Hyperlink { get; set; } = "MDHyperlink";

    public string ListOrdered { get; set; } = "MDListNumber";

    public string ListOrderedItem { get; set; } = "MDListNumberItem";
    public string ListBullet { get; set; } = "MDListBullet";
    public string ListBulletItem { get; set; } = "MDListBulletItem";
    
    public string? Quote { get; set; } = "MDQuotations";

    public string? HorizontalLine { get; set; } = "MDHorizontalLine";

    public string? CodeInline { get; set; } = "MDSourceText";
}
