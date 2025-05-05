using System.Collections.Generic;

namespace Markdig.Renderers.Docx;

public class DocumentStyles
{
    public Dictionary<string, string> MarkdownStyles = new()
    {
        ["UndefinedHeading"] = "MDHeading5", 
        ["UnknownFormatting"] = "MDNormal",
        
        ["Paragraph"] = "MDParagraphTextBody", 
        ["CodeBlock"] = "MDPreformattedText", 
        ["Quote"] = "MDQuotations", 
        ["HorizontalLine"] = "MDHorizontalLine", 

        ["ListOrdered"] = "MDListNumber", 
        ["ListOrderedItem"] = "MDListNumberItem", 
        ["ListBullet"] = "MDListBullet", 
        ["ListBulletItem"] = "MDListBulletItem", 

        ["Hyperlink"] = "MDHyperlink", 
        ["CodeInline"] = "CodeInline", 

        ["DefinitionTerm"] = "DefinitionTerm", 
        ["DefinitionItem"] = "DefinitionItem", 
    };

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

    public bool Contains(string styleName)
    {
        return MarkdownStyles.ContainsValue(styleName) || Headings.ContainsValue(styleName);
    }
}
