using System.Collections.Generic;

namespace Markdig.Renderers.Docx;

public class DocumentStyles
{
    public Dictionary<string, string> MarkdownStyles = new()
    {
        ["UnknownFormatting"] = "MDParagraph",
        
        ["Paragraph"] = "MDParagraph", 
        ["Quote"] = "MDQuote", 
        ["CodeBlock"] = "MDCodeBlock", 

        ["CodeInline"] = "MDCodeInline", 
        ["Hyperlink"] = "MDHyperlink", 

        ["HorizontalLine"] = "MDHorizontalLine", 

        ["ListBulletItem"] = "MDBulletedListItem", 
        ["ListOrderedItem"] = "MDOrderedListItem", 

        ["DefinitionTerm"] = "MDDefinitionTerm", 
        ["DefinitionItem"] = "MDDefinitionItem",

        ["Table"] = "MDTable",

        ["Heading1"] = "MDHeading1",
        ["Heading2"] = "MDHeading2",
        ["Heading3"] = "MDHeading3",
        ["Heading4"] = "MDHeading4",
        ["Heading5"] = "MDHeading5",
        ["Heading6"] = "MDHeading6"
    };

    public bool Contains(string styleName)
    {
        return MarkdownStyles.ContainsValue(styleName);
    }
}
