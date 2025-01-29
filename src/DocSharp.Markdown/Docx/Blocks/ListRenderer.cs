using System;
using System.Diagnostics;
using System.Linq;
using DocumentFormat.OpenXml.Wordprocessing;
using Markdig.Syntax;
using DocSharp.Docx;

namespace Markdig.Renderers.Docx.Blocks;

public class ListInfo
{
    public NumberingInstance? NumberingInstance { get; set; }
    
    public string? StyleId { get; set; }
    
    public int Level { get; set; }
}

public class ListRenderer : DocxObjectRenderer<ListBlock>
{
    private AbstractNum? _bulletListAbstractNum = null;
    private AbstractNum? _orderListAbstractNum = null;
    
    protected override void WriteObject(DocxDocumentRenderer renderer, ListBlock obj)
    {
        var listInfo = new ListInfo();
        
        if (renderer.ActiveList.Count == 0)
        {
            // We're creating new numbering for the list so item can reference it
            var listStyle = obj.IsOrdered ? renderer.Styles.MarkdownStyles["ListOrdered"] : renderer.Styles.MarkdownStyles["ListBullet"];
            var listItemStyle = obj.IsOrdered ? renderer.Styles.MarkdownStyles["ListOrderedItem"] : renderer.Styles.MarkdownStyles["ListBulletItem"];

            listInfo.StyleId = listItemStyle;
            listInfo.Level = 1;
            
            // Find or create abstract numbering associated with this style
            var numbering = renderer.Document.GetOrCreateNumbering().NumberingDefinitionsPart!.Numbering;

            var abstractNum = numbering
                .Elements<AbstractNum>().FirstOrDefault(e => e.StyleLink?.Val == listStyle);

            if (abstractNum == null)
            {
                throw new FormatException(
                    $"Failed to identify abstract numbering associated with list style ${listStyle}");
                // TODO Fallback and create this
            }

            var newNumberingId = numbering.Elements<NumberingInstance>().Count() + 1;
            listInfo.NumberingInstance = new NumberingInstance
            {
                NumberID = newNumberingId,
                AbstractNumId = new AbstractNumId {Val = abstractNum.AbstractNumberId}
            };

            if (obj.IsOrdered)
            {
                for (var i = 0; i <= 8; i++)
                {
                    var lvlOverride = new LevelOverride
                    {
                        LevelIndex = i,
                        StartOverrideNumberingValue = new StartOverrideNumberingValue() {Val = 1}
                    };
                    listInfo.NumberingInstance.AppendChild(lvlOverride);
                }
            }

            numbering.AddNumberingInstance(listInfo.NumberingInstance);
        }
        else
        {
            var previousList = renderer.ActiveList.Peek();
            listInfo.NumberingInstance = previousList.NumberingInstance;
            listInfo.StyleId = previousList.StyleId;
            listInfo.Level = previousList.Level + 1;
        }
        
        renderer.ActiveList.Push(listInfo);
        renderer.WriteChildren(obj);
        renderer.ActiveList.Pop();
    }
}
