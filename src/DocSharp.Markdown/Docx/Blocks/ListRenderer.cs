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
    //private AbstractNum? _bulletListAbstractNum = null;
    //private AbstractNum? _orderedListAbstractNum = null;

    protected override void WriteObject(DocxDocumentRenderer renderer, ListBlock obj)
    {
        var listInfo = new ListInfo();
        var numbering = renderer.Document.GetOrCreateNumbering().NumberingDefinitionsPart!.Numbering;
        var listStyle = obj.IsOrdered ? renderer.Styles.MarkdownStyles["ListOrdered"] : renderer.Styles.MarkdownStyles["ListBullet"];
        var listItemStyle = obj.IsOrdered ? renderer.Styles.MarkdownStyles["ListOrderedItem"] : renderer.Styles.MarkdownStyles["ListBulletItem"];
        listInfo.StyleId = listItemStyle;

        var abstractNum = numbering.Elements<AbstractNum>().FirstOrDefault(e => e.StyleLink?.Val == listStyle);
        if (abstractNum?.AbstractNumberId != null) // TODO: Fallback and create this
        {
            int abstractNumId = abstractNum.AbstractNumberId.Value;

            //var numberingId = numbering.Elements<NumberingInstance>().Where(n => n.AbstractNumId?.Val != null &&
            //                                                                     n.AbstractNumId.Val == abstractNumId)
            //                                                         .FirstOrDefault();

            var newNumberingId = numbering.Elements<NumberingInstance>().Count() + 1;
            var numberingInstance = new NumberingInstance
            {
                NumberID = newNumberingId,
                AbstractNumId = new AbstractNumId
                {
                    Val = abstractNum.AbstractNumberId
                }
            };
            listInfo.NumberingInstance = numberingInstance;
            if (obj.IsOrdered)
            {
                for (var i = 0; i <= 8; i++)
                {
                    var lvlOverride = new LevelOverride
                    {
                        LevelIndex = i,
                        StartOverrideNumberingValue = new StartOverrideNumberingValue()
                        {
                            Val = int.TryParse(obj.OrderedStart, out int startNumber) ? startNumber : 1
                        }
                    };
                    listInfo.NumberingInstance.AppendChild(lvlOverride);
                }
            }
            numbering.AddNumberingInstance(listInfo.NumberingInstance);

            if (renderer.ActiveList.Count == 0)
            {
                listInfo.Level = 0;
            }
            else
            {
                var previousList = renderer.ActiveList.Peek();
                listInfo.Level = Math.Min(previousList.Level + 1, 8); 
                // 8 seems to be the maximum level in DOCX documents (9 levels including 0).
            }
            renderer.ActiveList.Push(listInfo);
            renderer.WriteChildren(obj);
            renderer.ActiveList.Pop();
        }
    }
}
