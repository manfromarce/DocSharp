using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DocSharp.Helpers;

namespace DocSharp.Rtf.Model;

internal class TxtVisitor : INodeVisitor
{
    private TextWriter _writer;
    private bool isNumbered = false;

    public TxtVisitor(TextWriter writer)
    {
        _writer = writer;
    }

    public void Visit(RtfXml document)
    {
        var elements = document.Root.Elements().ToList();
        if (elements.Count == 1)
            elements[0].Visit(this);
        else
            document.Root.Visit(this);
    }

    public void Visit(Picture image)
    {
        // Not available in plain text
    }

    public void Visit(ExternalPicture image)
    {
        // Not available in plain text
    }

    void INodeVisitor.Visit(Anchor anchor)
    {
        // Not available in plain text
    }

    void INodeVisitor.Visit(HorizontalRule horizontalRule)
    {
        _writer.WriteLine();
        _writer.WriteLine("----------------------");
    }

    void INodeVisitor.Visit(Element element)
    {
        switch (element.Type)
        {           
            case ElementType.TableCell:
            case ElementType.TableHeaderCell:
                _writer.Write('|');
                break;
            case ElementType.List:
                isNumbered = false;
                break;
            case ElementType.OrderedList:
                isNumbered = true;
                break;
            case ElementType.ListItem:
                _writer.WriteLine();
                for (int index = 0; index < element.ListLevel; index++)
                {
                    _writer.Write("  ");
                }
                if (isNumbered)
                {
                    _writer.Write("1. ");
                }
                else
                {
                    _writer.Write("• ");
                }
                break;
        }
        foreach (var sub in element.Nodes())
        {
            sub.Visit(this);
        }
        switch (element.Type)
        {
            case ElementType.Heading1:
            case ElementType.Heading2:
            case ElementType.Heading3:
            case ElementType.Heading4:
            case ElementType.Heading5:
            case ElementType.Heading6:
            case ElementType.Paragraph:
            case ElementType.List:
            case ElementType.OrderedList:
                if (!element.IsLast())
                {
                    _writer.WriteLine();
                    if (!element.IsEmpty())
                    {
                        // Write additional blank line unless the paragraph is empty.
                        _writer.WriteLine();
                    }
                }
                break;
            case ElementType.Table:
                // Write blank line before the table.
                _writer.WriteLine();
                break;
            case ElementType.TableRow:
            case ElementType.TableHeader:
                _writer.Write('|');
                _writer.WriteLine();
                break;
        }
    }

    void INodeVisitor.Visit(Run run)
    {
        string font = run.Styles.OfType<Font>()?.FirstOrDefault()?.Name ?? string.Empty;
        string text = run.Value;

        foreach (char c in text)
        {
            _writer.Write(FontConverter.ToUnicode(font, c));
        }
    }
}
