using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DocSharp.Helpers;
using DocSharp.Rtf.Tokens;
using DocSharp.Writers;

namespace DocSharp.Rtf.Model;

internal class MarkdownVisitor : INodeVisitor
{
    internal MarkdownStringWriter _writer;

    private int imageCount = 0;
    private int rowIndex = 0;
    private int cellIndex = 0;
    private int headerCells = 0;
    private bool isNumbered = false;
    private bool isInTableCell = false;

    public RtfToMdSettings Settings { get; set; }

    public MarkdownVisitor(MarkdownStringWriter writer, RtfToMdSettings? settings = null)
    {
        _writer = writer;
        Settings = settings ?? new RtfToMdSettings();
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
        try
        {
            if (Settings.ImagesOutputFolder != null)
            {
                ++imageCount;
                string fileName = Settings.BaseImageFileName + imageCount.ToString() + image.FileType();
#if NETFRAMEWORK
                string actualFilePath = Path.Combine(Settings.ImagesOutputFolder, fileName);
#else
                string actualFilePath = Path.Join(Settings.ImagesOutputFolder, fileName);
#endif
                Uri uri;
                if (Settings.ImagesBaseUriOverride is null)
                {
                    uri = new Uri(actualFilePath, UriKind.Absolute);
                }
                else
                {
                    string baseUri = UriHelpers.NormalizeBaseUri(Settings.ImagesBaseUriOverride);
                    uri = new Uri(baseUri + fileName, UriKind.RelativeOrAbsolute);
                }
                File.WriteAllBytes(actualFilePath, image.Bytes);
                _writer.Write($" ![{fileName}]({uri}) ");
            }
        }
        catch (Exception ex)
        {
            // Probably an issue with the output directory.
            // Don't stop the conversion.
#if DEBUG
            Debug.WriteLine("Visit Picture error: " + ex.Message);
            return;
#endif
        }
    }

    public void Visit(ExternalPicture image)
    {
        _writer.Write($" ![image]({image.Uri}) ");
    }

    void INodeVisitor.Visit(Anchor anchor)
    {
        if (anchor.Type == AnchorType.Bookmark)
        {
            _writer.Write($"<a id=\"{anchor.Id}\"></a>");
        }
    }

    void INodeVisitor.Visit(HorizontalRule horizontalRule)
    {
        _writer.WriteLine();
        _writer.WriteLine("-----");
    }

    void INodeVisitor.Visit(Element element)
    {
        switch (element.Type)
        {
            case ElementType.Heading1:
                _writer.Write("# ");
                break;
            case ElementType.Heading2:
                _writer.Write("## ");
                break;
            case ElementType.Heading3:
                _writer.Write("### ");
                break;
            case ElementType.Heading4:
                _writer.Write("#### ");
                break;
            case ElementType.Heading5:
                _writer.Write("##### ");
                break;
            case ElementType.Heading6:
                _writer.Write("###### ");
                break;
            case ElementType.Emphasis:
                _writer.Write("*");
                break;
            case ElementType.Strong:
                _writer.Write("**");
                break;
            case ElementType.Underline:
                _writer.Write("<u>");
                break;
            case ElementType.Table:
                // Write blank line before the table.
                _writer.WriteLine();
                break;
            case ElementType.TableRow:
            case ElementType.TableHeader:
                if (rowIndex == 1)
                {
                    // Write horizontal line after the table header or first row (mandatory in Markdown).
                    for (int i = 0; i < headerCells; ++i)
                    {
                        _writer.Write("| --- ");
                    }
                    if (headerCells > 0)
                    {
                        // Close horizontal line
                        _writer.Write('|');
                    }
                    _writer.WriteLine();
                    // Proceed with the first non-header row.
                }
                _writer.Write("| "); // Delimiter before the first cell or between cells.
                break;
            case ElementType.TableCell:
            case ElementType.TableHeaderCell:
                // Table header may not be present in the model built from RTF, but is mandatory in Markdown,
                // so we need to keep track of the row index and number of cells.
                if (rowIndex == 0)
                {
                    ++headerCells;
                }
                else
                {
                    ++cellIndex;
                    if (cellIndex > headerCells)
                    {
                        return; // Not supported
                    }
                }
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
                    _writer.Write("    ");
                }
                if (isNumbered)
                {
                    _writer.Write("1. ");
                }
                else
                {
                    _writer.Write("- ");
                }
                break;
        }

        if (element.Type == ElementType.TableCell ||
            element.Type == ElementType.TableHeaderCell)
        {            
            // Table cells in markdown don't support multiple lines directly, 
            // so we have to use the HTML <br> element.
            // In addition, horizontal lines and nested tables are not supported inside a table cell.
            _writer.NewLine = "<br/>";
            isInTableCell = true;
            foreach (var sub in element.Nodes().Where(n => n is not HorizontalRule && !(n is Element el &&
                                                       (el.Type == ElementType.Table ||
                                                        el.Type == ElementType.TableBody ||
                                                        el.Type == ElementType.TableRow ||
                                                        el.Type == ElementType.TableCell ||
                                                        el.Type == ElementType.TableHeader ||
                                                        el.Type == ElementType.TableHeaderCell))))
            {
                sub.Visit(this);
            }
            isInTableCell = false;
            _writer.NewLine = Environment.NewLine;
        }
        else
        {
            foreach (var sub in element.Nodes())
            {
                sub.Visit(this);
            }
        }
        
        switch (element.Type)
        {
            case ElementType.Emphasis:
                _writer.Write("*");
                break;
            case ElementType.Strong:
                _writer.Write("**");
                break;
            case ElementType.Underline:
                _writer.Write("</u>");
                break;
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
            case ElementType.TableCell:
            case ElementType.TableHeaderCell:
                // Close table cell
                _writer.Write(" | ");               
                break;
            case ElementType.TableRow:
                if (rowIndex > 0)
                {
                    while (cellIndex < headerCells)
                    {
                        // Write empty cells if needed
                        _writer.Write(" | ");
                        ++cellIndex;
                    }
                }
                _writer.WriteLine();
                // End of table row, increase index.
                ++rowIndex;
                cellIndex = 0;
                break;
            case ElementType.TableHeader:
                _writer.WriteLine();
                // End of table row, increase index.
                ++rowIndex;
                cellIndex = 0;
                break;
            case ElementType.Table:
                // End of table, reset variables.
                rowIndex = 0;
                headerCells = 0;
                cellIndex = 0;
                _writer.WriteLine();
                break;
        }
    }

    void INodeVisitor.Visit(Run run)
    {
        var hyperlink = run.Styles.OfType<HyperlinkToken>().FirstOrDefault();

        if (hyperlink != null)
        {
            _writer.Write($" [{run.Value}]({hyperlink.Url}) ");
        }
        else
        {
            bool isBold = run.Styles.OfType<IsBold>().FirstOrDefault() != null;
            bool isItalic = run.Styles.OfType<IsItalic>().FirstOrDefault() != null;
            bool isUnderline = run.Styles.OfType<IsUnderline>().FirstOrDefault() != null;
            bool isMarked = run.Styles.OfType<BackgroundColor>().FirstOrDefault() is BackgroundColor bc && 
                            bc.Value != ColorValue.White;
            bool isSuperscript = run.Styles.OfType<SuperscriptStart>().FirstOrDefault() != null;
            bool isSubscript = run.Styles.OfType<SubscriptStart>().FirstOrDefault() != null;
            bool isStrikethrough = run.Styles.OfType<IsStrikethrough>().FirstOrDefault() != null ||
                                   run.Styles.OfType<IsDoubleStrike>().FirstOrDefault() != null;

            if (isItalic)
                _writer.Write("*");
            if (isBold)
                _writer.Write("**");
            if (isStrikethrough)
                _writer.Write("~~");
            if (isUnderline)
                _writer.Write("<u>");
            if (isSubscript)
                _writer.Write("<sub>");
            else if (isSuperscript)
                _writer.Write("<sup>");
            if (isMarked)
                _writer.Write("<mark>");

            string font = run.Styles.OfType<Font>()?.FirstOrDefault()?.Name ?? string.Empty;
            string text = run.Value;
            foreach (char c in text)
                _writer.WriteCharEscaped(c, font, isInTableCell);
            
            if (isMarked)
                _writer.Write("</mark>");
            if (isSubscript)
                _writer.Write("</sub>");
            else if (isSuperscript)
                _writer.Write("</sup>");
            if (isUnderline)
                _writer.Write("</u>");
            if (isStrikethrough)
                _writer.Write("~~");
            if (isItalic)
                _writer.Write("*");
            if (isBold)
                _writer.Write("**");
        }
    }
}
