using DocSharp.Helpers;
using DocSharp.Rtf.Tokens;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Xml;
using DocSharp.Writers;

namespace DocSharp.Rtf.Model;

internal class HtmlVisitor : INodeVisitor
{
    private UnitValue _defaultTabWidth = new UnitValue(0.5, UnitType.Inch);
    //private HtmlTag? _lastTag;
    private Stack<StyleList> _stack = new Stack<StyleList>();
    private IEnumerable<IToken> _stylesheet = new IToken[]
    {
        new ForegroundColor(new ColorValue(0, 0, 0)),
        new BackgroundColor(new ColorValue(255, 255, 255))
    };
    private readonly XmlWriter _writer;

    public UnitValue DefaultTabWidth
    {
        get { return _defaultTabWidth; }
        set { _defaultTabWidth = value.HasValue ? value : new UnitValue(0.5, UnitType.Inch); }
    }

    public RtfToHtmlSettings? Settings { get; set; }

    public HtmlVisitor(TextWriter writer, RtfToHtmlSettings? settings = null)
    {
        _writer = new HtmlTextWriter(writer, settings);
    }

    public HtmlVisitor(XmlWriter writer, RtfToHtmlSettings? settings = null)
    {
        _writer = writer;
        Settings = settings ?? new RtfToHtmlSettings();
    }

    private bool TryGetElementTag(ElementType primary, out HtmlTag? tag)
    {
        var tags = Settings?.ElementTags ?? RtfToHtmlSettings.DefaultTags;
        if (tags.TryGetValue(primary, out tag))
            return !string.IsNullOrEmpty(tag.Name);
        return false;
    }

    public void Visit(RtfXml document)
    {
        DefaultTabWidth = document.DefaultTabWidth;

        if (TryGetElementTag(ElementType.Meta, out var metaTag))
        {
            _writer.WriteStartElement("html");
            if (document.Metadata.Count > 0)
            {
                _writer.WriteStartElement("head");

                //_writer.WriteAttributeString("lang", ""); // TODO: retrieve language from RTF
                _writer.WriteStartElement(metaTag.Name);
                _writer.WriteAttributeString("http-equiv", "content-type");
                _writer.WriteAttributeString("content", "text/html; charset=utf-8");
                _writer.WriteEndElement();

                var title = document.Metadata.FirstOrDefault(k => k.Key is Title).Value as string;
                if (!string.IsNullOrEmpty(title))
                    _writer.WriteElementString("title", title);

                var baseUrl = document.Metadata.FirstOrDefault(k => k.Key is HyperlinkBase).Value as string;
                if (!string.IsNullOrEmpty(baseUrl))
                {
                    _writer.WriteStartElement("base");
                    _writer.WriteAttributeString("href", baseUrl);
                    _writer.WriteEndElement();
                }

                foreach (var meta in document.Metadata
                  .Where(k => !(k.Key is Title | k.Key is InternalVersion) && k.Value != null))
                {
                    _writer.WriteStartElement(metaTag.Name);
                    if (meta.Key is CreateTime)
                        _writer.WriteAttributeString("name", "DCTERMS.created");
                    else if (meta.Key is Subject)
                        _writer.WriteAttributeString("name", "DCTERMS.subject");
                    else if (meta.Key is Author)
                        _writer.WriteAttributeString("name", "DCTERMS.creator");
                    else if (meta.Key is Manager)
                        _writer.WriteAttributeString("name", "manager");
                    else if (meta.Key is Company)
                        _writer.WriteAttributeString("name", "company");
                    else if (meta.Key is Operator)
                        _writer.WriteAttributeString("name", "operator");
                    else if (meta.Key is Category)
                        _writer.WriteAttributeString("name", "category");
                    else if (meta.Key is Keywords)
                        _writer.WriteAttributeString("name", "keywords");
                    else if (meta.Key is Comment)
                        _writer.WriteAttributeString("name", "comment");
                    else if (meta.Key is DocComment)
                        _writer.WriteAttributeString("name", "comment");
                    else if (meta.Key is RevisionTime)
                        _writer.WriteAttributeString("name", "DCTERMS.modified");
                    else if (meta.Key is PrintTime)
                        _writer.WriteAttributeString("name", "print-time");
                    else if (meta.Key is BackupTime)
                        _writer.WriteAttributeString("name", "backup-time");
                    else if (meta.Key is Tokens.Version)
                        _writer.WriteAttributeString("name", "version");
                    else if (meta.Key is EditingTime)
                        _writer.WriteAttributeString("name", "editting-time");
                    else if (meta.Key is NumPages)
                        _writer.WriteAttributeString("name", "number-of-pages");
                    else if (meta.Key is NumWords)
                        _writer.WriteAttributeString("name", "number-of-words");
                    else if (meta.Key is NumChars)
                        _writer.WriteAttributeString("name", "number-of-characters");
                    else if (meta.Key is NumCharsWs)
                        _writer.WriteAttributeString("name", "number-of-non-whitespace-characters");

                    if (meta.Value is DateTime date)
                        _writer.WriteAttributeString("content", date.ToString("s"));
                    else
                        _writer.WriteAttributeString("content", meta.Value?.ToString() ?? "");

                    _writer.WriteEndElement();
                }
                Settings?.CustomHeadStyle?.Invoke(_writer);

                _writer.WriteEndElement();
            }
            document.Root.Visit(this);
            _writer.WriteEndElement();
        }
        else
        {
            var elements = document.Root.Elements().Where(e => TryGetElementTag(e.Type, out var _)).ToList();
            if (elements.Count == 1)
                elements[0].Visit(this);
            else
                document.Root.Visit(this);
        }
    }

    public void Visit(Element element)
    {
        if (!TryGetElementTag(element.Type, out var tag))
            return;

        if (element.Type == ElementType.TableCell && element.Styles.OfType<CellToken>().FirstOrDefault()?.MergeWithAbove == true)
            return;

        _writer.WriteStartElement(tag.Name);
        foreach (var attribute in tag.Attributes)
            _writer.WriteAttributeString(attribute.Key, attribute.Value);

        var elementStyles = (IEnumerable<IToken>)element.Styles;
        if (element.Type == ElementType.TableCell || element.Type == ElementType.TableHeaderCell)
            elementStyles = elementStyles.Concat(elementStyles.OfType<CellToken>().SelectMany(c => c.Styles));
        else if (element.Type == ElementType.Section || element.Type == ElementType.Document)
            elementStyles = elementStyles.Where(t => t.Type != TokenType.ParagraphFormat && t.Type != TokenType.RowFormat && t.Type != TokenType.CellFormat);
        var styleList = GetNewStyles(elementStyles, tag)
          .Where(t => !IsSpanElement(t))
          .ToList();
        if (element.Type != ElementType.TableCell && element.Type != ElementType.TableHeaderCell && element.Type != ElementType.TableRow && element.Type != ElementType.Table)
            styleList.RemoveWhere(t => t.Type == TokenType.CellFormat || t.Type == TokenType.RowFormat);

        if (element.Type == ElementType.OrderedList)
        {
            var numType = element.Styles.OfType<ListLevelType>().FirstOrDefault()?.Value
                ?? element.Styles.OfType<NumberingTypeToken>().FirstOrDefault()?.Value
                ?? NumberingType.Numbers;
            try
            {
                switch (numType)
                {
                    case NumberingType.LowerLetter:
                        _writer.WriteAttributeString("type", "a");
                        break;
                    case NumberingType.LowerRoman:
                        _writer.WriteAttributeString("type", "i");
                        break;
                    case NumberingType.UpperLetter:
                        _writer.WriteAttributeString("type", "A");
                        break;
                    case NumberingType.UpperRoman:
                        _writer.WriteAttributeString("type", "I");
                        break;
                }
            }
            catch (Exception e)
            {
                Debug.WriteLine(e.Message);
            }           

            var startAt = element.Styles.OfType<NumberingStart>().FirstOrDefault()?.Value ?? 1;
            if (startAt > 1)
                _writer.WriteAttributeString("start", startAt.ToString());
        }
        else if (element.Type == ElementType.TableCell || element.Type == ElementType.TableHeaderCell)
        {
            var colspan = element.Styles.OfType<CellToken>().FirstOrDefault()?.ColSpan ?? 1;
            if (colspan > 1)
                _writer.WriteAttributeString("colspan", colspan.ToString());

            var rowSpan = element.Styles.OfType<CellToken>().FirstOrDefault()?.RowSpan ?? 1;
            if (rowSpan > 1)
                _writer.WriteAttributeString("rowspan", rowSpan.ToString());
        }

        ProcessLeadingTabs(element, styleList);
        if (element.Type == ElementType.Section
          && element.Parent != null
          && element.Parent.Elements().First(e => e.Type == ElementType.Section) != element)
        {
            styleList.Add(new PageBreak());
        }

        if (styleList.Count > 0)
        {
            var css = new CssString(styleList, element.Type, elementStyles);
            if (css.Length > 0)
                _writer.WriteAttributeString("style", css.ToString());
        }

        if (element.Type == ElementType.Table)
            ProcessColumns(element);

        _stack.Push(new StyleList(tag.Styles.Where(t => t.Type == TokenType.CharacterFormat))
          .MergeRange(element.Styles));
        var anyNodes = false;
        foreach (var node in element.Nodes())
        {
            node.Visit(this);
            anyNodes = true;
        }
        if (!anyNodes)
        {
            if (element.Type == ElementType.Paragraph || element.Type == ElementType.ListItem)
                new Run() { Value = "\n" }.Visit(this);
            else if (element.Type == ElementType.TableCell || element.Type == ElementType.TableHeaderCell)
                new Run() { Value = "\u00a0" }.Visit(this);
        }
        _stack.Pop();

        _writer.WriteEndElement();
    }

    private void ProcessLeadingTabs(Element element, List<IToken> styles)
    {
        if (element.Nodes().FirstOrDefault() is Run run && run.Value.StartsWith("\t"))
        {
            var firstLineIndent = styles.OfType<FirstLineIndent>().FirstOrDefault()?.Value ?? UnitValue.Empty;

            var tabCount = 0;
            while (tabCount < run.Value.Length && run.Value[tabCount] == '\t')
                tabCount++;

            var newFirstLineIndent = IndentSize(styles, true, tabCount);
            if (newFirstLineIndent > firstLineIndent)
            {
                styles.RemoveWhere(t => t is FirstLineIndent);
                styles.Add(new FirstLineIndent(newFirstLineIndent));
            }
        }
    }

    private void ProcessColumns(Element table)
    {
        // Get all RightBoundary of header,body and all rows
        var rightBoundaries = new List<UnitValue>();
        foreach (var mainElement in table.Elements())
        {
            rightBoundaries.AddRange(mainElement.Styles.OfType<CellToken>().Select(c => c.RightBoundary));
            foreach (var row in mainElement.Elements())
            {
                rightBoundaries.AddRange(row.Styles.OfType<CellToken>().Select(c => c.RightBoundary));
            }
        }

        var boundaries = rightBoundaries
          .Distinct()
          .OrderBy(v => v)
          .Select(v => new CellIndex() { RightBoundary = v })
          .ToList();

        if (boundaries.Count < 1)
            return;

        var widths = new List<UnitValue>()
        {
          boundaries[0].RightBoundary
        };

        var cellIdx = 0;
        for (var i = 1; i < boundaries.Count; i++)
        {
            var width = boundaries[i].RightBoundary - boundaries[i - 1].RightBoundary;
            if (width.ToPx() > 0.25)
            {
                cellIdx++;
                widths.Add(width);
            }
            boundaries[i].Index = cellIdx;
        }
        var indexDict = boundaries.ToDictionary(b => b.RightBoundary, b => b.Index);

        var rows = table.Elements().ToList();
        if (!rows.Any(e => e.Type == ElementType.TableRow))
            rows = rows.SelectMany(e => e.Elements()).ToList();
        for (var j = 0; j < rows.Count; j++)
        {
            var row = rows[j];
            var cells = row.Elements().ToList();
            var startIndex = 0;
            for (var i = 0; i < cells.Count; i++)
            {
                var cell = cells[i].Styles.OfType<CellToken>().Single();
                int lastIndex;
                try { lastIndex = indexDict[cell.RightBoundary]; }
                catch { lastIndex = startIndex; }
                cell.Index = startIndex;
                cell.ColSpan = lastIndex - startIndex + 1;
                if (cell.ColSpan > 1)
                {

                }

                if (cell.MergeWithAbove)
                {
                    for (int k = j - 1; k >= 0; k--)
                    {
                        var prevRowCells = rows[k].Elements().ToArray();

                        var offset = 0;
                        foreach (var prevRowCell in prevRowCells)
                        {
                            var cellAbove = prevRowCell.Styles.OfType<CellToken>().Single();
                            if (offset == startIndex)
                            {

                                if (cellAbove.HasRowSpan)
                                {
                                    cellAbove.RowSpan++;
                                    break;
                                }
                            }

                            offset += cellAbove.ColSpan;
                        }
                    }
                }

                // Fix widths to be the widths instead of the right boundary when there is a discrepancy
                if (startIndex == lastIndex && cell.WidthUnit == CellWidthUnit.Twip)
                    widths[startIndex] = new UnitValue(cell.Width, UnitType.Twip);

                startIndex = lastIndex + 1;
            }
        }

        _writer.WriteStartElement("colgroup");
        foreach (var width in widths)
        {
            _writer.WriteStartElement("col");
            _writer.WriteAttributeString("style", new CssString().Append("width", width).ToString());
            _writer.WriteEndElement();
        }
        _writer.WriteEndElement();
    }

    public void Flush()
    {
        _writer.Flush();
    }

    private class CellIndex
    {
        public UnitValue RightBoundary { get; set; }
        public int Index { get; set; }
    }

    private IEnumerable<IToken> GetNewStyles(IEnumerable<IToken> styles, HtmlTag tag)
    {
        var existing = new StyleList(_stylesheet);
        if (_stack.Count > 0)
            existing.MergeRange(_stack.Peek());
        existing.MergeRange(tag.Styles);
        var requested = styles.ToList();
        var intersection = existing.Intersect(requested).ToList();

        var newStyles = requested.Where(t => !intersection.Contains(t)).ToList();
        var toNegate = existing.Where(t => !intersection.Contains(t)).ToList();
        foreach (var styleToNegate in toNegate)
        {
            if (styleToNegate is SpaceAfter && !newStyles.OfType<SpaceAfter>().Any())
                newStyles.Add(new SpaceAfter(new UnitValue(0, UnitType.Pixel)));
            else if (styleToNegate is SpaceBefore && !newStyles.OfType<SpaceBefore>().Any())
                newStyles.Add(new SpaceBefore(new UnitValue(0, UnitType.Pixel)));
            else if (styleToNegate is CellVerticalAlign && !newStyles.OfType<CellVerticalAlign>().Any())
                newStyles.Add(new CellVerticalAlign(VerticalAlignment.Top));
            else if (styleToNegate is BottomCellSpacing && !newStyles.OfType<BottomCellSpacing>().Any())
                newStyles.Add(new BottomCellSpacing(new UnitValue(0, UnitType.Pixel)));
            else if (styleToNegate is TopCellSpacing && !newStyles.OfType<TopCellSpacing>().Any())
                newStyles.Add(new TopCellSpacing(new UnitValue(0, UnitType.Pixel)));
            else if (styleToNegate is LeftCellSpacing && !newStyles.OfType<LeftCellSpacing>().Any())
                newStyles.Add(new LeftCellSpacing(new UnitValue(0, UnitType.Pixel)));
            else if (styleToNegate is RightCellSpacing && !newStyles.OfType<RightCellSpacing>().Any())
                newStyles.Add(new RightCellSpacing(new UnitValue(0, UnitType.Pixel)));
            else if (styleToNegate is TextAlign && tag.Styles.OfType<TextAlign>().Any())
                newStyles.Add(new TextAlign(TextAlignment.Left));
            else if (styleToNegate is ControlWord<bool> boolControl)
                newStyles.Add(ControlTag.Negate(boolControl));
        }

        return newStyles;
    }

    public void Visit(Run run)
    {
        var hyperlink = run.Styles.OfType<HyperlinkToken>().FirstOrDefault();
        var elementType = hyperlink == null ? ElementType.Span : ElementType.Hyperlink;

        if (!TryGetElementTag(elementType, out var tag))
            return;

        var styleList = new StyleList(GetNewStyles(run.Styles, tag)
          .Where(t => t.Type == TokenType.CharacterFormat));
        var stylesWritten = false;

        var convertSymbols = run.Styles.Any(t => t is Font font && 
                                                 FontConverter.IsNonUnicodeFont(font.Name));
        if (convertSymbols)
            styleList.RemoveWhere(t => t is Font);

        var endTags = 0;
        if (TryGetElementTag(ElementType.Strong, out var boldTag)
          && styleList.TryRemoveFirstTrue(out IsBold boldToken))
        {
            _writer.WriteStartElement(boldTag.Name);
            endTags++;
        }
        if (TryGetElementTag(ElementType.Emphasis, out var italicTag)
          && styleList.TryRemoveFirstTrue(out IsItalic italicToken))
        {
            _writer.WriteStartElement(italicTag.Name);
            endTags++;
        }
        if (hyperlink == null
          && TryGetElementTag(ElementType.Underline, out var underlineTag)
          && styleList.TryRemoveMany(StyleList.IsUnderline, out var underlineStyles))
        {
            _writer.WriteStartElement(underlineTag.Name);
            var underlineCss = new CssString(underlineStyles.Where(t => !(t is IsUnderline)), ElementType.Underline, run.Styles);
            if (underlineCss.Length > 0)
            {
                _writer.WriteAttributeString("style", underlineCss.ToString());
                stylesWritten = true;
            }
            endTags++;
        }
        if (styleList.TryRemoveFirstTrue(out IsStrikethrough strikeToken)
          || styleList.OfType<IsDoubleStrike>().FirstOrDefault()?.Value == true)
        {
            _writer.WriteStartElement("s");
            endTags++;
        }
        if (styleList.TryRemoveFirst(out SubscriptStart subToken))
        {
            _writer.WriteStartElement("sub");
            endTags++;
        }
        if (styleList.TryRemoveFirst(out SuperscriptStart superToken))
        {
            _writer.WriteStartElement("sup");
            endTags++;
        }
        if (styleList.TryRemoveFirst(out BackgroundColor highlight))
        {
            _writer.WriteStartElement("mark");
            styleList.RemoveWhere(s => s is ForegroundColor);
            var markCss = new CssString(GetNewStyles(run.Styles.Where(s => s is BackgroundColor || s is ForegroundColor), HtmlTag.Mark), ElementType.Span, run.Styles);
            if (markCss.Length > 0)
            {
                _writer.WriteAttributeString("style", markCss.ToString());
                stylesWritten = true;
            }
            endTags++;
        }

        var css = new CssString(styleList, elementType, run.Styles);
        if (hyperlink != null)
        {
            _writer.WriteStartElement(tag.Name);
            if (css.Length > 0)
                _writer.WriteAttributeString("style", css.ToString());
            if (!string.IsNullOrEmpty(hyperlink.Url))
                _writer.WriteAttributeString("href", hyperlink.Url);
            if (!string.IsNullOrEmpty(hyperlink.Target))
                _writer.WriteAttributeString("target", hyperlink.Target);
            if (!string.IsNullOrEmpty(hyperlink.Title))
                _writer.WriteAttributeString("title", hyperlink.Title);
            endTags++;
        }
        else if (css.Length > 0 && endTags > 0 && !stylesWritten)
        {
            _writer.WriteAttributeString("style", css.ToString());
        }
        else if (css.Length > 0)
        {
            _writer.WriteStartElement(tag.Name);
            _writer.WriteAttributeString("style", css.ToString());
            endTags++;
        }

        WriteRunText(run, convertSymbols);

        for (var j = 0; j < endTags; j++)
            _writer.WriteEndElement();
    }

    private void WriteRunText(Run run, bool convertSymbols)
    {
        var i = 0;
        var charBuffer = run.Value.ToCharArray();
        var eastAsian = run.Styles.OfType<Font>().Any(f => TextEncoding.IsEastAsian(f.Encoding));

        if (run.Value == " " && run.Parent?.Nodes().Count() == 1)
        {
            charBuffer[0] = eastAsian ? '\u2007' : '\u00a0';
        }
        else if (run.Parent?.Nodes().First() == run)
        {
            while (i < charBuffer.Length && charBuffer[i] == '\t')
                i++;
        }

        var start = i;
        var inTabList = false;
        while (i < charBuffer.Length)
        {
            if (charBuffer[i] == '\t')
            {
                if (!inTabList)
                {
                    if (start < i)
                        _writer.WriteChars(charBuffer, start, i - start);
                    start = i;
                    inTabList = true;
                }
            }
            else if (charBuffer[i] == '\n')
            {
                if (start < i && !inTabList)
                    _writer.WriteChars(charBuffer, start, i - start);
                inTabList = false;
                start = i + 1;
                _writer.WriteStartElement("br");
                _writer.WriteEndElement();
            }
            else
            {
                if (inTabList)
                {
                    WriteTabs(run.Parent?.Styles, start == 0 || charBuffer[start - 1] == '\n', i - start);
                    inTabList = false;
                    start = i;
                }
                else if (i > 0 && charBuffer[i] == ' '
                  && (charBuffer[i - 1] == ' ' || charBuffer[i - 1] == '\u00a0' || charBuffer[i - 1] == '\u2007'))
                {
                    charBuffer[i] = eastAsian ? '\u2007' : '\u00a0';
                }
                else if (convertSymbols)
                {
                    string fontName = run.Styles.OfType<Font>()?.FirstOrDefault()?.Name ?? "";
                    _writer.WriteValue(FontConverter.ToUnicode(fontName, charBuffer[i]));
                    start = i + 1;
                }
            }
            i++;
        }

        if (start < charBuffer.Length)
        {
            if (inTabList)
                WriteTabs(run.Parent?.Styles, start == 0 || charBuffer[start - 1] == '\n', charBuffer.Length - start);
            else
                _writer.WriteChars(charBuffer, start, charBuffer.Length - start);
        }

        if (charBuffer.Length > 0 && charBuffer[charBuffer.Length - 1] == '\n'
          && run.Parent != null && run.Parent.Nodes().Last() == run)
            _writer.WriteValue("\u00a0");
    }

    private UnitValue IndentSize(IEnumerable<IToken>? parentStyles, bool newLine, int tabCount)
    {
        var tabPositions = (parentStyles ?? Enumerable.Empty<IToken>()).OfType<TabPosition>().ToList();
        if (!newLine || tabPositions.Count < 1)
            return DefaultTabWidth * tabCount;
        else if (tabCount > tabPositions.Count)
            return tabPositions.Last().Value + DefaultTabWidth * (tabCount - tabPositions.Count);
        else
            return tabPositions[tabCount - 1].Value;
    }

    private void WriteTabs(IEnumerable<IToken>? parentStyles, bool newLine, int tabCount)
    {
        var size = IndentSize(parentStyles, newLine, tabCount);
        _writer.WriteStartElement("span");
        _writer.WriteAttributeString("style", $"display:inline-block;width:{size.ToPx().ToString(CultureInfo.InvariantCulture)}px");
        _writer.WriteEndElement();
    }

    internal static bool IsSpanElement(IToken token)
    {
        return token is IsBold
          || token is IsItalic
          || StyleList.IsUnderline(token)
          || token is IsStrikethrough
          || token is IsDoubleStrike
          || token is SubscriptStart
          || token is SuperscriptStart
          || token is HyperlinkToken
          || token is BackgroundColor;
    }

    public void Visit(Picture image)
    {
        var uri = Settings?.ImageUriGetter(image);
        if (!string.IsNullOrEmpty(uri))
        {
            _writer.WriteStartElement("img");

            if (image.Width.HasValue)
                _writer.WriteAttributeString("width", image.Width.ToPx().ToString("0", CultureInfo.InvariantCulture));

            if (image.Height.HasValue)
                _writer.WriteAttributeString("height", image.Height.ToPx().ToString("0", CultureInfo.InvariantCulture));

            _writer.WriteAttributeString("src", uri);
            _writer.WriteEndElement();
        }
    }

    public void Visit(ExternalPicture image)
    {
        _writer.WriteStartElement("img");
        _writer.WriteAttributeString("src", image.Uri);
        _writer.WriteEndElement();
    }

    public void Visit(Anchor anchor)
    {
        if (anchor.Type == AnchorType.Attachment)
        {
            Settings?.AttachmentRenderer(int.Parse(anchor.Id), _writer);
        }
        else
        {
            _writer.WriteStartElement("a");
            _writer.WriteAttributeString("id", anchor.Id);
            _writer.WriteEndElement();
        }
    }

    public void Visit(HorizontalRule horizontalRule)
    {
        _writer.WriteStartElement("hr");
        _writer.WriteAttributeString("style", "width:2in;border:0.5px solid black;margin-left:0");
        _writer.WriteEndElement();
    }

}
