using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DocumentFormat.OpenXml.Wordprocessing;
using DocSharp.Writers;
using System.IO;
using DocumentFormat.OpenXml.Packaging;
using System.Globalization;
using DocSharp.Helpers;

namespace DocSharp.Docx;

public partial class DocxToHtmlConverter : DocxToTextConverterBase<HtmlStringWriter>
{
    /// <summary>
    /// Image converter to preserve TIFF, EMF and other image types when converting to HTML. 
    /// If the DocSharp.ImageSharp or DocSharp.SystemDrawing package is installed, 
    /// this property can be set to a new instance of ImageSharpConverter or SystemDrawingConverter. 
    /// </summary>
    public IImageConverter? ImageConverter { get; set; } = null;

    /// <summary>
    /// If this property is set to an existing directory, images will be exported to that folder
    /// and a reference will be added in HTML syntax,
    /// otherwise images are preserved as base64. 
    /// NOTE: if the directory contains image files with the same names as in the DOCX document archive 
    /// (usually image1.*, image2.*, ...), they will be overwritten.
    /// </summary>
    public string? ImagesOutputFolder { get; set; } = string.Empty;

    /// <summary>
    /// This property is used in combination with ImagesOutputFolder to determine 
    /// how the image files URLs are specified in HTML.
    /// If images are exported as base64, this property is ignored.
    /// 
    /// If this property is set to null, an absolute path such as "file:///c:/.../image.jpg" 
    /// will be created using the ImagesOutputFolder value and the image file name.
    /// 
    /// Otherwise, the base path (excluding the image file name) is replaced by this value.
    /// Possible values:
    /// - empty string or "." : images are expected to be in the same folder as the HTML file.
    /// - relative paths such as "images" or "../images": images are expected to be in a subfolder or parent folder.
    /// - "/server/user/files/" or "C:\images": replaces the file path entirely
    /// (the image file name is still appended and Windows paths are converted to the file URI scheme).
    /// 
    /// This property does not affect where the images are actually saved, and can be useful if
    /// the HTML document is not saved to file, or in environments with limited file system access.
    /// </summary>
    public string? ImagesBaseUriOverride { get; set; } = null;

    internal override void ProcessDocument(Document document, HtmlStringWriter sb)
    {
        sb.AppendHtmlHeader(); // TODO: title
        sb.AppendStartTag("body");
        if (document.DocumentBackground is DocumentBackground bg)
        {
            ProcessDocumentBackground(bg, sb);
        }

        // Process body content
        if (document.Body is Body body)
        {
            base.ProcessBody(body, sb);
        }

        sb.AppendEndTag("body");
        sb.AppendEndTag("html");
    }

    internal override void ProcessDocumentBackground(DocumentBackground background, HtmlStringWriter sb)
    {
        if (background.Color != null)
        {
            string color = $"#{background.Color.Value}";
            sb.AppendTag("style", $"body {{ background-color: {color}; }}");
        }
        //else if (background.Background != null)
        //{
        // TODO (requires VML support)
        //}
    }

    private void ProcessTextDirection(TextDirectionValues value, ref List<string> styles)
    {
        if (value == TextDirectionValues.LefToRightTopToBottom ||
            value == TextDirectionValues.LeftToRightTopToBottom2010)
        {
            // Horizontal text, left to right (default)
            styles.Add("writing-mode: horizontal-tb;");
        }
        if (value == TextDirectionValues.TopToBottomRightToLeft ||
            value == TextDirectionValues.TopToBottomRightToLeft2010)
        {
            // Horizontal text, right to left
        }
        if (value == TextDirectionValues.BottomToTopLeftToRight ||
            value == TextDirectionValues.BottomToTopLeftToRight2010)
        {
            // Horizontal text, bottom to top
        }
        if (value == TextDirectionValues.LefttoRightTopToBottomRotated ||
            value == TextDirectionValues.LeftToRightTopToBottomRotated2010 ||
            value == TextDirectionValues.TopToBottomLeftToRightRotated ||
            value == TextDirectionValues.TopToBottomLeftToRightRotated2010)
        {
            // Vertical text
            styles.Add("writing-mode: vertical-lr;");
            styles.Add("text-orientation: upright;");
        }
        if (value == TextDirectionValues.TopToBottomRightToLeftRotated ||
            value == TextDirectionValues.TopToBottomRightToLeftRotated2010)
        {
            // Vertical text
            styles.Add("writing-mode: vertical-rl;");
            styles.Add("text-orientation: upright;");
        }
    }

    internal override void ProcessContentPart(ContentPart contentPart, HtmlStringWriter sb)
    {
        // MathML, SVG and SMIL are supported by most browsers
        var id = contentPart.Id;
        var mainDocumentPart = OpenXmlHelpers.GetMainDocumentPart(contentPart);
        if (id?.Value != null)
        {
            var part = mainDocumentPart?.GetPartById(id.Value);
            if (part != null)
            {
                // Read the part content
                using (var stream = part.GetStream())
                {
                    // Check if the part is a MathML, SVG or SMIL
                    if (part.ContentType == "application/mathml+xml" ||
                       part.ContentType == "application/mathml-presentation+xml" ||
                       part.ContentType == "application/mathml-content+xml" ||
                       part.ContentType == "image/svg+xml" ||
                       part.ContentType == "application/smil+xml")
                    {
                        // Read the content and append it to the HTML
                        using (var reader = new StreamReader(stream))
                        {
                            string content = reader.ReadToEnd();
                            if (content != null)
                            {
                                sb.AppendLine(content);
                            }
                        }
                    }
                }
            }
        }
    }

    internal override void ProcessText(Text text, HtmlStringWriter sb)
    {
        string font = string.Empty;
        if (text.Parent is Run run)
        {
            var fonts = OpenXmlHelpers.GetEffectiveProperty<RunFonts>(run);
            font = fonts?.Ascii?.Value?.ToLowerInvariant() ?? string.Empty;
        }
        string t = text.InnerText;
        var stringInfo = new StringInfo(t);
        for (int i = 0; i < stringInfo.LengthInTextElements; i++)
        {
            string textElement = stringInfo.SubstringByTextElements(i, 1);
            sb.Append(textElement, font);
        }
    }

    internal override void ProcessSymbolChar(SymbolChar symbolChar, HtmlStringWriter sb)
    {
        if (!string.IsNullOrEmpty(symbolChar?.Char?.Value))
        {
            string hexValue = symbolChar?.Char?.Value!;
            if (hexValue.StartsWith("0x", StringComparison.OrdinalIgnoreCase) ||
                hexValue.StartsWith("&h", StringComparison.OrdinalIgnoreCase))
            {
                hexValue = hexValue.Substring(2);
            }
            string htmlEntity = string.Empty;
            if (int.TryParse(hexValue, NumberStyles.HexNumber, CultureInfo.InvariantCulture, out int decimalValue))
            {
                if (!string.IsNullOrEmpty(symbolChar?.Font?.Value))
                {
                    htmlEntity = FontConverter.ToUnicode(symbolChar!.Font!.Value!, (char)decimalValue);
                }
            }
            if (string.IsNullOrEmpty(htmlEntity)) // If htmlEntity is empty, use the original char code
            {
                htmlEntity = $"&#{decimalValue.ToStringInvariant()};";
            }
            sb.Append(htmlEntity);
        }
    }

    internal override void ProcessBookmarkStart(BookmarkStart bookmark, HtmlStringWriter sb)
    {
        if (bookmark.Name != null)
            sb.AppendTag("a", null, ("id", bookmark.Name.Value));
    }

    internal override void ProcessBookmarkEnd(BookmarkEnd bookmarkEnd, HtmlStringWriter sb)
    {
    }

    internal override void ProcessBreak(Break @break, HtmlStringWriter sb)
    {
        if (@break.Type != null && @break.Type == BreakValues.Page)
        {
            sb.AppendTag("div", null, ("style", "break-after: page;"));
        }
        else if (@break.Type != null && @break.Type == BreakValues.Column)
        {
            sb.AppendTag("div", null, ("style", "break-after: column;"));
        }
        else
        {
            sb.AppendBreak();
        }
    }

    internal override void ProcessHyperlink(Hyperlink hyperlink, HtmlStringWriter sb)
    {
        bool hasUrl = false;
        if (hyperlink.Id?.Value is string rId)
        {
            var maindDocumentPart = OpenXmlHelpers.GetMainDocumentPart(hyperlink);
            if (maindDocumentPart?.HyperlinkRelationships.FirstOrDefault(x => x.Id == rId) is HyperlinkRelationship relationship)
            {
                string url = relationship.Uri.ToString();
                hasUrl = true;
                sb.Append($"<a href=\"{url}\">");
            }
        }
        else if (hyperlink.Anchor?.Value is string anchor)
        {
            hasUrl = true;
            sb.Append($"<a href=\"#{anchor}\">");
        }
        foreach (var element in hyperlink.Elements())
        {
            base.ProcessParagraphElement(element, sb);
        }
        if (hasUrl)
        {
            sb.Append("</a>");
        }
    }

    internal void ProcessShading(Shading? shading, ref List<string> styles)
    {
        if (shading != null && shading.Fill?.Value is string fill && fill.Length == 6)
        {
            styles.Add($"background-color: #{fill};");

            // Not supported: foreground (pattern)
        }
    }

    internal override void ProcessPageNumber(PageNumber pageNumber, HtmlStringWriter sb)
    {
    }

    internal override void ProcessPositionalTab(PositionalTab posTab, HtmlStringWriter sb)
    {
    }

    internal override void ProcessFieldChar(FieldChar field, HtmlStringWriter sb)
    {
    }

    internal override void ProcessFieldCode(FieldCode field, HtmlStringWriter sb)
    {
    }
}
