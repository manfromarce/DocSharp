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
using System.Xml;
using System.Diagnostics;
using DocumentFormat.OpenXml;

namespace DocSharp.Docx;

public partial class DocxToHtmlConverter : DocxToTextWriterBase<HtmlTextWriter>
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

    /// <summary>
    /// Since HTML is not paginated, only the header of the first section and
    /// footer of the last section are exported.
    /// Set this property to false to ignore headers and footers.
    /// </summary>
    public bool ExportHeaderFooter { get; set; } = true;

    /// <summary>
    /// Since HTML is not paginated, both footnotes and endnotes are exported at the end of the document.
    /// Set this property to false to ignore footnotes and endnotes.
    /// </summary>
    public bool ExportFootnotesEndnotes { get; set; } = true;

    /// <summary>
    /// Convert a DOCX file to the output format.
    /// </summary>
    /// <param name="inputDocument">The WordprocessingDocument to use.</param>
    /// <param name="writer">The output writer.</param>
    public override void Convert(WordprocessingDocument inputDocument, TextWriter writer)
    {
        using (var htmlWriter = new HtmlTextWriter(writer))
        {
            var document = inputDocument.MainDocumentPart?.Document;
            if (document != null)
            {
                ProcessDocument(document, htmlWriter);
            }
        }
    }

    internal override void ProcessDocument(Document document, HtmlTextWriter sb)
    {
        sb.WriteHtmlHeader(document.MainDocumentPart?.OpenXmlPackage.PackageProperties.Title);
        sb.WriteStartElement("body");
        if (document.DocumentBackground is DocumentBackground bg)
        {
            ProcessDocumentBackground(bg, sb);
        }

        // Process body content
        if (document.Body is Body body)
        {
            base.ProcessBody(body, sb);
        }

        sb.WriteEndElement("body");
        sb.WriteEndElement("html");
    }

    internal override void ProcessDocumentBackground(DocumentBackground background, HtmlTextWriter sb)
    {
        if (ColorHelpers.EnsureHexColor(background.Color?.Value) is string color)
        {
            sb.WriteElementString("style", $"body {{ background-color: #{color}; }}");
        }
        //else if (background.Background != null)
        //{
        // TODO (requires VML support)
        //}
    }

    internal override void EnsureSpace(HtmlTextWriter sb)
    {
        sb.WriteStartElement("p");
        sb.WriteAttributeString("style", "margin-top: 10px;");
        sb.WriteEndElement("p");
    }

    private void ProcessTextDirection(TextDirectionValues textDirection, ref List<string> styles, out bool isVertical)
    {
        // Possible CSS properties: direction, unicode-bidi, text-orientation, writing-mode

        isVertical = false;
        if (textDirection == TextDirectionValues.LefToRightTopToBottom ||
           textDirection == TextDirectionValues.LeftToRightTopToBottom2010 ||
           textDirection == TextDirectionValues.LefttoRightTopToBottomRotated ||
           textDirection == TextDirectionValues.LeftToRightTopToBottomRotated2010)
        {
            // Horizontal text, left to right
            // (there seems to be no difference in DOCX between LefToRightTopToBottom and LefttoRightTopToBottomRotated)
            styles.Add("writing-mode: horizontal-tb;");
        }        
        if (textDirection == TextDirectionValues.TopToBottomLeftToRightRotated ||
            textDirection == TextDirectionValues.TopToBottomLeftToRightRotated2010)
        {
            // Vertical text (rotated letters), top to bottom, left to right
            styles.Add("writing-mode: vertical-lr;");
            isVertical = true;
        }
        if (textDirection == TextDirectionValues.TopToBottomRightToLeft ||
            textDirection == TextDirectionValues.TopToBottomRightToLeft2010 ||
            textDirection == TextDirectionValues.TopToBottomRightToLeftRotated ||
            textDirection == TextDirectionValues.TopToBottomRightToLeftRotated2010)
        {
            // Vertical text (rotated letters), top to bottom, right to left
            // (there seems to be no difference in DOCX between TopToBottomRightToLeft and TopToBottomRightToLeftRotated)
            styles.Add("writing-mode: sideways-rl;"); // or vertical-rl
            isVertical = true;
        }
        if (textDirection == TextDirectionValues.BottomToTopLeftToRight ||
            textDirection == TextDirectionValues.BottomToTopLeftToRight2010)
        {
            // Vertical text (rotated letters), bottom to top, left to right
            styles.Add("writing-mode: sideways-lr;");
            isVertical = true;
        }        
    }

    internal override void ProcessContentPart(ContentPart contentPart, HtmlTextWriter writer)
    {
        // MathML, SVG and SMIL are supported by most browsers.
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
                       part.ContentType == "application/smil+xml" ||
                       part.ContentType == "text/html" ||
                       part.ContentType == "application/xhtml+xml")
                    {
                        // Read the content and append it to the HTML.
                        // TODO: skip DOCTYPE.
                        try
                        {
                            using (var reader = XmlReader.Create(stream))
                            {
                                writer.WriteNode(reader, true);
                            }
                        }
                        catch (Exception ex)
                        {
#if DEBUG
                            Debug.WriteLine("Error in ProcessContentPart: " + ex.Message);
#endif
                        }
                    }
                }
            }
        }
    }

    internal override void ProcessAltChunk(AltChunk altChunk, HtmlTextWriter writer)
    {
        var id = altChunk.Id;
        var mainDocumentPart = OpenXmlHelpers.GetMainDocumentPart(altChunk);
        if (id?.Value != null)
        {
            var part = mainDocumentPart?.GetPartById(id.Value);
            if (part is AlternativeFormatImportPart alternativeFormatImportPart)
            {
                try
                {
                    // Read the part content
                    using (var stream = part.GetStream())
                    {
                        // Check the AltChunk MIME type.
                        if (alternativeFormatImportPart.ContentType == AlternativeFormatImportPartType.Html.ContentType)
                        {
                            // Read the content and append it to the HTML.
                            // TODO: skip DOCTYPE, <html>, <body>
                            using (var sr = new StreamReader(stream))
                            {
                                writer.WriteRaw('\n');
                                writer.WriteRaw(sr.ReadToEnd());
                                writer.WriteRaw('\n');
                            }
                        }
                        else
                        {
                            base.ProcessAltChunk(altChunk, writer);
                        }
                    }
                }
                catch (Exception ex)
                {
#if DEBUG
                    Debug.WriteLine("Error in ProcessAltChunk: " + ex.Message);
#endif
                }
            }
        }
    }

    internal override void ProcessText(Text text, HtmlTextWriter sb)
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
            if (textElement == "\r")
            {
                // Ignore as it's usually followed by \n
            }
            else if (textElement == "\n")
            {
                // Line endings are not converted to <br> in the Write method, because 
                // it's not valide in other contexts such as attributes strings.
                // So, replace them here.
                sb.WriteBreak();
            }
            else
            {
                sb.Write(textElement, font);
            }
        }
    }

    internal override void ProcessSymbolChar(SymbolChar symbolChar, HtmlTextWriter sb)
    {
        if (!string.IsNullOrEmpty(symbolChar?.Char?.Value))
        {
            string hexValue = symbolChar?.Char?.Value!;
            if (hexValue.StartsWith("0x", StringComparison.OrdinalIgnoreCase) ||
                hexValue.StartsWith("&h", StringComparison.OrdinalIgnoreCase))
            {
                hexValue = hexValue.Substring(2);
            }
            if (int.TryParse(hexValue, NumberStyles.HexNumber, CultureInfo.InvariantCulture, out int decimalValue))
            {
                if (!string.IsNullOrEmpty(symbolChar?.Font?.Value))
                {
                    string unicode = FontConverter.ToUnicode(symbolChar!.Font!.Value!, (char)decimalValue);
                    sb.WriteString(unicode);
                }
                else // use the original char code
                {
                    sb.WriteString((char)decimalValue);
                }
            }
        }
    }

    internal override void ProcessBookmarkStart(BookmarkStart bookmark, HtmlTextWriter writer)
    {
        if (bookmark.Name != null)
        {
            writer.WriteStartElement("a");
            writer.WriteAttributeString("id", bookmark.Name.Value);
            writer.WriteEndElement();
        }
    }

    internal override void ProcessBookmarkEnd(BookmarkEnd bookmarkEnd, HtmlTextWriter sb)
    {
    }

    internal override void ProcessBreak(Break @break, HtmlTextWriter writer)
    {
        if (@break.Type != null && @break.Type == BreakValues.Page)
        {
            writer.WriteStartElement("div");
            writer.WriteAttributeString("style", "break-after: page;");
            writer.WriteEndElement();
        }
        else if (@break.Type != null && @break.Type == BreakValues.Column)
        {
            writer.WriteStartElement("div");
            writer.WriteAttributeString("style", "break-after: column;");
            writer.WriteEndElement();
        }
        else // Type == TextWrapping or not specified
        {
            writer.WriteBreak();
        }
    }

    internal override void ProcessHyperlink(Hyperlink hyperlink, HtmlTextWriter sb)
    {
        bool hasUrl = false;
        if (hyperlink.Id?.Value is string rId)
        {
            var maindDocumentPart = OpenXmlHelpers.GetMainDocumentPart(hyperlink);
            if (maindDocumentPart?.HyperlinkRelationships.FirstOrDefault(x => x.Id == rId) is HyperlinkRelationship relationship)
            {
                string url = relationship.Uri.OriginalString.Replace(" ", "%20");
                hasUrl = true;
                sb.WriteStartElement("a");
                sb.WriteAttributeString("href", url);
            }
        }
        else if (hyperlink.Anchor?.Value is string anchor)
        {
            hasUrl = true;
            sb.WriteStartElement("a");
            sb.WriteAttributeString("href", $"#{anchor}");
        }

        foreach (var element in hyperlink.Elements())
        {
            base.ProcessParagraphElement(element, sb);
        }

        if (hasUrl)
        {
            sb.WriteEndElement("a");
        }
    }

    internal void ProcessShading(Shading? shading, ref List<string> styles)
    {
        if (ColorHelpers.EnsureHexColor(shading?.Fill?.Value) is string fill)
        {
            styles.Add($"background-color: #{fill};");
            // Not supported: foreground (pattern)
        }
    }

    internal override void ProcessRuby(Ruby ruby, HtmlTextWriter writer)
    {
        writer.WriteStartElement("ruby");
        base.ProcessRuby(ruby, writer); // Processes RubyBase
        if (ruby.RubyContent != null)
        {
            writer.WriteStartElement("rt");
            // TODO: <rp> element for browsers that don't support Ruby
            foreach (var element in ruby.RubyContent.Elements())
            {
                ProcessRubyElement(element, writer);
            }
            writer.WriteEndElement("rt");
        }
        writer.WriteEndElement("ruby");
        // if (ruby.RubyProperties != null)
        // {
        // }
    }

    internal override void ProcessPageNumber(PageNumber pageNumber, HtmlTextWriter sb)
    {
    }

    internal override void ProcessPositionalTab(PositionalTab posTab, HtmlTextWriter sb)
    {
    }

    internal override void ProcessFieldChar(FieldChar field, HtmlTextWriter sb)
    {
    }

    internal override void ProcessFieldCode(FieldCode field, HtmlTextWriter sb)
    {
    }

    internal override void ProcessCommentStart(CommentRangeStart commentStart, HtmlTextWriter sb) { }
    internal override void ProcessCommentEnd(CommentRangeEnd commentEnd, HtmlTextWriter sb) { }
    internal override void ProcessAnnotationReference(AnnotationReferenceMark annotationRef, HtmlTextWriter sb) { }
    internal override void ProcessCommentReference(CommentReference commentRef, HtmlTextWriter sb) { }
}
