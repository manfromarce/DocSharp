using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DocSharp.Helpers;
using DocSharp.IO;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Office2019.Drawing.SVG;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Vml;
using DocumentFormat.OpenXml.Wordprocessing;
using DrawingML = DocumentFormat.OpenXml.Drawing;
using Path = System.IO.Path;
using M = DocumentFormat.OpenXml.Math;
using W14 = DocumentFormat.OpenXml.Office2010.Word;
using Wp = DocumentFormat.OpenXml.Drawing.Wordprocessing;
using A = DocumentFormat.OpenXml.Drawing;
using Pic = DocumentFormat.OpenXml.Drawing.Pictures;
using Wps = DocumentFormat.OpenXml.Office2010.Word.DrawingShape;
using V = DocumentFormat.OpenXml.Vml;
using System.Globalization;

namespace DocSharp.Docx;

public partial class DocxToHtmlConverter : DocxToTextConverterBase
{
    // Experimental - produce HTML with a fixed layout to preserve page setup (size, margins, ...) and page breaks.
    public bool FixedLayout { get; set; } = false;

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

    internal override void ProcessDocument(Document document, StringBuilder sb)
    {
        sb.AppendLine("<!DOCTYPE html>");
        sb.AppendLine("<html>");
        sb.AppendLine("<head><meta charset=\"utf-8\" /></head>");
        sb.AppendLine("<body>");
        if (document.DocumentBackground is DocumentBackground bg)
        {
            ProcessDocumentBackground(bg, sb);
        }

        // Process body content
        if (document.Body is Body body)
        {
            base.ProcessBody(body, sb);
        }               
        
        sb.AppendLine("</body>");
        sb.Append("</html>");
    }

    internal override void ProcessDocumentBackground(DocumentBackground background, StringBuilder sb)
    {
        if (background.Color != null)
        {
            string color = $"#{background.Color.Value}";
            sb.Append($"<style>body {{ background-color: {color}; }}</style>");
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

    internal override void ProcessContentPart(ContentPart contentPart, StringBuilder sb)
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
                    if(part.ContentType == "application/mathml+xml" ||
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

    internal override void ProcessPositionalTab(PositionalTab posTab, StringBuilder sb)
    {
    }

    internal override void ProcessFieldChar(FieldChar field, StringBuilder sb)
    {
    }

    internal override void ProcessFieldCode(FieldCode field, StringBuilder sb)
    {
    }
}
