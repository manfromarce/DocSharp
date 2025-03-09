using System;
using System.Diagnostics;
using System.IO;
using System.Linq;
using DocSharp.Docx;
using DocSharp.IO;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using Markdig.Syntax.Inlines;

namespace Markdig.Renderers.Docx.Inlines;

public class LinkInlineRenderer : DocxObjectRenderer<LinkInline>
{
    private uint _hyperlinkIdCounter = 1;
    private uint _imageIdCounter = 1;

    protected override void WriteObject(DocxDocumentRenderer renderer, LinkInline obj)
    {
        if (string.IsNullOrWhiteSpace(obj.Url))
        {
            renderer.WriteChildren(obj);
            return;
        }

        if (obj.IsImage)
        {
            if (!renderer.SkipImages)
            {
                ProcessImage(renderer, obj.Url, obj.Label, obj.Title);
            }
            return;
        }

        Uri? uri = null;

        var isAbsoluteUri = Uri.TryCreate(obj.Url, UriKind.Absolute, out uri);
        bool isExternal = true;
        string anchorName = string.Empty;
        
        if (!isAbsoluteUri)
        {
            isExternal = !obj.Url.StartsWith("#"); // anchor link
            
            if (isExternal)
            {
                string fixedUrl = string.Empty;
                if (!obj.Url.StartsWith("."))
                {
                    // Relative URIs like "file.txt" or "/file.txt" need be changed to
                    // "./file.txt" to work in Microsoft Word, otherwise the document will result corrupted.
                    fixedUrl = "./" + obj.Url.TrimStart(['/', '\\']);
                }
                else
                {
                    fixedUrl = obj.Url;
                }
                if (!string.IsNullOrWhiteSpace(fixedUrl))
                {
                    Uri.TryCreate(obj.Url, UriKind.Relative, out uri);
                }
            }
            else
            {
                // Note: bookmarks are currently created in HeadingRenderer (for headings).
                // It should be done in HtmlInline/HtmlBlock renderer too (for <a> tags),
                // but it's not implemented yet.
                anchorName = TryGetBookmark(renderer, obj.Url.Trim('#'));
            }
        }

        if (isExternal && uri != null && ((!isAbsoluteUri) || uri.IsFile))
        {
            // Remove anchor (if any) from local file URLs, as it does not work in Word.
            // Note that IsFile throws exception for relative URIs, it should be called only if URI is absolute
            int i = obj.Url.LastIndexOf('#');
            if (i >= 0)
            {
                string url = uri.OriginalString.Substring(0, i);
                uri = new Uri(url, isAbsoluteUri ? UriKind.Absolute : UriKind.Relative);
            }
        }

        if (uri == null && anchorName == string.Empty)
        {
            // Don't create the hyperlink if no valid Uri or anchor was created.
            renderer.WriteChildren(obj);
            return;
        }
                  
        var linkId = $"L{_hyperlinkIdCounter++}";
        Debug.Assert(renderer.Document.MainDocumentPart != null, "Document.MainDocumentPart != null");

        Hyperlink hl;
        if (isExternal)
        {
            renderer.Document.MainDocumentPart.AddHyperlinkRelationship(uri!, isExternal, linkId);
            hl = new Hyperlink()
            {
                Id = linkId,
            };
        }
        else
        {
            // Link to bookmark
            hl = new Hyperlink()
            {
                Anchor = anchorName
            };
        }

        renderer.Cursor.Write(hl);
        renderer.Cursor.GoInto(hl);
            
        renderer.TextStyle.Push(renderer.Styles.MarkdownStyles["Hyperlink"]);
        renderer.WriteChildren(obj);
        renderer.TextStyle.Pop();
            
        renderer.Cursor.PopAndAdvanceAfter(hl);
    }

    private void ProcessImage(DocxDocumentRenderer renderer, string url, string? label, string? title)
    {
        Uri? uri = null;

        var isAbsoluteUri = Uri.TryCreate(url, UriKind.Absolute, out uri);

        if (!isAbsoluteUri)
        {
            // Could be a relative URL
            if (Uri.TryCreate(url, UriKind.Relative, out uri) && !string.IsNullOrEmpty(renderer.ImagesBaseUri))
            {
                // Relative URI is well formatted, check ImagesBaseUri and add a final slash,
                // otherwise it is interpreted as file and relative links starting with . or .. won't work properly.
                // Note that ImagesPathUri should not be a file path.
                string normalizedBaseUri = renderer.ImagesBaseUri.Trim('\\', '/') + @"\";
                if (Uri.TryCreate(Path.TrimEndingDirectorySeparator(normalizedBaseUri) + '/', 
                                  UriKind.Absolute, out Uri? baseUri) && baseUri != null)
                {
                    uri = new Uri(baseUri, uri);
                }
            }
        } // else the URI can be directly processed as absolute

        if (uri != null)
        {
            try
            {
                if (uri.IsFile)
                {
                    InsertImage(renderer, uri.LocalPath, label, title);
                }
                else if (uri.Scheme == Uri.UriSchemeHttp || uri.Scheme == Uri.UriSchemeHttps)
                {
                    // Only HTTP and HTTPS are supported for automatic download
                    var bytes = ResourceDownloader.DownloadFile(url);
                    if (bytes != null)
                    {
                        InsertImage(renderer, bytes, label, title);
                    }
                }
            }
            catch (Exception ex)
            {
                //Probably non-existent file, inaccessible URL or permission issue, do not stop the conversion
                #if DEBUG
                Debug.WriteLine("InsertImage exception: " + ex.Message);
                #endif
            }
        }
    }

    private void InsertImage(DocxDocumentRenderer renderer, string filePath, string? label, string? title)
    {
        using (var fs = new FileStream(filePath, FileMode.Open, FileAccess.Read))
        {
            InsertImage(renderer, fs, label, title);
        }
    }

    private void InsertImage(DocxDocumentRenderer renderer, byte[] bytes, string? label, string? title)
    {
        using (var ms = new MemoryStream(bytes))
        {
            InsertImage(renderer, ms, label, title);
        }
    }

    private void InsertImage(DocxDocumentRenderer renderer, Stream stream, string? label, string? title)
    {
        if (ImageHeader.TryDetectFileType(stream, out var fileType))
        {
            PartTypeInfo imageFormat;
            switch (fileType)
            {
                case ImageHeader.FileType.Bitmap:
                    imageFormat = ImagePartType.Bmp;
                    break;
                case ImageHeader.FileType.Gif:
                    imageFormat = ImagePartType.Gif;
                    break;
                case ImageHeader.FileType.Jpeg:
                    imageFormat = ImagePartType.Jpeg;
                    break;
                case ImageHeader.FileType.Png:
                    imageFormat = ImagePartType.Png;
                    break;
                case ImageHeader.FileType.Tiff:
                    imageFormat = ImagePartType.Tiff;
                    break;
                case ImageHeader.FileType.Svg:
                    imageFormat = ImagePartType.Svg;
                    break;
                case ImageHeader.FileType.Ico:
                    imageFormat = ImagePartType.Icon;
                    break;
                default:
                    // Note: WEBP and AVIF images are supported by web browsers but not by DOCX.
                    return;
            }

            var imagePart = renderer.Document.MainDocumentPart?.AddImagePart(stream, imageFormat);
            if (imagePart != null && renderer.Document.MainDocumentPart?.GetIdOfPart(imagePart) is string rId)
            {
                System.Drawing.Size size = System.Drawing.Size.Empty;
                try
                {
                    size = ImageHeader.GetDimensions(stream, fileType);
                }
                catch
                {
                    return;
                }
                if (size == System.Drawing.Size.Empty || size.Width < 0 || size.Height < 0)
                {
                    return;
                }
                // GetDimensions returns width and height in pixels, except for WMF
                // whose dimensions are return in inches as it's not device-independent.
                var unit = fileType == ImageHeader.FileType.Wmf ? DocSharp.UnitMetric.Inch : DocSharp.UnitMetric.Pixel;
                
                // Convert to EMUs
                var width = DocSharp.UnitMetricHelper.ConvertToEmus(size.Width, unit);
                var height = DocSharp.UnitMetricHelper.ConvertToEmus(size.Height, unit);

                // Try to scale image size to fit the page.
                var pageSize = renderer.Document.GetEffectivePageSize();
                // Convert twips to EMUs (used for image dimensions),
                // and assume 75% of the page size as maximum (empirical). 
                var maxWidth = pageSize.Width * 635 * 0.75;
                var maxHeight = pageSize.Height * 635 * 0.75;

                ScaleImageSize(ref width, ref height, maxWidth, maxHeight);

                var imageElement = ImageHelpers.CreateImage(rId, width, height, _imageIdCounter, label, title);
                ++_imageIdCounter;
                if (renderer.Cursor.Container is Run run)
                {
                    renderer.Cursor.Write(imageElement);
                }
                else
                {
                    renderer.Cursor.Write(new Run(imageElement));
                }
            }
        }
    }

    // Scale image dimensions to be less than max width and height, keeping aspect ratio.
    private static void ScaleImageSize(ref long width, ref long height, double maxWidth, double maxHeight)
    {
        if (width > maxWidth || height > maxHeight)
        {
            var ratioX = maxWidth / width;
            var ratioY = maxHeight / height;
            var ratio = Math.Min(ratioX, ratioY);
            width = Math.Max((int)(width * ratio), 1);
            height = Math.Max((int)(height * ratio), 1);
        }
    }

    private string TryGetBookmark(DocxDocumentRenderer renderer, string anchorId)
    {
        // To be improved
        var bookmarkName = renderer.Document.MainDocumentPart?.Document.Body?
                           .Descendants<BookmarkStart>()
                           .Select(bs => bs.Name)
                           .Where(name => name != null && 
                                          name.Value != null && 
                                          name.Value.Equals(anchorId, StringComparison.OrdinalIgnoreCase));
        return bookmarkName?.FirstOrDefault()?.Value ?? "";
    }
}
