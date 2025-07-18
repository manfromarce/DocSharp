using System;
using System.Diagnostics;
using System.IO;
using System.Linq;
using DocSharp.Docx;
using DocSharp.Helpers;
using DocSharp.IO;
using Markdig.Syntax.Inlines;

namespace Markdig.Renderers.Rtf.Inlines;

public class LinkInlineRenderer : RtfObjectRenderer<LinkInline>
{
    protected override void WriteObject(RtfRenderer renderer, LinkInline obj)
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

        renderer.RtfWriter.Write(@"{\field{\*\fldinst{HYPERLINK ");

        if (isExternal)
        {
            renderer.RtfWriter.Write(@"""" + uri + @"""}}");
        }
        else
        {
            // Link to bookmark
            renderer.RtfWriter.Write(@"\\l """ + anchorName + @"""}}");
        }
        renderer.RtfWriter.Write(@"{\fldrslt{\cf17\ul ");
        renderer.WriteChildren(obj);
        renderer.RtfWriter.Write(@"}}}");
    }

    private void ProcessImage(RtfRenderer renderer, string url, string? label, string? title)
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
                string normalizedBaseUri = renderer.ImagesBaseUri.TrimEnd('\\', '/') + @"/";
                if (Uri.TryCreate(normalizedBaseUri, UriKind.Absolute, out Uri? baseUri)
                    && baseUri != null)
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
                //Probably non-existent file or permission issue, do not stop the conversion
                #if DEBUG
                Console.WriteLine("InsertImage exception: " + ex.Message);
                #endif
            }
        }
    }

    private void InsertImage(RtfRenderer renderer, string filePath, string? label, string? title)
    {
        using (var fs = new FileStream(filePath, FileMode.Open, FileAccess.Read))
        {
            InsertImage(renderer, fs, label, title);
        }
    }

    private void InsertImage(RtfRenderer renderer, byte[] bytes, string? label, string? title)
    {
        using (var ms = new MemoryStream(bytes))
        {
            InsertImage(renderer, ms, label, title);
        }
    }

    private void InsertImage(RtfRenderer renderer, Stream stream, string? label, string? title)
    {
        if (ImageHeader.TryDetectFileType(stream, out var fileType))
        {
            // Support only JPEG and PNG formats for now
            if (fileType != ImageFormat.Jpeg && fileType != ImageFormat.Png)
            {
                #if DEBUG
                Console.WriteLine($"Unsupported image format: {fileType}");
                #endif
                return;
            }

            // Read bytes
            byte[] imageBytes;
            using (var ms = new MemoryStream())
            {
                stream.CopyTo(ms);
                imageBytes = ms.ToArray();
            }

            // Calculate image dimensions and scale them
            System.Drawing.Size size = ImageHeader.GetDimensions(new MemoryStream(imageBytes), fileType);
            // GetDimensions returns width and height in pixels, except for WMF
            // whose dimensions are returned in inches as it's not device-independent.
            var unit = fileType == ImageFormat.Wmf ? DocSharp.UnitMetric.Inch : DocSharp.UnitMetric.Pixel;
            // Convert to twips
            var width = DocSharp.UnitMetricHelper.ConvertToTwips(size.Width, unit);
            var height = DocSharp.UnitMetricHelper.ConvertToTwips(size.Height, unit);

            // Consider 75% of A4 page size in twips (1/1440 inch)
            long maxWidth = 8929; 
            long maxHeight = 12585;
            ScaleImageSize(ref width, ref height, maxWidth, maxHeight);

            // Write RTF syntax for image
            renderer.RtfWriter.Write(@"{\pict");
            renderer.RtfWriter.Write(fileType == ImageFormat.Jpeg ? @"\jpegblip" : @"\pngblip");
            renderer.RtfWriter.Write($@"\picw{width}\pich{height}\picwgoal{width}\pichgoal{height} ");
            foreach (var b in imageBytes)
            {
                renderer.RtfWriter.Write(b.ToString("x2"));
            }
            renderer.RtfWriter.WriteLine(@"}");
        }
    }

    // Scale image dimensions to be less than max width and height, keeping aspect ratio.
    private static void ScaleImageSize(ref long width, ref long height, long maxWidth, long maxHeight)
    {
        if (width > maxWidth || height > maxHeight)
        {
            var ratioX = (double)maxWidth / width;
            var ratioY = (double)maxHeight / height;
            var ratio = Math.Min(ratioX, ratioY);
            width = Math.Max((int)(width * ratio), 1);
            height = Math.Max((int)(height * ratio), 1);
        }
    }

    private string TryGetBookmark(RtfRenderer renderer, string anchorId)
    {
        var bookmarkName = renderer.Bookmarks
                           .Where(b => b.Equals(anchorId, StringComparison.OrdinalIgnoreCase));
        return bookmarkName?.FirstOrDefault() ?? "";
    }
}
