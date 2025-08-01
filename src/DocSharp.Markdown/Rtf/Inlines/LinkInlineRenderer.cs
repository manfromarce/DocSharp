using System;
using System.Diagnostics;
using System.IO;
using System.Linq;
using DocSharp.Docx;
using DocSharp.Helpers;
using DocSharp.IO;
using DocSharp.Markdown.Common;
using Markdig.Renderers.Html;
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
                long width, height;
                LinkImageRenderHelper.GetImageAttributes(obj, out width, out height);
                ProcessImage(renderer, obj.Url!, obj.Label, obj.Title, width, height);
            }
        }
        else
        {
            bool isAnchor = obj.Url!.StartsWith("#");
            string anchorName = string.Empty;
            Uri? uri = null;
            if (isAnchor)
            {
                // Note: bookmarks are currently created in HeadingRenderer (for headings).
                // It should be done in HtmlInline/HtmlBlock renderer too (for <a> tags),
                // but it's not implemented yet.
                anchorName = TryGetBookmark(renderer, obj.Url.Trim('#'));
            }
            else
            {
                uri = LinkImageRenderHelper.NormalizeLinkUri(obj.Url, renderer.LinksBaseUri);
            }

            if (uri == null && anchorName == string.Empty)
            {
                // Don't create the hyperlink if no valid Uri or anchor was created.
                renderer.WriteChildren(obj);
                return;
            }

            renderer.RtfWriter.Write(@"{\field{\*\fldinst{HYPERLINK ");

            if (isAnchor)
            {
                // Link to bookmark
                renderer.RtfWriter.Write(@"\\l """ + anchorName + @"""}}");
            }
            else
            {
                renderer.RtfWriter.Write(@"""" + uri + @"""}}");
            }
            renderer.RtfWriter.Write(@"{\fldrslt{\cf17\ul ");
            renderer.WriteChildren(obj);
            renderer.RtfWriter.Write(@"}}}");
        }
    }

    private void ProcessImage(RtfRenderer renderer, string url, string? label, string? title, long width, long height)
    {
        Uri? uri = LinkImageRenderHelper.NormalizeImageUri(url, renderer.ImagesBaseUri);
        if (uri != null)
        {
            try
            {
                if (uri.IsAbsoluteUri && uri.IsFile)
                {
                    using (var fs = new FileStream(uri.LocalPath, FileMode.Open, FileAccess.Read))
                    {
                        InsertImage(renderer, fs, label, title, width, height);
                    }
                }
                else if (uri.Scheme == Uri.UriSchemeHttp || uri.Scheme == Uri.UriSchemeHttps)
                {
                    // Only HTTP and HTTPS are supported for automatic download
                    using (var stream = ResourceDownloader.GetDownloadStream(url))
                    {
                        if (stream != null)
                        {
                            InsertImage(renderer, stream, label, title, width, height);
                        }
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

    private void InsertImage(RtfRenderer renderer, Stream stream, string? label, string? title, long desiredWidth, long desiredHeight)
    {
        try
        {
            var pageSize = new System.Drawing.Size(9026, 13958); // A4 page size - margins (in twips)
            using (var tempStream = LinkImageRenderHelper.ConvertAndScaleImage(stream,
                                                           out ImageFormat fileType,
                                                           pageSize,
                                                           desiredWidth, desiredHeight,
                                                           out long calculatedWidth, out long calculatedHeight,
                                                           true, renderer.ImageConverter))
            {
                if (calculatedWidth > 0 && calculatedHeight > 0 && tempStream != null)
                {
                    // Write RTF syntax for image
                    renderer.RtfWriter.Write(@"{\pict");
                    renderer.RtfWriter.Write(fileType == ImageFormat.Jpeg ? @"\jpegblip" : @"\pngblip");
                    renderer.RtfWriter.Write($@"\picw{calculatedWidth.ToStringInvariant()}\pich{calculatedHeight.ToStringInvariant()}\picwgoal{calculatedWidth.ToStringInvariant()}\pichgoal{calculatedHeight.ToStringInvariant()} ");
                    int b;
                    while ((b = tempStream.ReadByte()) != -1)
                    {
                        renderer.RtfWriter.Write(b.ToString("x2"));
                    }
                    renderer.RtfWriter.WriteLine(@"}");
                }
            }
        }
        catch (Exception ex)
        {
#if DEBUG
            Debug.WriteLine($"Exception in InsertImage: {ex.Message}");
#endif
            return;
        }
    }

    private string TryGetBookmark(RtfRenderer renderer, string anchorId)
    {
        var bookmarkName = renderer.Bookmarks
                           .Where(b => b.Equals(anchorId, StringComparison.OrdinalIgnoreCase));
        return bookmarkName?.FirstOrDefault() ?? "";
    }
}
