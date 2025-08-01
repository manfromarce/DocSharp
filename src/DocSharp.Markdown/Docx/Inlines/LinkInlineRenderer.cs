using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Globalization;
using System.IO;
using System.Linq;
using DocSharp.Docx;
using DocSharp.IO;
using DocSharp.Markdown.Common;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using Markdig.Renderers.Html;
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
                LinkImageRenderHelper.GetImageAttributes(obj, out long widthInTwips, out long heightInTwips);
                ProcessImage(renderer, obj.Url!, obj.Label, obj.Title, widthInTwips, heightInTwips);
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
            
            var linkId = $"L{_hyperlinkIdCounter++}";
            Debug.Assert(renderer.Document.MainDocumentPart != null, "Document.MainDocumentPart != null");

            Hyperlink hl;
            if (isAnchor)
            {
                // Link to bookmark
                hl = new Hyperlink()
                {
                    Anchor = anchorName
                };
            }
            else
            {
                renderer.Document.MainDocumentPart.AddHyperlinkRelationship(uri!, true, linkId);
                hl = new Hyperlink()
                {
                    Id = linkId,
                };
            }

            renderer.Cursor.Write(hl);
            renderer.Cursor.GoInto(hl);
            
            renderer.TextStyle.Push(renderer.Styles.MarkdownStyles["Hyperlink"]);
            renderer.WriteChildren(obj);
            renderer.TextStyle.Pop();
            
            renderer.Cursor.PopAndAdvanceAfter(hl);
        }
    }

    private void ProcessImage(DocxDocumentRenderer renderer, string url, string? label, string? title, long widthInTwips, long heightInTwips)
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
                        InsertImage(renderer, fs, label, title, widthInTwips, heightInTwips);
                    }
                }
                else if (uri.Scheme == Uri.UriSchemeHttp || uri.Scheme == Uri.UriSchemeHttps)
                {
                    // Only HTTP and HTTPS are supported for automatic download
                    using (var stream = ResourceDownloader.GetDownloadStream(url))
                    {
                        if (stream != null)
                        {
                            InsertImage(renderer, stream, label, title, widthInTwips, heightInTwips);
                        }
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

    private void InsertImage(DocxDocumentRenderer renderer, Stream stream, string? label, string? title, long widthInTwips, long heightInTwips)
    {
        try
        {
            using (var tempStream = LinkImageRenderHelper.ConvertAndScaleImage(stream,
                                                           out ImageFormat fileType,
                                                           renderer.Document.GetEffectivePageSize(), 
                                                           widthInTwips, heightInTwips,
                                                           out long calculatedWidth, out long calculatedHeight,
                                                           false, renderer.ImageConverter))
            {
                PartTypeInfo? imagePartType = fileType.ToImagePartType();
                if (tempStream != null && imagePartType != null && imagePartType.HasValue && calculatedWidth > 0 && calculatedHeight > 0)
                {
                    tempStream.Position = 0;
                    var imagePart = renderer.Document.MainDocumentPart?.AddImagePart(tempStream, imagePartType.Value);
                    if (imagePart != null && renderer.Document.MainDocumentPart?.GetIdOfPart(imagePart) is string rId)
                    {
                        var imageElement = ImageHelpers.CreateImage(rId, calculatedWidth, calculatedHeight, _imageIdCounter, label, title);
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
        }
        catch (Exception ex)
        {
#if DEBUG
            Debug.WriteLine($"Exception in InsertImage: {ex.Message}");
#endif
            return;
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
