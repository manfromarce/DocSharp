using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Xml.Linq;
using DocSharp.Xml;
using AngleSharp;
using AngleSharp.Dom;
using AngleSharp.Html.Parser;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace DocSharp.Ebook;

internal class HtmlUtils
{
    public static string NormalizeHtml(string htmlContent, string tempDir, List<string> chapterFileNames)
    {
        HtmlParser parser = new HtmlParser();
        IDocument document;
        try
        {
            document = parser.ParseDocument(htmlContent);
        }
        catch
        {
            // If the document cannot be parsed, return the original content unchanged.
            return htmlContent;
        }

        if (document == null || document.DocumentElement == null)
            return htmlContent;

        // Search for links (<a> tags)
        foreach (var link in document.QuerySelectorAll("a"))
        {
            var href = link.GetAttribute("href");
            if (!string.IsNullOrEmpty(href))
            {
                link.SetAttribute("href", FixLink(href, tempDir, chapterFileNames));
            }
        }

        // Search for images (<img> tags)
        foreach (var img in document.QuerySelectorAll("img"))
        {
            var src = img.GetAttribute("src");
            if (!string.IsNullOrEmpty(src))
            {
                img.SetAttribute("src", FixImageSource(src, tempDir));
            }
        }

        // Return the normalized HTML. Use the full document element for compatibility.
        return document.DocumentElement.OuterHtml;
    }

    public static string FixImageSource(string link, string tempDir)
    {
        if (Uri.TryCreate(link, UriKind.RelativeOrAbsolute, out Uri uri))
        {
            if (uri.IsAbsoluteUri)
            {
                // The URI is already absolute
                if (uri.IsFile) // Absolute file path (not valid, most likely not found)
                    return string.Empty;
                else // Online image source (https, ftp, ...), return unchanged
                    return uri.AbsoluteUri;
            }
            else 
            {
                // The URI is relative, combine it with the base path.
                string absolute =  Path.GetFullPath(Path.Combine(tempDir, link));
                // Convert the absolute file path to a file:/// URL
                if (Uri.TryCreate(absolute, UriKind.Absolute, out Uri absoluteFileUri))
                    return absoluteFileUri.AbsoluteUri;
                else // An invalid combined URI was produced
                    return string.Empty;
            }
        }
        else // The URI is not valid.
            return string.Empty;
    }

    public static string FixLink(string link, string tempDir, List<string> chapterFileNames)
    {
        if (Uri.TryCreate(link, UriKind.RelativeOrAbsolute, out Uri uri))
        {
            if (uri.IsAbsoluteUri)
            {
                // The URI is already absolute
                if (uri.IsFile) // Absolute file path (not valid, most likely not found)
                    return string.Empty;
                else // Online URL (https, mailto, ...), return unchanged
                    return uri.AbsoluteUri;
            }
            else 
            {                
                if (link.Contains('#')) // The URL is or contains an anchor
                {                    
                    // Anchors will be preserved in the final document, 
                    // but remove the URL (if any) because everything will in one file.
                    string anchor = link.Substring(link.LastIndexOf('#'));                
                    return anchor; // HtmlToOpenXml wil convert the anchor to a DOCX bookmark.
                }
                else
                {                    
                    // The URI is relative and will not work in the final document.
                    // If it points to a chapter, replace it with an anchor (will be created later).
                    
                    // Get file name after the last slash or reverse slash (if any)
                    string fileName = Path.GetFileName(link).ToLower();
                    if (chapterFileNames.Contains(fileName))
                    {
                        string anchor = $"#_{fileName.Replace(" ", "_")}";
                        return anchor;                  
                    }
                    else
                    {
                        // The URI is not valid.
                        return string.Empty;
                    }
                }

            }
        }
        else // The URI is not valid.
        {
            if (link.Contains('#')) // The URL is or contains an anchor
            {
                // Anchors will be preserved in the final document, 
                // but remove the URL (if any) because everything will in one file.
                string anchor = link.Substring(link.LastIndexOf('#'));                
                return anchor; // HtmlToOpenXml wil convert the anchor to a DOCX bookmark.
            }
            else
            {
                // Unrecognized link type
                return string.Empty;
            }
        }
    }
}
