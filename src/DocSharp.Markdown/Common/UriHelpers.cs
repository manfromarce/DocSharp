using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.Linq;
using DocSharp.IO;
using Markdig.Renderers.Html;
using Markdig.Syntax.Inlines;
#if NETFRAMEWORK
using DocSharp.Helpers;
#endif

namespace DocSharp.Markdown.Common;

internal static class UriHelpers
{
    public static Uri? NormalizeImageUri(string url, string? imagesBaseUri)
    {
        if (string.IsNullOrWhiteSpace(url))
        {
            return null;
        }

        Uri? uri = null;
        var isAbsoluteUri = Uri.TryCreate(url.Trim('"'), UriKind.Absolute, out uri);
        // var isAbsoluteUri = Uri.TryCreate(url, UriKind.Absolute, out uri);
        if (!isAbsoluteUri)
        {
            // The URL is relative or invalid.
            if (Uri.TryCreate(url, UriKind.Relative, out uri) && !string.IsNullOrEmpty(imagesBaseUri))
            {
                // Relative URI is well formatted, check ImagesBaseUri and add a final slash,
                // otherwise it is interpreted as file and relative links starting with . or .. won't work properly.
                // Note that ImagesPathUri should obviously not be a file path.
                string normalizedBaseUri = imagesBaseUri!.TrimEnd('\\', '/') + @"/";
                if (Uri.TryCreate(normalizedBaseUri, UriKind.Absolute, out Uri? baseUri) && baseUri != null)
                {
                    uri = new Uri(baseUri, uri);
                }
            }
        }
        return uri;
    }

    public static Uri? NormalizeLinkUri(string url, string? linksBaseUri)
    {
        if (string.IsNullOrWhiteSpace(url))
        {
            return null;
        }

        Uri? uri = null;

        var isAbsolute = Uri.TryCreate(url.Trim('"'), UriKind.Absolute, out uri);
        // var isAbsolute = Uri.TryCreate(url, UriKind.Absolute, out uri);
        if (!isAbsolute)
        {
            // The URL is relative or invalid.
            if (string.IsNullOrEmpty(linksBaseUri))
            {
                // If LinksBaseUri is null or empty, keep the URI as relative.
                string fixedUrl;
                if (!url.StartsWith('.'))
                {
                    // Relative URIs like "file.txt" or "/file.txt" need be changed to
                    // "./file.txt" to work in Microsoft Word, otherwise the document will result corrupted.
                    fixedUrl = "./" + url.TrimStart(['/', '\\']);
                }
                else
                {
                    fixedUrl = url;
                }
                if (!string.IsNullOrWhiteSpace(fixedUrl))
                {
                    Uri.TryCreate(fixedUrl, UriKind.Relative, out uri);
                }
            }
            else if (Uri.TryCreate(url, UriKind.Relative, out uri))
            {
                // If LinksBaseUri is NOT null/empty and url is a valid relative URI, 
                // combine them to make an absolute URI.
                // Add a final slash to the base URI first, otherwise it is interpreted as file and won't work properly.
                string normalizedBaseUri = linksBaseUri!.TrimEnd('\\', '/') + @"/";
                if (Uri.TryCreate(normalizedBaseUri, UriKind.Absolute, out Uri? baseUri) && baseUri != null)
                {
                    uri = new Uri(baseUri, uri);
                    isAbsolute = true;
                }
            }
        }

        if (uri != null && ((!isAbsolute) || uri.IsFile))
        {
            // Remove anchor (if any) from local file URLs, as it does not work in Word.
            // Note that IsFile throws exception for relative URIs, it should be called only if URI is absolute;
            // and online URIs cannot be relative.
            int i = uri.OriginalString.LastIndexOf('#');
            if (i >= 0)
            {
                string url2 = uri.OriginalString.Substring(0, i);
                uri = new Uri(url2, isAbsolute ? UriKind.Absolute : UriKind.Relative);
            }
        }
        return uri;
    }   
}