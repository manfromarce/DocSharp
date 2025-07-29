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

namespace DocSharp.Markdown.Common;

public static class LinkImageRenderHelper
{
    public static void GetImageAttributes(LinkInline obj, out double width, out double height)
    {
        width = -1;
        height = -1;
        var attrs = obj.GetAttributes();
        if (attrs.Properties != null)
        {
            var attrWidth = attrs.Properties.FirstOrDefault(kvp => kvp.Key.Equals("width", StringComparison.OrdinalIgnoreCase));
            var attrHeight = attrs.Properties.FirstOrDefault(kvp => kvp.Key.Equals("height", StringComparison.OrdinalIgnoreCase));
            if (!attrWidth.Equals(default(KeyValuePair<string, string?>)) &&
                !attrHeight.Equals(default(KeyValuePair<string, string?>)) &&
                double.TryParse(attrWidth.Value, NumberStyles.Number, CultureInfo.InvariantCulture, out double w) &&
                double.TryParse(attrHeight.Value, NumberStyles.Number, CultureInfo.InvariantCulture, out double h))
            {
                width = w;
                height = h;
            }
        }
    }

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
                string normalizedBaseUri = imagesBaseUri.TrimEnd('\\', '/') + @"/";
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
                string normalizedBaseUri = linksBaseUri.TrimEnd('\\', '/') + @"/";
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

    // Scale image dimensions to be less than max width and height, keeping aspect ratio.
    public static void ScaleImageSize(ref long width, ref long height, double maxWidth, double maxHeight)
    {
        if (width > maxWidth || height > maxHeight)
        {
            var ratioX = ((decimal)maxWidth) / width; // use decimal as double may not be able to contain a long
            var ratioY = ((decimal)maxHeight) / height;
            var ratio = Math.Min(ratioX, ratioY);
            width = (long)(width * ratio);
            height = (long)(height * ratio);
        }
    }

    // Returns the original stream or a second stream with the image converted to PNG
    internal static Stream? ConvertAndScaleImage(Stream imageData, out ImageFormat fileType,
                                            Size pageSize, UnitMetric maxSizeUnit,
                                            double desiredWidth, double desiredHeight, UnitMetric desiredSizeUnit,
                                            out long calculatedWidth, out long calculatedHeight,
                                            bool isRtf, IImageConverter? imageConverter)
    {
        Stream? outputStream = null;
        fileType = ImageFormat.Unknown;
        if (ImageHeader.TryDetectFileType(imageData, out fileType))
        {
            if ((isRtf && fileType.IsSupportedInRtf()) || (!isRtf && fileType.IsSupportedInOpenXml()))
            {
                if (imageData.CanSeek)
                {
                    outputStream = imageData;
                }
                else
                {
                    // Fixes issues with network streams and other non-seekable stream.
                    outputStream = new MemoryStream();
                    imageData.CopyTo(outputStream);
                }
            }
            else
            {
                outputStream = new MemoryStream();
                if (!(imageConverter != null && imageConverter.ConvertToPng(imageData, outputStream, fileType)))
                {
#if DEBUG
                    Debug.WriteLine($"Error in ConvertAndScaleImage - conversion failed.");
#endif
                    calculatedWidth = 0;
                    calculatedHeight = 0;
                    return null;
                }
            }
            outputStream.Position = 0;
            if (desiredWidth > 0 && desiredHeight > 0)
            {
                calculatedWidth = isRtf ? UnitMetricHelper.ConvertToTwips(desiredWidth, desiredSizeUnit) : UnitMetricHelper.ConvertToEmus(desiredWidth, desiredSizeUnit);
                calculatedHeight = isRtf ? UnitMetricHelper.ConvertToTwips(desiredHeight, desiredSizeUnit) : UnitMetricHelper.ConvertToEmus(desiredHeight, desiredSizeUnit);
            }
            else
            {
                var maxW = isRtf ? UnitMetricHelper.ConvertToTwips(pageSize.Width, maxSizeUnit) : UnitMetricHelper.ConvertToEmus(pageSize.Width, maxSizeUnit);
                var maxH = isRtf ? UnitMetricHelper.ConvertToTwips(pageSize.Height, maxSizeUnit) : UnitMetricHelper.ConvertToEmus(pageSize.Height, maxSizeUnit);

                // Assume 75% of the page size as maximum (empirical, if 100% is used the image would be moved to the next page too often). 
                maxW = (long)(maxW * 0.75m); // use decimal as double may not be able to contain a long
                maxH = (long)(maxH * 0.75m);

                var originalSize = ImageHeader.GetDimensions(outputStream, fileType);
                outputStream.Position = 0;

                if (originalSize == Size.Empty || originalSize.Width < 0 || originalSize.Height < 0)
                {
#if DEBUG
                    Debug.WriteLine($"Error in ConvertAndScaleImage - empty size returned.");
#endif
                    calculatedWidth = 0;
                    calculatedHeight = 0;
                    return null;
                }

                // GetDimensions returns width and height in pixels, except for WMF
                // whose dimensions are returned in inches as it's not device-independent.
                var unit = fileType == ImageFormat.Wmf ? DocSharp.UnitMetric.Inch : DocSharp.UnitMetric.Pixel;

                calculatedWidth = isRtf ? UnitMetricHelper.ConvertToTwips(originalSize.Width, unit) : UnitMetricHelper.ConvertToEmus(originalSize.Width, unit);
                calculatedHeight = isRtf ? UnitMetricHelper.ConvertToTwips(originalSize.Height, unit) : UnitMetricHelper.ConvertToEmus(originalSize.Height, unit);

                ScaleImageSize(ref calculatedWidth, ref calculatedHeight, maxW, maxH);
            }
            return outputStream;
        }
        else
        {
#if DEBUG
            Debug.WriteLine($"Error in ConvertAndScaleImage - image type not recognized.");
#endif
            calculatedWidth = 0;
            calculatedHeight = 0;
            return null;
        }
    }
}