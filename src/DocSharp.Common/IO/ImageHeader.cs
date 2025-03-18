/* Copyright (C) Olivier Nizet https://github.com/onizet/html2openxml - All Rights Reserved
 * 
 * This source is subject to the Microsoft Permissive License.
 * Please see the License.txt file for more information.
 * All other rights reserved.
 * 
 * THIS CODE AND INFORMATION ARE PROVIDED "AS IS" WITHOUT WARRANTY OF ANY 
 * KIND, EITHER EXPRESSED OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE
 * IMPLIED WARRANTIES OF MERCHANTABILITY AND/OR FITNESS FOR A
 * PARTICULAR PURPOSE.
 *
 * Original source code from Andy Wilson: http://www.codeproject.com/KB/cs/ReadingImageHeaders.aspx
 * http://stackoverflow.com/questions/111345/getting-image-dimensions-without-reading-the-entire-file/111349
 * EMF Specifications: https://docs.microsoft.com/en-us/openspecs/windows_protocols/ms-emf/ae7e7437-cfe5-485e-84ea-c74b51b000be
 */
using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Xml.XPath;

namespace DocSharp.IO;

/// <summary>
/// Utility class to extract some information of an image file without reading the entire file.
/// </summary>
public static class ImageHeader
{
    // https://en.wikipedia.org/wiki/List_of_file_signatures

    private static readonly byte[] pngSignatureBytes = [0x89, 0x50, 0x4E, 0x47, 0x0D, 0x0A, 0x1A, 0x0A];

    private static readonly Dictionary<byte[], ImageFormat> imageFormatDecoders = new()
    {
        { new byte[] { 0x42, 0x4D }, ImageFormat.Bitmap }, // bmp, dib
        { Encoding.UTF8.GetBytes("GIF87a"), ImageFormat.Gif },
        { Encoding.UTF8.GetBytes("GIF89a"), ImageFormat.Gif }, // animated gif
        { pngSignatureBytes, ImageFormat.Png },
        { new byte[] { 0xff, 0xd8, 0xff }, ImageFormat.Jpeg }, // JPEG
        { new byte[] { 0x49, 0x49, 0xbc, 0x01 }, ImageFormat.Jxr }, // JPEG-XR / HD Photo (jxr, wdp, hdp, wmp) - Not a TIFF
        { new byte[] { 0x49, 0x49, 0xbc, 0x00 }, ImageFormat.Jxr }, // HD Photo pre-release
        { new byte[] { 0x49, 0x49, 0x2a, 0x00 }, ImageFormat.Tiff }, // TIFF (little-endian)
        { new byte[] { 0x4d, 0x4d, 0x00, 0x2a }, ImageFormat.Tiff }, // TIFF (big-endian)
        { new byte[] { 0x49, 0x49, 0x2b, 0x00 }, ImageFormat.Tiff }, // BigTIFF (little-endian)
        { new byte[] { 0x4d, 0x4d, 0x00, 0x2b }, ImageFormat.Tiff }, // BigTIFF (big-endian)
        { new byte[] { 0x52, 0x49, 0x46, 0x46 }, ImageFormat.Webp }, // WebP or other RIFF-based format (needs further analysis)
        { new byte[] { 0x00, 0x00, 0x00, 0x0C, 0x4A, 0x58, 0x4C, 0x20, 0x0D, 0x0A, 0x87, 0x0A }, ImageFormat.Jxl },
        { new byte[] { 0xFF, 0x0A }, ImageFormat.Jxl },
        { new byte[] { 0x00, 0x00, 0x00, 0x0C, 0x6A, 0x50, 0x20, 0x20, 0x0D, 0x0A, 0x87, 0x0A }, ImageFormat.Jpeg2000 }, // JPEG 2000 (jp2, jpf, jpx, jpm, mj2, jph)
        { new byte[] { 0xFF, 0x4F, 0xFF, 0x51 }, ImageFormat.Jpeg2000 }, // JPEG 2000 codestream (j2k, j2c, jpc)
        { new byte[] { 0xD7, 0xCD, 0xC6, 0x9A }, ImageFormat.Wmf },
        { new byte[] { 0x01, 0x00, 0x09, 0x00 }, ImageFormat.Wmf },
        { new byte[] { 0x02, 0x00, 0x09, 0x00 }, ImageFormat.Wmf },
        { new byte[] { 0x1, 0, 0, 0 }, ImageFormat.Emf },
        { new byte[] { 0, 0, 0x1, 0 }, ImageFormat.Ico }, 
        { new byte[] { 0, 0, 0x2, 0 }, ImageFormat.Cur },
        { new byte[] { 0x0A }, ImageFormat.Pcx }, // needs further analysis
        
        // Signatures for AVIF and HEIF start from the 5th byte (the first 4 can be skipped)
        { Encoding.UTF8.GetBytes("ftypavif"), ImageFormat.Avif },
        { Encoding.UTF8.GetBytes("ftypavis"), ImageFormat.Avif }, // Avif image sequence
        { Encoding.UTF8.GetBytes("ftypmif1"), ImageFormat.Avif },
        { Encoding.UTF8.GetBytes("ftypheic"), ImageFormat.Heif },
        { Encoding.UTF8.GetBytes("ftypheim"), ImageFormat.Heif },
        { Encoding.UTF8.GetBytes("ftypheis"), ImageFormat.Heif },
        { Encoding.UTF8.GetBytes("ftypheix"), ImageFormat.Heif },
        { Encoding.UTF8.GetBytes("ftyphevc"), ImageFormat.Heif },
        { Encoding.UTF8.GetBytes("ftyphevm"), ImageFormat.Heif },
        { Encoding.UTF8.GetBytes("ftyphevs"), ImageFormat.Heif },
    };

    private static readonly int MaxMagicBytesLength = imageFormatDecoders
        .Keys.OrderByDescending(x => x.Length).First().Length;

    /// <summary>
    /// Read a image stream and try to detect its file type.
    /// </summary>
    /// <param name="stream">The readable image stream</param>
    /// <param name="type">The guess file type.</param>
    /// <returns>Returns true if the detection was successful.</returns>
    public static bool TryDetectFileType(Stream stream, out ImageFormat type)
    {
        using (SequentialBinaryReader reader = new SequentialBinaryReader(stream, leaveOpen: true))
        {
            type = DetectFileType(reader);
            stream.Seek(0L, SeekOrigin.Begin);
        }
        if (type == ImageFormat.Unknown)
        {
            // Check for SVG
            if (stream.Length > 4)
            {
                using (var sr = new StreamReader(stream, encoding: Encoding.UTF8, detectEncodingFromByteOrderMarks: true, bufferSize: 1024, leaveOpen: true))
                {
                    char[] buffer = new char[5];
                    sr.Read(buffer, 0, 5);
                    var res = new string(buffer).Trim();
                    if (res.StartsWith("<svg") || res.StartsWith("<?xml"))
                    {
                        type = ImageFormat.Svg;
                    }                    
                }
            }
        }
        stream.Seek(0L, SeekOrigin.Begin);
        return type != ImageFormat.Unknown;
    }

    /// <summary>
    /// Examines the first bytes of the file and estimates its image type if possible.
    /// </summary>
    /// <returns>Returns <see cref="ImageFormat.Unrecognized"/> if not recognized.</returns>
    private static ImageFormat DetectFileType(SequentialBinaryReader reader)
    {
        byte[] magicBytes = new byte[MaxMagicBytesLength];
        for (int i = 0; i < MaxMagicBytesLength; i += 1)
        {
            magicBytes[i] = reader.ReadByte();
            foreach (var kvPair in imageFormatDecoders)
            {
                int startIndex = 0;
                if (kvPair.Value == ImageFormat.Avif || kvPair.Value == ImageFormat.Heif)
                {
                    // Skip first 4 values (0, 1, 2, 3)
                    startIndex = 4;
                }

                if (i >= startIndex && StartsWith(magicBytes, kvPair.Key, startIndex))
                {
                    return kvPair.Value;
                }
            }
        }
        return ImageFormat.Unknown;
    }

    /// <summary>
    /// Gets the dimensions of a supported image in pixels (except for WMF which will be in inches).
    /// This function should be called passing a file type determined using TryDetectFileType,
    /// as it expects a specific file header and file extensions are not guaranteed to match the actual content type.
    /// </summary>
    /// <param name="stream">The image stream.</param>
    /// <param name="type">The image type.</param>
    /// <returns>The dimensions of the specified image.</returns>
    public static Size GetDimensions(Stream stream, ImageFormat type)
    {
        using (SequentialBinaryReader reader = new SequentialBinaryReader(stream, leaveOpen: true))
        {
            stream.Seek(0L, SeekOrigin.Begin);
            switch (type)
            {
                case ImageFormat.Bitmap: return DecodeBitmap(reader);
                case ImageFormat.Gif: return DecodeGif(reader);
                case ImageFormat.Jpeg: return DecodeJfif(reader);
                case ImageFormat.Png: return DecodePng(reader);
                case ImageFormat.Ico: return DecodeIco(reader);
                case ImageFormat.Svg: return DecodeXml(stream);

                // Other formats are not supported by web browsers or DOCX or neither,
                // so reading their dimensions is not needed anywhere in the library.
                // However, the following are handled for the sake of completeness and possible future use.
                case ImageFormat.Webp: return DecodeWebP(reader); // supported by web browsers
                case ImageFormat.Avif: // supported by web browsers
                case ImageFormat.Heif: // supported by Safari
                    return DecodeAvifOrHeif(reader);
                case ImageFormat.Jxr: return DecodeJxr(reader); // supported by IE 11
                case ImageFormat.Tiff: return DecodeTiff(reader); // supported in DOCX and Safari
                case ImageFormat.Jpeg2000: return DecodeJpeg2000(reader); // supported in DOCX and previous Safari versions
                case ImageFormat.Emf: return DecodeEmf(reader); // supported in DOCX and RTF
                case ImageFormat.Wmf: return DecodeWmf(reader); // supported in DOCX and RTF
                // Note: WMF size is in inches, all the others are in pixels.
                
                default: return Size.Empty;
            }
        }
    }  

    /// <summary>
    /// Determines whether the beginning of this byte array instance matches the specified byte array.
    /// If <paramref name="startIndex"/> is > 0, bytes from "thisBytes" are checked starting from a specific index.
    /// </summary>
    /// <returns>Returns true if the first array starts with the bytes of the second array.</returns>
    private static bool StartsWith(byte[] thisBytes, byte[] thatBytes, int startIndex = 0)
    {
        for (int i = 0; i < thatBytes.Length; i += 1)
        {
            if (thisBytes[i + startIndex] != thatBytes[i])
            {
                return false;
            }
        }
        return true;
    }

    private static Size DecodeBitmap(SequentialBinaryReader reader)
    {
        reader.IsBigEndian = false;
        if (reader.ReadUInt16() != 0x4D_42)
        {
            throw new Exception("Invalid BMP signature.");
        }
        reader.Skip(4 + 2 + 2 + 4); // skip past the rest of the file header

        int headerSize = reader.ReadInt32();
        switch (headerSize)
        {
            // https://en.wikipedia.org/wiki/BMP_file_format#DIB_header_(bitmap_information_header)
            case 12: // BITMAPCOREHEADER, OS21XBITMAPHEADER
                return new Size(reader.ReadInt16(),
                                Math.Abs(reader.ReadInt16()));
            case 40: // BITMAPINFOHEADER
            case 52: // BITMAPV2INFOHEADER
            case 64: // OS22XBITMAPHEADER
            case 16: // OS22XBITMAPHEADER
            case 56: // BITMAPV3INFOHEADER
            case 104: // BITMAPV4HEADER
            case 124: // BITMAPV5HEADER
                return new Size(reader.ReadInt32(), 
                                Math.Abs(reader.ReadInt32()));
            default:
                return Size.Empty;
        }
    }

    private static Size DecodeGif(SequentialBinaryReader reader)
    {
        reader.IsBigEndian = false;
        
        // 3 - signature: "GIF"
        if (new string(reader.ReadChars(3)) != "GIF")
        {
            throw new Exception("Invalid GIF signature.");
        }

        // 3 - version: either "87a" or "89a"
        reader.Skip(3);

        int width = reader.ReadInt16();
        int height = reader.ReadInt16();
        return new Size(width, height);
    }

    private static Size DecodeJfif(SequentialBinaryReader reader)
    {
        reader.IsBigEndian = true;
        var magicNumber = reader.ReadUInt16(); // first two bytes should be JPEG magic number (FF D8)
        if (magicNumber != 0xFFD8)
        {
            throw new Exception("Invalid JPEG signature.");
        }
        do
        {
            // Find next segment marker. Markers are zero or more 0xFF bytes, followed
            // by a 0xFF and then a byte not equal to 0x00 or 0xFF.
            byte segmentIdentifier = reader.ReadByte();
            byte segmentType = reader.ReadByte();

            // Read until we have a 0xFF byte followed by a byte that is not 0xFF or 0x00
            while (segmentIdentifier != 0xFF || segmentType == 0xFF || segmentType == 0)
            {
                segmentIdentifier = segmentType;
                segmentType = reader.ReadByte();
            }

            if (segmentType == 0xD9) // EOF?
                return Size.Empty;

            // next 2-bytes are <segment-size>: [high-byte] [low-byte]
            var segmentLength = (int)reader.ReadUInt16();

            // segment length includes size bytes, so subtract two
            segmentLength -= 2;

            if (segmentType == 0xC0 || segmentType == 0xC2)
            {
                reader.ReadByte(); // bits/sample, usually 8
                int height = (int) reader.ReadUInt16();
                int width = (int) reader.ReadUInt16();
                return new Size(width, height);
            }
            else
            {
                // skip this segment
                reader.Skip(segmentLength);
            }
        }
        while (true);
    }

    private static Size DecodePng(SequentialBinaryReader reader)
    {
        reader.IsBigEndian = true;
        reader.ReadBytes(pngSignatureBytes.Length);
        reader.Skip(8);

        int width = reader.ReadInt32();
        int height = reader.ReadInt32();
        return new Size(width, height);
    }

    private static Size DecodeIco(SequentialBinaryReader reader)
    {
        reader.IsBigEndian = false;
        if (reader.ReadUInt32() != 0x0001_0000)
        {
            throw new Exception("Invalid ICO signature");
        };
        // Note: the file may still not be an icon, the signature is also used by other formats such as JBIG.
        int imageCount = reader.ReadUInt16();
        int maxWidth = 0;
        int maxHeight = 0;
        for (int i = 0; i < imageCount; i++)
        {
            int width = reader.ReadByte();
            int height = reader.ReadByte();

            if (width == 0)
                width = 256;

            if (height == 0)
                height = 256;

            if (width > maxWidth && height > maxHeight)
            {
                maxWidth = width;
                maxHeight = height;
            }
            if (maxWidth == 256 && maxHeight == 256)
            {
                return new Size(256, 256);
            }
            reader.ReadBytes(14); // Skip the next 14 bytes for each image
        }
        // TODO: detect sizes larger than 256 px by reading PNG images
        return new Size(maxWidth, maxHeight);
    }

    // Credits: https://stackoverflow.com/questions/111345/getting-image-dimensions-without-reading-the-entire-file,
    // https://metadataconsulting.blogspot.com/2020/09/CSharp-dotNET-How-to-get-image-dimensions-from-header-of-webP-for-all-formats-lossy-lossless-extended-partially-load-image.html?m=1, 
    // https://developers.google.com/speed/webp/docs/riff_container
    private static Size DecodeWebP(SequentialBinaryReader reader)
    {
        reader.IsBigEndian = false;

        string header = new string(reader.ReadChars(12));
        if ((!header.StartsWith("RIFF")) || !header.EndsWith("WEBP"))
        {
            throw new Exception("Invalid WebP signature.");
        }       

        string format = new string(reader.ReadChars(4)); // VP8, VP9L or VP8X
        if (format == "VP8 ")
        {
            reader.ReadBytes(10); // go to width and height
            var width = reader.ReadUInt16() & 0b0011_1111_1111_1111; // 14 bits width
            var height = reader.ReadUInt16() & 0b0011_1111_1111_1111; // 14 bits height
            return new Size(width, height);
        }
        else if (format == "VP8L")
        {
            reader.ReadUInt32(); // size
            byte signature = reader.ReadByte(); // 0x2f signature
            if (signature != 0x2f)
            {
                throw new InvalidDataException("Invalid VP8L signature");
            }
            byte[] wh = reader.ReadBytes(4); //width and height in 1 read
            var width = 1 + (((wh[1] & 0x3F) << 8) | wh[0]); // 14 bits width - https://blog.tcl.tk/38137  
            var height = 1 + (((wh[3] & 0xF) << 10) | (wh[2] << 2) | ((wh[1] & 0xC0) >> 6)); // 14 bits height >> 6))}]
            return new Size(width, height);
        }
        else if (format == "VP8X")
        {
            reader.ReadBytes(8); // skip flags and optional fields 

            byte[] w = reader.ReadBytes(3); //24 bits for width
            var width = 1 + (w[2] << 16 | w[1] << 8 | w[0]); //little endian

            byte[] h = reader.ReadBytes(3); //24 bits for height
            var height = 1 + (h[2] << 16 | h[1] << 8 | h[0]);

            return new Size(width, height);
        }
        else
        {
            throw new InvalidDataException("Unsupported WebP format");
        }
    }

    private static Size DecodeTiff(SequentialBinaryReader reader)
    {
        // Currently returns dimensions of the first image only; this could be improved for multi-frame tiff by returning minimum or maximum dimensions.
        var signature = reader.ReadChars(2); // 42-42 for Little Endian, 4D-4D for Big Endian
        reader.IsBigEndian = new string(signature) == "MM";

        int width = 0;
        int height = 0;
        var version = reader.ReadUInt16();
        switch (version)
        {
            case 0x2A: // TIFF
            case 0x2B: // BigTIFF
                var ifdOffset = reader.ReadUInt32();
                reader.BaseStream.Position = ifdOffset;
                var ifdTagCount = reader.ReadUInt16();
                for (int i = 0; i < ifdTagCount; i++)
                {
                    var tagPosition = reader.BaseStream.Position;
                    var tagId = reader.ReadUInt16();
                    switch (tagId)
                    {
                        case 0x100: // width
                            var dataType = reader.ReadUInt16();
                            reader.ReadUInt32(); // Data count (always 1 for this tag)
                            width = (int)(dataType == 3 ? reader.ReadUInt16() : reader.ReadUInt32());
                            break;
                        case 0x101: // height
                            var dataType2 = reader.ReadUInt16();
                            reader.ReadUInt32(); // Data count (always 1 for this tag)
                            height = (int)(dataType2 == 3 ? reader.ReadUInt16() : reader.ReadUInt32());
                            break;
                    }
                    if (width > 0 && height > 0)
                    {
                        // Both width and height found, exit loop
                        break;
                    }
                    else
                    {
                        // Go to next tag
                        reader.Skip(12 - (int)(reader.BaseStream.Position - tagPosition));
                    }
                }
                break;
            default:
                throw new InvalidDataException("Invalid TIFF signature.");
        }
        return new Size(width, height);
    }

    private static Size DecodeJxr(SequentialBinaryReader reader)
    {
        reader.IsBigEndian = false;
        var signature = reader.ReadUInt32();
        if (signature != 0x01BC4949 && signature != 0x00BC4949)
        {
            throw new Exception("Invalid JPEG-XR / HD Photo signature.");
        }
        var ifdOffset = reader.ReadUInt32();
        reader.BaseStream.Position = ifdOffset;
        var ifdTagCount = reader.ReadUInt16();
        int width = 0;
        int height = 0;
        for (int i = 0; i < ifdTagCount; i++)
        {
            var tagPosition = reader.BaseStream.Position;
            var tagId = reader.ReadUInt16();
            switch (tagId)
            {
                case 0xBC80: // width
                    var dataType = reader.ReadUInt16();
                    reader.ReadUInt32(); // Data count (always 1 for this tag)
                    width = (int)(dataType == 3 ? reader.ReadUInt16() : reader.ReadUInt32());
                    break;
                case 0xBC81: // height
                    var dataType2 = reader.ReadUInt16();
                    reader.ReadUInt32(); // Data count (always 1 for this tag)
                    height = (int)(dataType2 == 3 ? reader.ReadUInt16() : reader.ReadUInt32());
                    break;
            }
            if (width > 0 && height > 0)
            {
                // Both width and height found, exit loop
                break;
            }
            else
            {
                // Go to next tag
                reader.Skip(12 - (int)(reader.BaseStream.Position - tagPosition));
            }
        }
        return new Size(width, height);
    }

    //private static Size DecodeJxl(SequentialBinaryReader reader)
    //{
    //    reader.IsBigEndian = true;
    //}

    private static Size DecodeJpeg2000(SequentialBinaryReader reader)
    {
        reader.IsBigEndian = true;
        var signature = reader.ReadUInt32();
        if (signature == 0xFF4FFF51)
        {
            // Read width and height directly (see below)
        }
        else if (signature == 0x0000000C)
        {
            // Container format
            bool isFirstMarker = true;
            while (true)
            {
                var b = reader.ReadByte();
                if (b == 0xFF)
                {
                    isFirstMarker = false;
                    byte b2 = reader.ReadByte();
                    byte b3 = reader.ReadByte();
                    byte segmentType = reader.ReadByte();

                    if (b2 == 0xD9)
                    {
                        // End of file, exit
                        return Size.Empty;
                    }
                    else if (b2 == 0x00 || b2 == 0xFF)
                    {
                        // not a marker
                        continue;
                    }
                    else if (b2 == 0x4f && b3 == 0xFF && segmentType == 0x51)
                    {
                        // SIZ marker found
                        break;
                    }
                    else
                    {
                        // Go to next segment
                        var segmentLength = (int)reader.ReadUInt16();
                        segmentLength -= 2; // Segment length includes size bytes, so subtract two
                        reader.Skip(segmentLength);
                        break;
                    }
                }
                else if (b != 0xFF && !isFirstMarker)
                {
                    throw new InvalidDataException("Invalid JPEG 2000 data.");
                }
            }
        }
        else
        {
            throw new Exception("Invalid JPEG 2000 signature.");
        }

        reader.Skip(4); // skip to width and height
        var width = reader.ReadUInt32();
        var height = reader.ReadUInt32();
        return new Size((int)width, (int)height);
    }

    private static Size DecodeAvifOrHeif(SequentialBinaryReader reader)
    {
        // To be improved; works for AVIF but HEIC images often have more than one 'ispe' blocks
        reader.IsBigEndian = true;
        var ispe = Encoding.ASCII.GetBytes("ispe");
        int foundChars = 0;
        while (true)
        {
            var c = reader.ReadChar();
            if (c == 'i' && foundChars == 0)
            {
                ++foundChars;
            }
            else if (c == 's' && foundChars == 1)
            {
                ++foundChars;
            }
            else if (c == 'p' && foundChars == 2)
            {
                ++foundChars;
            }
            else if (c == 'e' && foundChars == 3)
            {
                // 'ispe' type found
                break;
            }
            else
            {
                // Not found
                foundChars = 0;
            }
        }
        reader.Skip(4); // version + flags
        var width = reader.ReadUInt32();
        var height = reader.ReadUInt32();
        return new Size((int)width, (int)height);
    }

    private static Size DecodeWmf(SequentialBinaryReader reader)
    {
        reader.IsBigEndian = false;
        if (reader.ReadUInt32() != 0x9AC6CDD7)
        {
            // To be improved; how to get dimensions for metafiles directly starting
            // with the WMF header 01 00 or 02 00 ?
            throw new Exception("Invalid or unsupported WMF file.");
        }
        reader.Skip(2); // HWmf
        int left = reader.ReadInt16();
        int top = reader.ReadInt16();
        int right = reader.ReadInt16();
        int bottom = reader.ReadInt16();
        int scale = reader.ReadUInt16(); // number of metafile units per inch

        // Note: WMF size is returned in inches
        int widthInInches = (int)Math.Round((right - left) / (decimal)scale);
        int heightInInches = (int)Math.Round((bottom - top) / (decimal)scale);

        return new Size(widthInInches, heightInInches);
    }

    private static Size DecodeEmf(SequentialBinaryReader reader)
    {
        reader.IsBigEndian = false;

        // EMR_HEADER: Type + Size + Bounds
        reader.Skip(4 + 4 + 16);

        // Frame: specify the rectangular inclusive-inclusive dimensions, 
        // in .01 millimeter units, of a rectangle that surrounds the image stored in the metafile.
        int left   = reader.ReadInt32();
        int top    = reader.ReadInt32();
        int right  = reader.ReadInt32();
        int bottom = reader.ReadInt32();

        // Skip other headers:
        // Signature (4) + Version (4) + Size (4) + Records (4) + Handles (2)
        // + nReserved (2) + nDescription (4) + offDescription (4) + PalEntries (4)
        reader.Skip(32);

        // Next 8 bytes specify the size of the reference device, in pixels
        int deviceSizeInPixelX = reader.ReadInt32();
        int deviceSizeInPixelY = reader.ReadInt32();

        // Then 8 more to specify the size of the reference device, in millimeters
        int deviceSizeInMlmX = reader.ReadInt32();
        int deviceSizeInMlmY = reader.ReadInt32();

        int widthInPixel = (int) Math.Round(0.5 + (right - left + 1.0) * deviceSizeInPixelX / deviceSizeInMlmX / 100.0);
        int heightInPixel = (int) Math.Round(0.5 + (bottom - top + 1.0) * deviceSizeInPixelY / deviceSizeInMlmY / 100.0);

        return new Size(widthInPixel, heightInPixel);
    }

    private static Size DecodeXml(Stream stream)
    {
        try
        {
            var nav = new XPathDocument(stream).CreateNavigator();
            // use local-name() to ignore any xml namespace
            nav = nav.SelectSingleNode("/*[local-name() = 'svg']");
            if (nav is not null)
            {
                var width = Unit.Parse(nav.GetAttribute("width", string.Empty), UnitMetric.Pixel);
                var height = Unit.Parse(nav.GetAttribute("height", string.Empty), UnitMetric.Pixel);
                if (width.IsValid && width.Type.IsValid() && height.IsValid && height.Type.IsValid())
                    return new Size(width.ValueInPx, height.ValueInPx);

                // If width or height are not found or use unsupported units (%, auto, unitless),
                // try to get the viewBox
                var viewBox = nav.GetAttribute("viewBox", string.Empty);
                if (!string.IsNullOrWhiteSpace(viewBox))
                {
                    var rectParts = viewBox.Split([' ', ',', ';'], StringSplitOptions.RemoveEmptyEntries);
                    if (rectParts.Length == 4)
                    {
                        width = Unit.Parse(rectParts[2], UnitMetric.Pixel);
                        height = Unit.Parse(rectParts[3], UnitMetric.Pixel);
                        return new Size(width.ValueInPx, height.ValueInPx);
                    }
                }

                // If viewBox is not found assume unsupported units as pixels 
                // (at least aspect ratio is preserved, better than returning 0).
                if (width.IsValid && height.IsValid)
                    return new Size(width.ValueInPx, height.ValueInPx);
            }
        }
        catch (SystemException)
        {
            return Size.Empty;
        }
        return Size.Empty;
    }
}
