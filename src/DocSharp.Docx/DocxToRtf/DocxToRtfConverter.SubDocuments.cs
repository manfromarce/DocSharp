using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using DocumentFormat.OpenXml.Wordprocessing;
using Shadow14 = DocumentFormat.OpenXml.Office2010.Word.Shadow;
using Outline14 = DocumentFormat.OpenXml.Office2010.Word.TextOutlineEffect;
using DocSharp.Helpers;
using M = DocumentFormat.OpenXml.Math;
using DocSharp.Writers;
using System.Diagnostics;
using DocumentFormat.OpenXml.Packaging;
using System.IO;
using DocSharp.Collections;

namespace DocSharp.Docx;

public partial class DocxToRtfConverter : DocxToStringWriterBase<RtfStringWriter>
{
    internal void WriteFileTable(FastStringCollection files, MainDocumentPart mainPart, RtfStringWriter sb)
    {
        if (!(string.IsNullOrWhiteSpace(OriginalFolderPath) || string.IsNullOrWhiteSpace(OutputFolderPath)))
        {
            sb.Write(@"{\*\filetbl ");
            foreach (var file in files)
            {
                var rel = mainPart.ExternalRelationships.Where(r => r.Id != null && r.Id == file.Key).FirstOrDefault();
                if (rel?.Uri != null)
                {
                    string unescapedPath;
                    string outputFilePath = OutputFolderPath!;
                    try
                    {
                        if (!Directory.Exists(OutputFolderPath))
                        {
                            Directory.CreateDirectory(OutputFolderPath);
                        }

                        string url = rel.Uri.OriginalString;
                        unescapedPath = Uri.UnescapeDataString(url); // Unescapes sequences such as %20
                        unescapedPath = Path.Combine(OriginalFolderPath, unescapedPath);

                        if (File.Exists(unescapedPath)) // Ensure the original subdocument exists
                        {
                            // Build file path for the converted subdocument
                            string outputFileName = Path.GetFileNameWithoutExtension(unescapedPath) + ".rtf";
                            outputFilePath = Path.Combine(OutputFolderPath, outputFileName);
                            using (var secondDoc = WordprocessingDocument.Open(unescapedPath, false))
                            {
                                // Convert the subdocument
                                var secondConverter = new DocxToRtfConverter()
                                {
                                    ImageConverter = this.ImageConverter,
                                    DefaultSettings = this.DefaultSettings
                                };
                                secondConverter.Convert(secondDoc, outputFilePath);

                                if (File.Exists(outputFilePath)) // Ensure the converted subdocument exists
                                {
                                    sb.Write(@"{\file ");
                                    sb.Write(@$"\fid{file.Value} "); // index (referenced by \subdocumentN in the document)

                                    if (!rel.Uri.IsAbsoluteUri)
                                    {
                                        // If the original URI was relative, keep it relative.
                                        int i = outputFilePath.IndexOf(outputFileName);
                                        if (i >= 0)
                                            sb.Write(@$"\frelative{i} ");
                                    }

                                    // TODO
                                    // if (true)
                                    // {
                                    sb.Write(@"\fvalidntfs ");
                                    // sb.Write(@"\fvalidmac ");
                                    // sb.Write(@"\fvaliddos ");
                                    // sb.Write(@"\fvalidhpfs ");
                                    // }

                                    if (rel.Uri.IsAbsoluteUri && rel.Uri.IsUnc)
                                    {
                                        sb.Write(@"\fnetwork ");
                                    }

                                    sb.WriteRtfEscaped(outputFilePath); // Escape chars that are valid for filenames but not valid in RTF
                                    sb.WriteLine('}');
                                }
                            }
                        }
                    }
                    catch (Exception ex)
                    {
#if DEBUG
                        Debug.WriteLine($"Exception in processing subdocument: {ex.Message}");
#endif
                        break;
                    }
                }
            }
            sb.WriteLine("}");
        }
    }

    internal override void ProcessSubDocumentReference(SubDocumentReference subDocReference, RtfStringWriter sb)
    {
        if (!string.IsNullOrWhiteSpace(OriginalFolderPath) &&
            !string.IsNullOrWhiteSpace(OutputFolderPath) &&
            subDocReference.Id?.Value != null)
        {
            // Keep track of the file in the file table
            files.TryAddAndGetIndex(subDocReference.Id.Value, out int fileIndex);
            sb.Write($"\\subdocument{fileIndex}");
        }
    }

    internal override void ProcessAltChunk(AltChunk altChunk, RtfStringWriter writer)
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
                        if (alternativeFormatImportPart.ContentType == AlternativeFormatImportPartType.Rtf.ContentType)
                        {
                            // Read the content and append it to the RTF.
                            using (var sr = new StreamReader(stream))
                            {
                                writer.WriteLine();
                                writer.Write('{');
                                // TODO: skip the RTF header
                                writer.Write(sr.ReadToEnd());
                                writer.WriteLine('}');
                            }
                        }
                        // else if (alternativeFormatImportPart.ContentType == AlternativeFormatImportPartType.Html.ContentType)
                        // {
                        //      using (var sr = new StreamReader(stream))
                        //     {
                        //         ProcessHtml(sr.ReadToEnd(), writer);
                        //     }
                        // }
                        // else if (alternativeFormatImportPart.ContentType == AlternativeFormatImportPartType.Mht.ContentType)
                        // {
                        // }
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

    internal void ProcessHtml(string html, RtfStringWriter writer)
    {
        // TODO

        // https://learn.microsoft.com/en-us/openspecs/exchange_server_protocols/ms-oxrtfex/fc107c62-c62f-48d4-b114-d7e3d1c8f54b
        // https://learn.microsoft.com/en-us/openspecs/exchange_server_protocols/ms-oxrtfex/4f09a809-9910-43f3-a67c-3506b09ca5ac
    }
}
