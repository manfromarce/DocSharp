using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace DocSharp.Docx;

public interface ISaveOptions { }

public class DocxSaveOptions : ISaveOptions
{
    public WordprocessingDocumentType DocumentType { get; set; } = WordprocessingDocumentType.Document;
}

public class RtfSaveOptions : ISaveOptions
{
    /// <summary>
    /// Gets or set the default font and paragraph properties used in (rare) cases where 
    /// they are not specified in in neither the document body, styles or default style. 
    /// In these cases, different word processors and versions behave differently. 
    /// If not set, DocSharp will emulate recent Microsoft Word versions. 
    /// </summary>
    public DocumentDefaultSettings DefaultSettings { get; set; } = new DocumentDefaultSettings();
    
    /// <summary>
    /// Get or set the output folder path, which will be used for saving sub-documents (if any) and calculating their relative file paths.
    /// If null or empty, sub-documents will not be preserved.
    /// </summary>
    public string? OutputFolderPath { get; set; }

    /// <summary>
    /// Image converter to preserve TIFF, GIF and other image types when converting to RTF. 
    /// If the DocSharp.ImageSharp or DocSharp.SystemDrawing package is installed, 
    /// this property can be set to a new instance of ImageSharpConverter or SystemDrawingConverter. 
    /// </summary>
    public IImageConverter? ImageConverter { get; set; } = null;

    /// <summary>
    /// Get or set the base file path for processing external sub-documents (if any).
    /// If null or empty, sub-documents will not be preserved.
    /// </summary>
    public string? OriginalFolderPath { get; set; }
}

public class HtmlSaveOptions : ISaveOptions
{
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

    /// <summary>
    /// Since HTML is not paginated, only the header of the first section and
    /// footer of the last section are exported.
    /// Set this property to false to ignore headers and footers.
    /// </summary>
    public bool ExportHeaderFooter { get; set; } = true;

    /// <summary>
    /// Since HTML is not paginated, both footnotes and endnotes are exported at the end of the document.
    /// Set this property to false to ignore footnotes and endnotes.
    /// </summary>
    public bool ExportFootnotesEndnotes { get; set; } = true;

    /// <summary>
    /// Get or set the base file path for processing external sub-documents (if any).
    /// If null or empty, sub-documents will not be preserved.
    /// </summary>
    public string? OriginalFolderPath { get; set; }
}

public class MarkdownSaveOptions : ISaveOptions
{
    /// <summary>
    /// If this property is set to a directory, images will be exported to that folder
    /// and a reference will be added in Markdown syntax,
    /// otherwise images are not converted. 
    /// If the directory does not exist, it will be created.
    /// NOTE: if the directory contains image files with the same names as in the DOCX document archive 
    /// (usually image1.*, image2.*, ...), they will be overwritten.
    /// </summary>
    public string? ImagesOutputFolder { get; set; } = string.Empty;

    /// <summary>
    /// This property is used in combination with ImagesOutputFolder to determine 
    /// how the image files are specified in Markdown.
    /// 
    /// If this property is set to null, an absolute path such as "file:///c:/.../image.jpg" 
    /// will be created using the ImagesOutputFolder value and the image file name.
    /// 
    /// Otherwise, the base path (exluding the image file name) is replaced by this value.
    /// Possible values:
    /// - empty string or "." : images are expected to be in the same folder as the Markdown file.
    /// - relative paths such as "images" or "../images": images are expected to be in a subfolder or parent folder.
    /// - "/server/user/files/" or "C:\images": replaces the file path entirely
    /// (the image file name is still appended and Windows paths are converted to the file URI scheme).
    /// 
    /// This property does not affect where the images are actually saved, and can be useful if
    /// the Markdown document is not saved to file, or in environments with limited file system access.
    /// </summary>
    public string? ImagesBaseUriOverride { get; set; } = null;

    /// <summary>
    /// Image converter to preserve TIFF, EMF and other image types when converting to Markdown. 
    /// If the DocSharp.ImageSharp or DocSharp.SystemDrawing package is installed, 
    /// this property can be set to a new instance of ImageSharpConverter or SystemDrawingConverter. 
    /// </summary>
    public IImageConverter? ImageConverter { get; set; } = null;

    /// <summary>
    /// Since Markdown is not paginated, only the header of the first section and
    /// footer of the last section are exported.
    /// Set this property to false to ignore headers and footers.
    /// </summary>
    public bool ExportHeaderFooter { get; set; } = true;

    /// <summary>
    /// Since Markdown is not paginated, both footnotes and endnotes are exported at the end of the document.
    /// Set this property to false to ignore footnotes and endnotes.
    /// </summary>
    public bool ExportFootnotesEndnotes { get; set; } = true;

    /// <summary>
    /// Get or set the base file path for processing external sub-documents (if any).
    /// If null or empty, sub-documents will not be preserved.
    /// </summary>
    public string? OriginalFolderPath { get; set; }
}

public class TxtSaveOptions : ISaveOptions
{
    /// <summary>
    /// Since plain text is not paginated, only the header of the first section and
    /// footer of the last section are exported.
    /// Set this property to false to ignore headers and footers.
    /// </summary>
    public bool ExportHeaderFooter { get; set; } = true;

    /// <summary>
    /// Since plain text is not paginated, both footnotes and endnotes are exported at the end of the document.
    /// Set this property to false to ignore footnotes and endnotes.
    /// </summary>
    public bool ExportFootnotesEndnotes { get; set; } = true;

    /// <summary>
    /// Get or set the base file path for processing external sub-documents (if any).
    /// If null or empty, sub-documents will not be preserved.
    /// </summary>
    public string? OriginalFolderPath { get; set; }
}