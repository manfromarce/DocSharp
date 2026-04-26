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

public interface ISaveOptions
{
    SaveFormat Format { get; }
}

public class DocxSaveOptions : ISaveOptions
{
    public SaveFormat Format => SaveFormat.Docx;
    public WordprocessingDocumentType DocumentType { get; set; } = WordprocessingDocumentType.Document;
}

public class RtfSaveOptions : ISaveOptions
{
    public SaveFormat Format => SaveFormat.Rtf;

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

    /// <summary>
    /// Specifies the default code page number, that will be written in the RTF header and used to encode special characters. 
    /// For example, "à" is encoded as \'e0 when the encoding is Windows-1252. 
    /// For characters not available in the code page, Unicode will be used, for example (e.g. \uc1\u915). 
    /// By default, the code page of the system region is used and changing it is not recommended. 
    /// A full list of code pages can be found at https://learn.microsoft.com/en-us/windows/win32/intl/code-page-identifiers, 
    /// but note that not all code pages are supported in RTF (notably, UTF-8 is not supported). 
    /// The RTF specification can be found in the repository. 
    /// </summary>
    public int DefaultCodePage { get; set; } = Encodings.SystemCodePage;
}

public class HtmlSaveOptions : ISaveOptions
{
    public SaveFormat Format => SaveFormat.Html;

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
    /// Get or set the base file path for processing external sub-documents (if any).
    /// If null or empty, sub-documents will not be preserved.
    /// </summary>
    public string? OriginalFolderPath { get; set; }

    public HeaderFooterExportOptions HeaderFooterExportOptions { get; set; } = HeaderFooterExportOptions.FirstHeaderLastFooter;

    public FootnoteEndnoteExportOptions FootnoteEndnoteExportOptions { get; set; } = FootnoteEndnoteExportOptions.EndOfDocument;

    /// <summary>
    /// Used to map DOCX styles by name. The default <see cref="DefaultStyleNamingResolver"/> can be overriden to customize style mappings.
    /// </summary>
    public IStyleNamingResolver StyleNamingResolver { get; set; } = new DefaultStyleNamingResolver();

    /// <summary>
    /// Get or set whether an horizontal rule (---) should be written between different sections.
    /// </summary>
    public bool HorizontalRuleForSectionBreaks { get; set; } = false;

    /// <summary>
    /// Get or set whether an horizontal rule (---) should be written after forced page breaks.
    /// </summary>
    public bool HorizontalRuleForPageBreaks { get; set; } = false;

    /// <summary>
    /// If true, the converter will produce a fixed-layout page-like container for sections
    /// (applies width, padding/margins and borders to the section `div` and centers it).
    /// Default is false (fluid layout).
    /// </summary>
    public bool FixedLayout { get; set; } = false;

    /// <summary>
    /// By default only inline images are supported,  
    /// because other DOCX image layouts have no direct equivalent in HTML/Markdown and can lead to unexpected results.  
    /// However, if desired, this property can be set to ImageLayoutType.InlineAndAnchored 
    /// to preserve the "top and bottom", "square", "tight" and "through" wrap layouts too, 
    /// or to ImageLayoutType.All to preserve absolutely positioned images ("in front of"/"behind" text) too.
    /// </summary>
    public ImageLayoutType SupportedImagesLayout { get; set; } = ImageLayoutType.Inline;
}

public class MarkdownSaveOptions : ISaveOptions
{
    public SaveFormat Format => SaveFormat.Markdown;

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

    /// <summary>
    /// Get or set whether top/bottom/between paragraph borders in DOCX should produce an horizontal rule (---) in Markdown.
    /// </summary>
    public bool HorizontalRuleForTopBottomBorders { get; set; } = false;

    /// <summary>
    /// Get or set whether an horizontal rule (---) should be written between different sections.
    /// </summary>
    public bool HorizontalRuleForSectionBreaks { get; set; } = true;

    /// <summary>
    /// Get or set whether an horizontal rule (---) should be written after forced page breaks.
    /// </summary>
    public bool HorizontalRuleForPageBreaks { get; set; } = false;

    /// <summary>
    /// This property can be used to optionally set font families (e.g. Courier New, Cascadia Code) 
    /// that should be mapped to an inline code element in Markdown. 
    /// </summary>
    public string[]? CodeFontFamilies { get; set; } = null;

    /// <summary>
    /// By default only inline images are supported,  
    /// because other DOCX image layouts have no direct equivalent in HTML/Markdown and can lead to unexpected results.  
    /// However, if desired, this property can be set to ImageLayoutType.InlineAndAnchored 
    /// to preserve the "top and bottom", "square", "tight" and "through" wrap layouts too, 
    /// or to ImageLayoutType.All to preserve absolutely positioned images ("in front of"/"behind" text) too.
    /// </summary>
    public ImageLayoutType SupportedImagesLayout { get; set; } = ImageLayoutType.Inline;

    /// <summary>
    /// Used to map DOCX styles by name. The default <see cref="DefaultStyleNamingResolver"/> can be overriden to customize style mappings.
    /// </summary>
    public IStyleNamingResolver StyleNamingResolver { get; set; } = new DefaultStyleNamingResolver();
}

public class TxtSaveOptions : ISaveOptions
{
    public SaveFormat Format => SaveFormat.Txt;

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

    /// <summary>
    /// Sets whether the image description or alternate text (if available) 
    /// should be written in the output plain text instead of the image (default is false). 
    /// </summary>
    public bool WriteImageDescription { get; set; } = false;

    /// <summary>
    /// Get or set whether special horizontal line shapes in DOCX should produce an horizontal line (---) in plain text.
    /// </summary>
    public bool HorizontalRuleForHorizontalLineShapes { get; set; } = true;

    /// <summary>
    /// Get or set whether top/bottom/between paragraph borders in DOCX should produce an horizontal line (---) in plain text.
    /// </summary>
    public bool HorizontalRuleForTopBottomBorders { get; set; } = false;

    /// <summary>
    /// Get or set whether an horizontal line (---) should be written between different sections.
    /// </summary>
    public bool HorizontalRuleForSectionBreaks { get; set; } = false;

    /// <summary>
    /// Get or set whether an horizontal line (---) should be written after forced page breaks.
    /// </summary>
    public bool HorizontalRuleForPageBreaks { get; set; } = false;
}
